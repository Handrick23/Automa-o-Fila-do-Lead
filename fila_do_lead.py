from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import customtkinter as ctk
import pandas as pd
import math
import os
import tkinter.filedialog as filedialog
from tkinter import messagebox

# =================================================================
# VARIÁVEIS GLOBAIS: Armazenam os DataFrames carregados via Excel
# =================================================================
df_semana = None      # Armazena os dados da aba "BASE LEAD" (Vendas da semana)
df_consultores = None  # Armazena os dados da aba "CONSULTORES" (Lista mestre e status)
df_mensal = None       # Armazena os dados da aba "BASE MENSAL" (Vendas acumuladas)

def processar_vendas(dataframe, sufixo=""):
    """
    CONSOLIDAÇÃO DE DADOS:
    Transforma uma lista de transações individuais em um resumo por consultor.
    - Agrupa por 'Consultor' e 'Tipo Cliente' (Novo/Existente).
    - Soma os valores de venda.
    - Calcula o volume total.
    """
    if dataframe is None or dataframe.empty:
        return pd.DataFrame(columns=['Consultor', f'Venda Novo{sufixo}', f'Venda Existente{sufixo}', f'Total{sufixo}'])
    
    # Padronização de nomes de colunas (Remove espaços e Capitaliza)
    dataframe.columns = [str(c).strip().title() for c in dataframe.columns]
    
    # Identifica a coluna numérica de valor (busca por 'Venda' ou a primeira disponível)
    col_valor = 'Venda'
    if col_valor not in dataframe.columns:
        cols_num = dataframe.select_dtypes(include=['number']).columns
        col_valor = cols_num[0] if len(cols_num) > 0 else 'Venda'

    # Cria uma tabela dinâmica: Consultor nas linhas e Tipo Cliente nas colunas
    resumo = dataframe.pivot_table(
        index='Consultor', 
        columns='Tipo Cliente', 
        values=col_valor, 
        aggfunc='sum'
    ).fillna(0).reset_index()

    # Garante que as colunas 'Novo' e 'Existente' existam mesmo que não haja vendas no tipo
    for col in ['Novo', 'Existente']:
        if col not in resumo.columns:
            resumo[col] = 0
            
    # Renomeia colunas para facilitar o merge posterior (ex: Venda Novo_Mensal)
    resumo = resumo.rename(columns={'Novo': f'Venda Novo{sufixo}', 'Existente': f'Venda Existente{sufixo}'})
    
    # Cálculo do Total Geral por consultor (Novo + Existente)
    resumo[f'Total{sufixo}'] = resumo[f'Venda Novo{sufixo}'] + resumo[f'Venda Existente{sufixo}']
    return resumo

def gerar_fila_do_lead():
    """
    CORE DA REGRA DE NEGÓCIO: 
    Aplica o ranqueamento em camadas para Fila 1 e Fila 2.
    """
    global df_semana, df_consultores, df_mensal

    if df_semana is None or df_consultores is None:
        messagebox.showwarning("Aviso", "Por favor, carregue o arquivo Excel primeiro.")
        return

    # 1. Processamento Prévio: Transforma abas em resumos de performance
    vendas_semana = processar_vendas(df_semana)
    vendas_mensal = processar_vendas(df_mensal, sufixo="_Mensal")

    # 2. Filtragem de Disponibilidade: Remove quem está com status 'FÉRIAS'
    base_fila = df_consultores[['Consultor', 'Equipe', 'Justificativa']].rename(columns={'Justificativa': 'Status'})
    base_fila = base_fila[base_fila['Status'].astype(str).str.upper() != 'FÉRIAS'].copy()
    
    # 3. Cruzamento de Dados (Merge): Une consultores com suas vendas semanais e mensais
    base_fila = pd.merge(base_fila, vendas_semana, on='Consultor', how='left').fillna(0)
    base_fila = pd.merge(base_fila, vendas_mensal, on='Consultor', how='left').fillna(0)

    # 4. Normalização de Equipes: Mapeia diversos nomes para uma sigla de Filial comum
    def mapear_unidade_comercial(equipe):
        e = str(equipe).upper().strip()
        mapeamento = {
            'GRANDES CONTAS SP': 'SPO', 'SP 1': 'SPO', 'SPO': 'SPO', 'SP': 'SPO',
            'GRANDES CONTAS BH': 'BHZ', 'MG 1': 'BHZ', 'BHZ': 'BHZ', 'BH': 'BHZ',
            'GRANDES CONTAS RJ': 'RJO', 'RJ 1': 'RJO', 'RJO': 'RJO', 'RJ': 'RJO',
            'GRANDES CONTAS CTA': 'CTA', 'CURITIBA': 'CTA', 'CTA': 'CTA',
            'SP INTERIOR': 'SP INTERIOR', 'SPI': 'SP INTERIOR'
        }
        return mapeamento.get(e, e)

    base_fila['Filial_Final'] = base_fila['Equipe'].apply(mapear_unidade_comercial)

    # 5. Setup do Excel Final (Estilização de bordas, cores e fontes)
    wb = Workbook()
    ws = wb.active
    ws.title = "Fila do Lead"
    
    side_m = Side(style='medium', color='000000')
    border_top = Border(left=side_m, right=side_m, top=side_m)
    border_mid = Border(left=side_m, right=side_m)
    border_bot = Border(left=side_m, right=side_m, bottom=side_m)
    fill_header = PatternFill(start_color="384B59", end_color="384B59", fill_type="solid")
    font_header = Font(bold=True, color="FFFFFF", size=11)
    font_red = Font(bold=True, color="FF0000", size=11)
    align_center = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions['A'].width = 45
    row_idx = 1

    # 6. Processamento por Filial (Geração das Filas)
    for filial in sorted(base_fila['Filial_Final'].unique()):
        grupo = base_fila[base_fila['Filial_Final'] == filial].copy()

        # --- REGRA 1: FILA 1 (Prioridade Venda Novo na Semana) ---
        # Ordenação: Quem vendeu mais para clientes NOVOS na semana fica no topo
        cat_a = grupo[grupo['Total'] > 0].sort_values(by=['Venda Novo', 'Total'], ascending=False)
        
        # Divisão F1: A Fila 1 recebe 50% dos vendedores que performaram na semana
        num_vendedores_semana = len(cat_a)
        # Cálculo: $n\_f1 = \lceil \text{num\_vendedores\_semana} / 2 \rceil$
        n_f1 = math.ceil(num_vendedores_semana / 2) if num_vendedores_semana > 0 else 0
        lista_f1 = cat_a.head(n_f1)
        
        # --- REGRA 2: FILA 2 - CAMADA 1 (Vendedores da semana que sobraram da F1) ---
        # Ordenação: Re-ordenados pelo TOTAL geral da semana (Novo + Existente)
        sobras_cat_a = cat_a.tail(num_vendedores_semana - n_f1).sort_values(by='Total', ascending=False)
        
        # --- REGRA 3: FILA 2 - CAMADA 2 (Quem não vendeu na semana, mas vendeu no mês) ---
        # Ordenação: Baseada puramente no Total Mensal acumulado
        cat_b = grupo[(grupo['Total'] == 0) & (grupo['Total_Mensal'] > 0)].sort_values(by='Total_Mensal', ascending=False)
        
        # CAMADA 3: Consultores sem nenhuma venda (Fila de entrada ou baixa performance)
        # Ordenação: Aleatória (Sorteio) para garantir rotatividade
        cat_c = grupo[(grupo['Total'] == 0) & (grupo['Total_Mensal'] == 0)].sample(frac=1)

        # Concatenação da Fila 2 respeitando a hierarquia das 3 camadas
        lista_f2 = pd.concat([sobras_cat_a, cat_b, cat_c])

        # --- ESCRITA DOS BLOCOS NO EXCEL ---
        # Título da Filial
        cell = ws.cell(row=row_idx, column=1, value=f"{filial} Comercial")
        cell.fill, cell.font, cell.alignment, cell.border = fill_header, font_header, align_center, border_top
        row_idx += 1

        # Bloco Fila 1
        ws.cell(row=row_idx, column=1, value="Fila 1").font = font_red
        ws.cell(row=row_idx, column=1).alignment = align_center
        ws.cell(row=row_idx, column=1).border = border_mid
        row_idx += 1
        for nome in lista_f1['Consultor']:
            c = ws.cell(row=row_idx, column=1, value=nome.title())
            c.alignment, c.border = align_center, border_mid
            row_idx += 1

        # Espaçadores estéticos entre filas
        for _ in range(2): 
            ws.cell(row=row_idx, column=1).border = border_mid
            row_idx += 1

        # Bloco Fila 2
        ws.cell(row=row_idx, column=1, value="Fila 2").font = font_red
        ws.cell(row=row_idx, column=1).alignment = align_center
        ws.cell(row=row_idx, column=1).border = border_mid
        row_idx += 1
        total_f2 = len(lista_f2)
        for i, nome in enumerate(lista_f2['Consultor']):
            c = ws.cell(row=row_idx, column=1, value=nome.title())
            c.alignment = align_center
            # Aplica borda inferior apenas no último nome do bloco
            c.border = border_bot if i == total_f2 - 1 else border_mid
            row_idx += 1
        row_idx += 2 

    # Salva e abre o arquivo automaticamente
    nome_final = "Fila_do_Lead.xlsx"
    wb.save(nome_final)
    messagebox.showinfo("Sucesso", "Fila gerada com as novas regras!")
    os.startfile(nome_final)

def importar_planilha():
    """
    UI LOGIC: Abre o seletor de arquivos e carrega as abas específicas.
    """
    global df_semana, df_consultores, df_mensal
    caminho = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    if caminho:
        try:
            xls = pd.ExcelFile(caminho)
            # Cria um mapa de abas em maiúsculo para busca flexível de nomes
            abas_map = {a.strip().upper(): a for a in xls.sheet_names}
            
            # Identifica as abas obrigatórias
            aba_sem = abas_map.get('BASE LEAD') or abas_map.get('BASE SEMANAL')
            aba_men = abas_map.get('BASE MENSAL')
            aba_con = abas_map.get('CONSULTORES')

            if aba_sem and aba_men and aba_con:
                df_semana = pd.read_excel(xls, aba_sem)
                df_mensal = pd.read_excel(xls, aba_men)
                df_consultores = pd.read_excel(xls, aba_con)
                
                # Normaliza a coluna 'Consultor' para garantir que o Merge (VLOOKUP) funcione
                for d in [df_semana, df_mensal, df_consultores]:
                    d.columns = [str(c).strip().title() for c in d.columns]
                    if 'Consultor' in d.columns:
                        d['Consultor'] = d['Consultor'].astype(str).str.strip().str.upper()
                messagebox.showinfo("Sucesso", "Bases carregadas com sucesso!")
            else:
                messagebox.showerror("Erro", "Abas obrigatórias não encontradas!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar arquivo: {e}")

# =================================================================
# INTERFACE GRÁFICA (CustomTkinter)
# =================================================================
ctk.set_appearance_mode("dark")
app = ctk.CTk()
app.title("Fila Lead Pro 2026")
app.geometry('400x280')

ctk.CTkLabel(app, text="Fila do Lead", font=("Arial", 24, "bold")).pack(pady=30)

f_main = ctk.CTkFrame(app)
f_main.pack(pady=10, padx=30, fill="both", expand=True)

ctk.CTkButton(f_main, text="Carregar Planilha Excel", command=importar_planilha, width=250).pack(pady=20)
ctk.CTkButton(f_main, text="Gerar Fila Ranqueada", command=gerar_fila_do_lead, fg_color="#27ae60", width=250).pack(pady=10)

app.mainloop()
