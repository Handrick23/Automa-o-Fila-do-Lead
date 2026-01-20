from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import customtkinter as ctk
import pandas as pd
import math
import os
import tkinter.filedialog as filedialog
from tkinter import messagebox

# =================================================================
# VARIÁVEIS GLOBAIS: Armazenam os dados na memória após o upload
# =================================================================
df_semana = None      # Dados de vendas da semana atual
df_consultores = None  # Lista mestre de consultores e seus status
df_mensal = None      # Dados acumulados do mês para critério de desempate

def processar_vendas(dataframe, sufixo=""):
    """
    REGRA DE NEGÓCIO: Transforma uma lista bruta de vendas em um resumo por consultor.
    - Consolida vendas por 'Tipo Cliente' (Novo ou Existente).
    - Trata erros de nomenclatura (espaços extras e maiúsculas).
    - Garante que colunas de valor existam mesmo que não haja vendas.
    """
    if dataframe is None or dataframe.empty:
        return pd.DataFrame(columns=['Consultor', f'Venda Novo{sufixo}', f'Venda Existente{sufixo}', f'Total{sufixo}'])
    
    # Padronização: Remove espaços e coloca em 'Title Case' para evitar erros de busca
    dataframe.columns = [str(c).strip().title() for c in dataframe.columns]
    
    # Identifica a coluna de valor dinamicamente (busca 'Venda' ou a primeira numérica)
    col_valor = 'Venda'
    if col_valor not in dataframe.columns:
        cols_num = dataframe.select_dtypes(include=['number']).columns
        if len(cols_num) > 0:
            col_valor = cols_num[0]
        else:
            return pd.DataFrame(columns=['Consultor', f'Venda Novo{sufixo}', f'Venda Existente{sufixo}', f'Total{sufixo}'])

    # Pivotagem: Soma os valores cruzando Consultor x Tipo de Cliente
    resumo = dataframe.pivot_table(
        index='Consultor', 
        columns='Tipo Cliente', 
        values=col_valor, 
        aggfunc='sum'
    ).fillna(0).reset_index()

    # Garante a existência das colunas obrigatórias para o cálculo da fila
    for col in ['Novo', 'Existente']:
        if col not in resumo.columns:
            resumo[col] = 0
            
    resumo = resumo.rename(columns={
        'Novo': f'Venda Novo{sufixo}', 
        'Existente': f'Venda Existente{sufixo}'
    })
    
    # Cálculo do volume total (Soma de novos + existentes)
    resumo[f'Total{sufixo}'] = resumo[f'Venda Novo{sufixo}'] + resumo[f'Venda Existente{sufixo}']
    return resumo

def gerar_fila_do_lead():
    """
    CORE DO PROGRAMA: Aplica a lógica de ranqueamento e gera o Excel final.
    """
    global df_semana, df_consultores, df_mensal

    if df_semana is None or df_consultores is None:
        messagebox.showwarning("Aviso", "Por favor, carregue o arquivo Excel primeiro.")
        return

    # 1. Processamento Prévio
    vendas_semana = processar_vendas(df_semana)
    vendas_mensal = processar_vendas(df_mensal, sufixo="_Mensal")

    # 2. Filtragem de Status: Remove quem está de Férias (Regra de disponibilidade)
    base_fila = df_consultores[['Consultor', 'Equipe', 'Justificativa']].rename(columns={'Justificativa': 'Status'})
    base_fila = base_fila[base_fila['Status'].astype(str).str.upper() != 'FÉRIAS'].copy()
    
    # 3. Cruzamento de Dados (Join): Une os cadastros com as vendas semanais e mensais
    base_fila = pd.merge(base_fila, vendas_semana, on='Consultor', how='left').fillna(0)
    base_fila = pd.merge(base_fila, vendas_mensal, on='Consultor', how='left').fillna(0)

    # 4. Padronização de Filiais: Agrupa diferentes nomes de equipes em siglas regionais
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

    # 5. Configuração Estética do Excel (Bordas e Cores)
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

    # 6. Aplicação do Ranqueamento por Filial
    for filial in sorted(base_fila['Filial_Final'].unique()):
        grupo = base_fila[base_fila['Filial_Final'] == filial].copy()

        # REGRA DE NEGÓCIO - CATEGORIZAÇÃO:
        # Cat A: Vendeu na semana (Prioridade máxima)
        # Cat B: Não vendeu na semana, mas vendeu no mês (Recuperação)
        # Cat C: Não vendeu nada (Fila de entrada/baixa performance)
        cat_a = grupo[grupo['Total'] > 0].sort_values(by=['Venda Novo', 'Total'], ascending=False)
        cat_b = grupo[(grupo['Total'] == 0) & (grupo['Total_Mensal'] > 0)].sort_values(by=['Venda Novo_Mensal', 'Total_Mensal'], ascending=False)
        cat_c = grupo[(grupo['Total'] == 0) & (grupo['Total_Mensal'] == 0)].sample(frac=1) # Aleatório para Cat C

        # DIVISÃO DA FILA 1 E FILA 2:
        # Metade superior da Cat A vai para Fila 1. O restante compõe a Fila 2.
        num_vendedores_semana = len(cat_a)
        n_f1 = math.ceil(num_vendedores_semana / 2) if num_vendedores_semana > 0 else 0
        
        lista_f1 = cat_a.head(n_f1)
        lista_f2 = pd.concat([cat_a.tail(num_vendedores_semana - n_f1), cat_b, cat_c])

        # --- ESCREVER NO EXCEL (Bloco Visual) ---
        cell = ws.cell(row=row_idx, column=1, value=f"{filial} Comercial")
        cell.fill, cell.font, cell.alignment, cell.border = fill_header, font_header, align_center, border_top
        row_idx += 1

        f1_label = ws.cell(row=row_idx, column=1, value="Fila 1")
        f1_label.font, f1_label.alignment, f1_label.border = font_red, align_center, border_mid
        row_idx += 1
        
        for nome in lista_f1['Consultor']:
            c = ws.cell(row=row_idx, column=1, value=nome.title())
            c.alignment, c.border = align_center, border_mid
            row_idx += 1
        
        # Espaçamento estético entre Fila 1 e Fila 2
        for _ in range(2): 
            ws.cell(row=row_idx, column=1).border = border_mid
            row_idx += 1

        f2_label = ws.cell(row=row_idx, column=1, value="Fila 2")
        f2_label.font, f2_label.alignment, f2_label.border = font_red, align_center, border_mid
        row_idx += 1
        
        total_f2 = len(lista_f2)
        for i, nome in enumerate(lista_f2['Consultor']):
            c = ws.cell(row=row_idx, column=1, value=nome.title())
            c.alignment = align_center
            c.border = border_bot if i == total_f2 - 1 else border_mid
            row_idx += 1
        
        if total_f2 == 0: f2_label.border = border_bot
        row_idx += 2 

    # Finalização
    nome_final = "Fila_do_Lead.xlsx"
    wb.save(nome_final)
    messagebox.showinfo("Sucesso", f"Arquivo '{nome_final}' gerado com sucesso!")
    os.startfile(nome_final)

def importar_planilha():
    """
    INTERAÇÃO COM ARQUIVO: Lê o Excel enviado pelo usuário e mapeia as abas.
    """
    global df_semana, df_consultores, df_mensal
    caminho = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    if caminho:
        try:
            xls = pd.ExcelFile(caminho)
            abas_originais = xls.sheet_names
            # Normaliza nomes de abas para busca flexível
            abas_map = {a.strip().upper(): a for a in abas_originais}
            
            # Busca as abas necessárias (atende variações de nome)
            aba_sem = abas_map.get('BASE LEAD') or abas_map.get('BASE SEMANAL')
            aba_men = abas_map.get('BASE MENSAL')
            aba_con = abas_map.get('CONSULTORES')

            if aba_sem and aba_men and aba_con:
                df_semana = pd.read_excel(xls, aba_sem)
                df_mensal = pd.read_excel(xls, aba_men)
                df_consultores = pd.read_excel(xls, aba_con)
                
                # Normalização das colunas de texto para cruzamento de dados (VLOOKUP/Merge)
                for d in [df_semana, df_mensal, df_consultores]:
                    d.columns = [str(c).strip().title() for c in d.columns]
                    if 'Consultor' in d.columns:
                        d['Consultor'] = d['Consultor'].astype(str).str.strip().str.upper()
                
                messagebox.showinfo("Sucesso", "Todas as bases carregadas com sucesso!")
            else:
                messagebox.showerror("Erro", f"Abas obrigatórias não encontradas!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha crítica: {e}")

# =================================================================
# INTERFACE GRÁFICA (GUI)
# =================================================================
ctk.set_appearance_mode("dark")
app = ctk.CTk()
app.title("Fila Lead Pro")
app.geometry('400x280')

fonte_titulo = ctk.CTkFont(size=24, weight="bold")
fonte_botao = ctk.CTkFont(size=14, weight="bold")

ctk.CTkLabel(app, text="Fila do Lead", font=fonte_titulo).pack(pady=30)

f_main = ctk.CTkFrame(app)
f_main.pack(pady=10, padx=30, fill="both", expand=True)

ctk.CTkButton(f_main, text="Carregar Planilha de Vendas", command=importar_planilha, 
              width=250, height=40, font=fonte_botao).pack(pady=20)

ctk.CTkButton(f_main, text="Gerar Fila do Lead", command=gerar_fila_do_lead, 
              fg_color="#27ae60", hover_color="#1e8449", width=250, height=40, font=fonte_botao).pack(pady=10)

app.mainloop()