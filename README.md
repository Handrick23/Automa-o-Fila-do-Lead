
# 🚀 Automação Fila do Lead

A **Automação Fila Lead** é uma solução de automação em Python desenvolvida para transformar bases brutas de vendas em uma **Fila de Atendimento Comercial** inteligente e formatada. O sistema aplica regras de ranqueamento baseadas em performance real, garantindo uma distribuição estratégica de leads para a equipe de vendas.

---

## 📋 Sumário

* [Visão Geral](#visão-geral)
* [Regras de Negócio](#regras-de-negócio)
* [Funcionamento do Algoritmo](#funcionamento-do-algoritmo)
* [Tecnologias Utilizadas](#tecnologias-utilizadas)
* [Estrutura da Planilha de Entrada](#estrutura-da-planilha-de-entrada)
* [Como Utilizar](#como-utilizar)

---

## Visão Geral

O programa automatiza o cruzamento de três fontes de dados (Base Semanal, Base Mensal e Cadastro de Consultores). Em vez de uma ordenação simples, ele utiliza critérios de meritocracia para priorizar quem está performando melhor no período atual.

---

## 🧠 Regras de Negócio

O sistema aplica quatro pilares de decisão para organizar os consultores:

### 1. Disponibilidade (Filtro de Status)

O primeiro passo é a exclusão de consultores indisponíveis.

* **Regra:** Se o campo `Status` (ou `Justificativa`) contiver o termo **"FÉRIAS"**, o consultor é ignorado, independentemente de sua performance anterior.

### 2. Categorização de Performance (ABC)

Os consultores ativos são segmentados em três categorias de acordo com o volume de vendas:

* **Categoria A (Alta Performance):** Consultores que realizaram pelo menos uma venda na **semana passada**.
* **Categoria B (Recuperação):** Consultores que não venderam na semana em questão, mas possuem vendas acumuladas no **mês**.
* **Categoria C (Base/Entrada):** Consultores sem vendas na semana e sem vendas no mês.

### 3. Fila 1 vs Fila 2

Para cada filial regional, a distribuição segue a regra da metade superior:

* **Fila 1 (Prioridade Máxima):** Composta geralmente pelos **50% "melhores"** da Categoria A.
* **Fila 2 (Fluxo Geral):** Composta pelos 50% restantes da Categoria A, somados aos consultores das Categorias B e C, ordenados pela venda total semanal (e, posteriormente, mensal, para os que não venderam na semana em questão).

### 4. Critérios de Desempate e Priorização

A ordenação dentro de cada categoria segue esta hierarquia:

1. **Venda Novo (New Logo):** Prioridade para quem traz novos clientes.
2. **Venda Total:** Volume financeiro total (Clientes Existentes + Clientes Novos).
3. **Aleatoriedade (Shuffle):** Para a Categoria C (quem não vendeu período histórico selecionado), o sistema realiza um sorteio aleatório a cada geração, garantindo que a ordem de recebimento de leads seja justa e não alfabética.

---

## ⚙️ Funcionamento do Algoritmo

O processamento matemático para a divisão das filas utiliza arredondamento para cima, garantindo que em equipes com número ímpar de vendedores, a Fila 1 não seja prejudicada:

Onde  é o número total de vendedores que venderam na semana.

### Tratamento de Dados (Data Cleaning)

Para evitar que erros humanos nas planilhas interrompam o processo, o algoritmo executa:

* **Normalização de Strings:** Remove espaços em branco (`strip`) e converte textos para maiúsculas (`upper`) para garantir o "match" entre as bases.
* **Mapeamento de Filiais:** Agrupa diferentes nomenclaturas de equipes em siglas regionais padrão (Ex: "SP 1" e "Grandes Contas SP" são consolidados como "SPO").
* **Busca Flexível:** O sistema identifica as abas necessárias mesmo que o usuário mude o nome de "Base Lead" para "Base Semanal".

---

## 🛠 Tecnologias Utilizadas

* **Python 3.x**: Linguagem base.
* **Pandas**: Processamento de dados e pivotagem de tabelas.
* **CustomTkinter**: Interface gráfica moderna (GUI) com suporte a Dark Mode.
* **Openpyxl**: Criação e estilização de arquivos Excel célula a célula.
* **Math & OS**: Operações matemáticas e comandos de sistema operacional.

---

## 📊 Estrutura da Planilha de Entrada

Para o funcionamento correto, o arquivo Excel deve conter as seguintes abas (nomes flexíveis):

1. **Base Semanal:** Colunas obrigatórias: `Consultor`, `Tipo Cliente` e `Venda`.
2. **Base Mensal:** Mesma estrutura, mas com o histórico do mês.
3. **Consultores:** Colunas obrigatórias: `Consultor`, `Equipe` e `Justificativa` (Status).

---

## 🚀 Como Utilizar

1. **Execução:** Inicie o programa via terminal ou executável.
2. **Upload:** Clique em "Carregar Planilha de Vendas" e selecione seu arquivo `.xlsx`.
3. **Geração:** Clique em "Gerar Fila do Lead".
4. **Resultado:** O sistema abrirá automaticamente o arquivo `Fila_do_Lead.xlsx` formatado com cabeçalhos azuis e bordas organizadas por filial.

---

> **Desenvolvido por:** Handrick Guimarães
> **Finalidade:** Automação de Inteligência Comercial e Processamento de Dados.

---
