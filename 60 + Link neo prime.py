import pandas as pd

# 1. Carregar os dados
tabela = pd.read_excel("relatorio_os_09-03-2026 05h25.xlsx", sheet_name=0)
tabela_carteira = pd.read_excel("Carteiras Gerenciadoras.xlsx", sheet_name=0)

# 2. Remover espaços dos nomes das colunas
tabela.columns = tabela.columns.str.strip()
tabela_carteira.columns = tabela_carteira.columns.str.strip()

# 3. Merge para trazer o colunas da carteira
if 'Validação NF' in tabela.columns:
    tabela.drop(columns='Validação NF', inplace=True)

if 'Atendimento' in tabela.columns:
    tabela.drop(columns='Atendimento', inplace=True)
    
tabela_final = pd.merge(
    tabela,
    tabela_carteira[['Cod', 'Lider', 'Status', 'Validação NF', 'Atendimento']],
    left_on='Cod Cliente',
    right_on='Cod',
    how='left'
)

tabela_final.drop(columns='Cod', inplace=True)

tabela_final['Lider'] = tabela_final['Lider'].replace('Paulo Videira', 'Paulo Videira/Diego rosa')

# 4. Criar coluna combinando 'Cod Cliente' e 'Cod OS' para controle de duplicatas
tabela_final['cliente_os'] = (
    tabela_final['Cod Cliente'].astype(str) + "-" + tabela_final['Cod OS'].astype(str)
)

# Ver quantidade antes de filtros
print(f"Total inicial: {len(tabela_final)} linhas")


# 8. Filtro: manter apenas registros com 'Qtd Dias em Atraso' >= 60
tabela_final = tabela_final[tabela_final["Qtd Dias em Atraso"] >= 60]
print(f"Após manter Qtd Dias em Atraso >= 60: {len(tabela_final)} linhas")

# 6. Filtro: remover Status OS indesejados
status_indesejados = ["APROVADA", "CANCELADA", "SERVICO REJEITADO"]
tabela_final = tabela_final[~tabela_final["Status OS"].isin(status_indesejados)]
print(f"Após remover Status OS indesejados: {len(tabela_final)} linhas")

# 7. Filtro: remover duplicadas com base na coluna 'cliente_os'
tabela_final = tabela_final.drop_duplicates(subset="cliente_os")
print(f"Após remover duplicadas (cliente_os): {len(tabela_final)} linhas")

#Colocar NF para o Kassio nf
nome_excecao = 'Itamar da Silva Machado Junior'

condicao_nf = (
           (tabela_final['Aba OS'] == 'Validação NF') &
           (tabela_final['Validação NF'].str.lower() == 'equipe plataforma') &
           (tabela_final['Lider'] )
)
tabela_final.loc[condicao_nf, 'Lider'] = 'NF'

condicao_nf_problema = (
           (tabela_final['Aba OS'] == 'NF c/ Problema') &
           (tabela_final['Validação NF'].str.lower() == 'equipe plataforma') &
           (tabela_final['Lider'])
)
tabela_final.loc[condicao_nf_problema, 'Lider'] = 'NF/Problema'

#colocar BOP

cod_bop = (
    (tabela_final['Lider'] == 'Anderson de Oliveira')  &
    (tabela_final['Atendimento'].str.lower() == 'especializada'))

tabela_final.loc[cod_bop, 'Lider'] = 'BOP'

# tirar orçamentista
orcamentista = [
    'Liberação Aprovação',
     'NF Validadas', 
     'Saldo Insuf.', 
     'Lib. Aprovação',
     'Validação 1º Orçamento']

cond_orcamentista = (
            tabela_final['Aba OS'].isin(orcamentista) &
            (tabela_final['Atendimento'] == 'Orçamentista')
        )

tabela_final.loc[cond_orcamentista, 'Lider'] = 'Cliente'


#alterar abas para cliente
abas = [
    'Aguard. Aprovação',
    'Aguard. Envio Veículo',
    'NF Validadas',
    'Saldo Insuf. Lib. Aprovação',
    'Fim Serviço',
    'Inicio Serviço']

cod_abas = (
    tabela_final['Aba OS'].isin(abas)
)
tabela_final.loc[cod_abas, 'Lider'] = 'Cliente'

#alterar OS que é validação cliente
cond_remover = (
    (tabela_final['Validação NF'] == 'Cliente') &
    (tabela_final['Aba OS'].isin(['NF c/ Problema','Validação NF']))
)
tabela_final = tabela_final[~cond_remover]

cliente_so_nf =(
    (tabela_final['Atendimento'] == "Validação de NF ") &
    (tabela_final['Aba OS'].isin(['OS Plataforma',
'Aguard. 1º Orçamento',
'Aguard. Reavaliação',
'Cotações',
'Liberação Aprovação',
'NF Validadas',
'Saldo Insuf. Lib. Aprovação',
'Validação 1º Orçamento'
]))
)
tabela_final = tabela_final[~cliente_so_nf]

#Remover inativos
tabela_final = tabela_final[~tabela_final['Status'].isin(['Inativo'])]
                                          
print(f"Após alterar orçamentistas orçamentistas: {len(tabela_final)} linhas")

# Exibir resultado final
print(tabela_final)


tabela_final.to_excel("tabela_filtrada_final.xlsx", index=False)
