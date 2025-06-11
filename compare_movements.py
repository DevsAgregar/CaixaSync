import re
import pandas as pd
import os
from typing import Dict, Set, Tuple
from utils import (
    normalizar_filial, sanitizar_nome_arquivo,
    formatar_data, obter_conta_bancaria,
    obter_centro_custo
)

def normalizar_filial_formatada(filial):
    if pd.isnull(filial):
        return ''
    match = re.search(r'Loja\s*0?(\d+)', str(filial), re.IGNORECASE)
    if match:
        return f'Loja {int(match.group(1))}'
    return ''

def normalizar_filial_movimentacoes(filial):
    if pd.isnull(filial):
        return ''
    match = re.search(r'LJ0?(\d+)', str(filial), re.IGNORECASE)
    if match:
        return f'Loja {int(match.group(1))}'
    return ''

def sanitizar_nome_arquivo(nome):
    caracteres_invalidos = r'[\/:*?"<>|]'
    return re.sub(caracteres_invalidos, '_', nome)

def formatar_data(data):
    if pd.isnull(data) or data == '':
        return ''
    return pd.to_datetime(data, dayfirst=True).strftime('%d/%m/%Y')

def conta_bancaria(fp, filial):
    if filial == "Loja 1":
        if fp == "Dinheiro":
            return "CAIXA 01"
        elif fp == "Transferência Pix":
            return "SICOOB"
        elif fp in ["Cartão de Débito VISA/ MASTER", "Cartão de Crédito VISA / MASTER"]:
            return "MAQUINETA ÚNICA PETROLINA"
        elif fp == "Cartão de Débito ELO":
            return "MAQUINETA ÚNICA PETROLINA"
        elif fp == "Cartão de Crédito ELO":
            return "MAQUINETA ÚNICA PETROLINA"
        
    elif filial == "Loja 2":
        if fp == "Dinheiro":
            return "CAIXA 02"
        elif fp == "Transferência Pix":
            return "BRADESCO C/C"
        elif fp == "PIx Instantâneo Bradesco LJ02":
            return "BRADESCO C/C"
        elif fp in ["Cartão de Débito VISA/ MASTER", "Cartão de Crédito VISA / MASTER"]:
            return "MAQUINETA ÚNICA SÃO FRANCISCO"
        elif fp == "Cartão de Débito ELO":
            return "MAQUINETA ÚNICA SÃO FRANCISCO"
        elif fp == "Cartão de Crédito ELO":
            return "MAQUINETA ÚNICA SÃO FRANCISCO"
    return ""

def cruzar_planilhas_movimentacao(arquivo_formatado: str, arquivo_movimentacoes: str, pasta_saida: str) -> None:
    """
    Cruza as planilhas de movimentação e gera os arquivos de saída.
    
    Args:
        arquivo_formatado: Caminho do arquivo formatado da etapa anterior
        arquivo_movimentacoes: Caminho do arquivo de movimentações
        pasta_saida: Pasta onde serão salvos os arquivos resultantes
    """
    df_formatada = pd.read_excel(arquivo_formatado, dtype=str)
    df_mov = pd.read_excel(arquivo_movimentacoes, dtype=str)

    df_formatada['Movimentação'] = df_formatada['Movimentação'].str.strip()
    df_formatada['Valor'] = pd.to_numeric(df_formatada['Valor'].replace(',', '.', regex=True)).round(2)
    df_formatada['Filial'] = df_formatada['Filial'].apply(normalizar_filial_formatada)

    df_mov['Código'] = df_mov['Código'].astype(str).str.strip()
    df_mov['Valor (R$)'] = pd.to_numeric(df_mov['Valor (R$)'].replace(',', '.', regex=True)).round(2)
    df_mov['Filial'] = df_mov['Filial'].apply(normalizar_filial_movimentacoes)

    forma_pagamento_map = {
        (str(linha['Movimentação']), linha['Valor'], linha['Filial']): linha['Forma de Pagamento']
        for _, linha in df_formatada.iterrows()
    }

    chaves_formatada = set(forma_pagamento_map.keys())
    chaves_mov = set((str(linha['Código']), linha['Valor (R$)'], linha['Filial']) for _, linha in df_mov.iterrows())
    chaves_nao_relacionadas = chaves_formatada - chaves_mov

    nao_relacionados = df_formatada[
        df_formatada.apply(lambda linha: (linha['Movimentação'], linha['Valor'], linha['Filial']) in chaves_nao_relacionadas, axis=1)
    ]

    if not nao_relacionados.empty:
        caminho_arquivo_nao_relacionados = os.path.join(pasta_saida, 'Não Relacionados.xlsx')
        nao_relacionados.to_excel(caminho_arquivo_nao_relacionados, index=False)
        print(f'Planilha de lançamentos não relacionados salva em: {caminho_arquivo_nao_relacionados}')

    formas = []
    for _, linha in df_mov.iterrows():
        chave = (str(linha['Código']), linha['Valor (R$)'], linha['Filial'])
        forma = forma_pagamento_map.get(chave, '')
        formas.append(forma)
    df_mov['Forma de Pagamento'] = formas

    df_mov['Conta Bancária'] = [
        conta_bancaria(fp, filial)
        for fp, filial in zip(df_mov['Forma de Pagamento'], df_mov['Filial'])
    ]

    contas = [conta for conta in df_mov['Conta Bancária'].dropna().unique() if str(conta).strip() != ""]

    if not contas:
        print("Nenhum dado compatível encontrado. Nenhuma planilha foi gerada.")
        return

    for conta in contas:
        df_conta = df_mov[df_mov['Conta Bancária'] == conta].copy()

        # Reordenar e formatar as colunas
        df_conta['Data de Competência'] = df_conta['Data Movimentação'].apply(formatar_data)
        df_conta['Data de Vencimento'] = df_conta['Data Movimentação'].apply(formatar_data)
        df_conta['Data de Pagamento'] = ''  # Mantém vazio
        df_conta['Valor'] = pd.to_numeric(df_conta['Valor (R$)']).round(2)  # Garante exatamente 2 casas decimais
        df_conta['Categoria'] = 'Receitas de Vendas'
        df_conta['Descrição'] = df_conta['Código'].apply(lambda x: f"Recebimento Mov. Nº {x}")
        df_conta['Centro de Custo'] = df_conta['Filial'].apply(lambda f: "Loja 01 - Petrolina" if f == "Loja 1" else "Loja 02 - São Francisco")
        df_conta['Observações'] = ''  # Mantém vazio
        df_conta['CNPJ/CPF Cliente/Fornecedor'] = ''

        # Seleciona e reordena as colunas
        df_conta = df_conta[[
            'Data de Competência', 'Data de Vencimento', 'Data de Pagamento',
            'Valor', 'Categoria', 'Descrição', 'Cliente/Fornecedor',
            'CNPJ/CPF Cliente/Fornecedor', 'Centro de Custo', 'Observações'
        ]]

        nome_arquivo = f"{sanitizar_nome_arquivo(conta)}.xlsx"
        caminho_arquivo = os.path.join(pasta_saida, nome_arquivo)
        
        # Criar um ExcelWriter para formatar as células
        with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
            df_conta.to_excel(writer, index=False)
            
            # Obter a planilha ativa
            worksheet = writer.sheets['Sheet1']
            
            # Formatar colunas de data (A, B, C)
            for col in ['A', 'B', 'C']:
                for row in range(2, len(df_conta) + 2):  # +2 porque o Excel começa em 1 e tem cabeçalho
                    cell = f"{col}{row}"
                    if worksheet[cell].value:  # Só formata se tiver valor
                        worksheet[cell].number_format = 'dd/mm/yyyy'
            
            # Formatar coluna de valor (D) - Agora com formato brasileiro
            for row in range(2, len(df_conta) + 2):
                cell = f"D{row}"
                worksheet[cell].number_format = '0.00'  # Formato mais simples para garantir 2 casas decimais
            
            # Formatar colunas de texto (E até J)
            for col in ['E', 'F', 'G', 'H', 'I', 'J']:
                for row in range(2, len(df_conta) + 2):
                    cell = f"{col}{row}"
                    worksheet[cell].number_format = '@'
            
            # Ajustar largura das colunas
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width

        print(f'Arquivo separado salvo para conta "{conta}": {caminho_arquivo}')