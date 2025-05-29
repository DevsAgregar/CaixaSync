import pandas as pd
import os
from typing import List, Dict, Any
from utils import extrair_loja, parse_valor

class ProcessadorPlanilha:
    """Classe responsável por processar e transformar planilhas HTML desformatadas."""
    
    FORMAS_PAGAMENTO_VALIDAS = {
        'Dinheiro', 'Transferência Pix',
        'Cartão de Débito VISA/ MASTER',
        'Cartão de Crédito VISA / MASTER',
        'PIx Instantâneo Bradesco LJ02'
    }
    
    TIPOS_OPERACAO = {'Entrada', 'Saída'}
    
    def __init__(self, caminho_entrada: str, caminho_saida: str):
        """
        Inicializa o processador de planilhas.
        
        Args:
            caminho_entrada: Caminho do arquivo de entrada
            caminho_saida: Caminho onde será salvo o arquivo processado
        """
        self.caminho_entrada = caminho_entrada
        self.caminho_saida = self._validar_caminho_saida(caminho_saida)
        self.dados_formatados: List[Dict[str, Any]] = []
        print(f"\nCaminhos configurados:")
        print(f"Entrada: {self.caminho_entrada}")
        print(f"Saída: {self.caminho_saida}")
    
    def _validar_caminho_saida(self, caminho: str) -> str:
        """Valida e ajusta o caminho de saída."""
        try:
            # Normaliza o caminho
            caminho = os.path.normpath(caminho)
            
            # Adiciona extensão se necessário
            if not caminho.lower().endswith(('.xls', '.xlsx')):
                caminho += '.xlsx'
                print(f"Adicionada extensão .xlsx ao caminho de saída: {caminho}")
            
            # Cria diretório se não existir
            output_dir = os.path.dirname(caminho)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
                print(f"Diretório criado: {output_dir}")
            elif output_dir:
                print(f"Diretório de saída existe: {output_dir}")
            
            # Verifica permissões de escrita
            if output_dir and not os.access(output_dir, os.W_OK):
                raise PermissionError(f"Sem permissão de escrita no diretório: {output_dir}")
            
            return caminho
            
        except Exception as e:
            print(f"Erro ao validar caminho de saída: {e}")
            raise
    
    def _processar_linha_dados(self, row: pd.Series, movimentacao: str, tipo_operacao: str) -> Dict[str, Any]:
        """Processa uma linha de dados da planilha."""
        valor_num = parse_valor(row[8])
        
        if tipo_operacao == 'Saída' and valor_num > 0:
            valor_num = -valor_num
        elif tipo_operacao == 'Entrada':
            valor_num = abs(valor_num)
            
        return {
            'Movimentação': movimentacao,
            'Código': str(row[0]).strip(),
            'Cliente/Fornecedor': str(row[1]).strip() if pd.notna(row[1]) else '',
            'Documento': str(row[5]).strip() if pd.notna(row[5]) else '',
            'Valor': valor_num,
            'Forma de Pagamento': None
        }
    
    def _eh_movimentacao(self, valor: str) -> bool:
        """Verifica se uma string é um número de movimentação válido."""
        try:
            return valor.isdigit() and len(valor) == 6
        except:
            return False
    
    def _eh_linha_dados(self, valor: str) -> bool:
        """Verifica se uma string é um código de linha de dados válido."""
        try:
            return valor.isdigit() and len(valor) <= 5
        except:
            return False
    
    def processar(self) -> None:
        """Processa a planilha de entrada e gera a saída formatada."""
        try:
            print(f"\nIniciando processamento do arquivo: {self.caminho_entrada}")
            df = pd.read_excel(self.caminho_entrada, header=None)
            print(f"Arquivo lido com sucesso. Total de linhas: {len(df)}")
            
            # Debug: Mostra as primeiras linhas da planilha
            print("\nPrimeiras 5 linhas da planilha:")
            for idx, row in df.head().iterrows():
                print(f"Linha {idx}: {row.tolist()}")
                
        except Exception as e:
            print(f"Erro ao ler o arquivo de entrada: {e}")
            return

        bloco_atual = []
        movimentacao_atual = None
        tipo_operacao = None
        forma_pagamento = None
        
        total_movimentacoes = 0
        total_linhas_dados = 0
        
        print("\nIniciando processamento linha a linha:")

        for idx, row in df.iterrows():
            # Verifica se a primeira coluna tem valor
            if pd.isna(row[0]):
                continue
                
            valor_col0 = str(row[0]).strip()
            
            # Debug: Mostra informações da linha atual
            print(f"\nProcessando linha {idx}:")
            print(f"Valor coluna 0: '{valor_col0}'")
            if pd.notna(row[4]):
                print(f"Valor coluna 4: '{str(row[4]).strip()}'")
            
            # Nova movimentação (6 dígitos)
            if self._eh_movimentacao(valor_col0):
                if bloco_atual:
                    for item in bloco_atual:
                        item['Forma de Pagamento'] = forma_pagamento or ''
                        self.dados_formatados.append(item)
                    bloco_atual = []
                    forma_pagamento = None

                movimentacao_atual = valor_col0
                tipo_operacao = None
                total_movimentacoes += 1
                print(f"Nova movimentação encontrada: {movimentacao_atual}")
                continue

            # Tipo de operação
            if pd.notna(row[4]):
                valor_col4 = str(row[4]).strip()
                if valor_col4 in self.TIPOS_OPERACAO:
                    tipo_operacao = valor_col4
                    print(f"Tipo de operação definido para movimentação {movimentacao_atual}: {tipo_operacao}")
                else:
                    print(f"Valor na coluna 4 não é um tipo de operação válido: '{valor_col4}'")
                continue

            # Forma de pagamento
            if valor_col0 in self.FORMAS_PAGAMENTO_VALIDAS:
                forma_pagamento = valor_col0
                print(f"Forma de pagamento definida para movimentação {movimentacao_atual}: {forma_pagamento}")
                continue

            # Linhas de dados (códigos de até 5 dígitos)
            if self._eh_linha_dados(valor_col0):
                print(f"Código de linha de dados encontrado: {valor_col0}")
                if movimentacao_atual and tipo_operacao:
                    bloco_atual.append(
                        self._processar_linha_dados(row, movimentacao_atual, tipo_operacao)
                    )
                    total_linhas_dados += 1
                    print(f"Linha de dados processada. Total atual: {total_linhas_dados}")
                else:
                    print(f"AVISO: Linha de dados ignorada - movimentação: {movimentacao_atual}, tipo_operacao: {tipo_operacao}")

        # Processa o último bloco
        if bloco_atual:
            for item in bloco_atual:
                item['Forma de Pagamento'] = forma_pagamento or ''
                self.dados_formatados.append(item)

        print(f"\nResumo do processamento:")
        print(f"Total de movimentações encontradas: {total_movimentacoes}")
        print(f"Total de linhas de dados processadas: {total_linhas_dados}")
        print(f"Total de registros formatados: {len(self.dados_formatados)}")

        self._salvar_resultado()
    
    def _salvar_resultado(self) -> None:
        """Salva o resultado processado em um arquivo Excel."""
        if not self.dados_formatados:
            print("\nNenhum dado para processar. Verifique se:")
            print("1. A planilha contém movimentações (números de 6 dígitos)")
            print("2. Cada movimentação tem um tipo de operação (Entrada/Saída)")
            print("3. Existem linhas de dados (códigos de até 5 dígitos)")
            print("4. A estrutura da planilha está no formato esperado")
            return
            
        try:
            print(f"\nPreparando dados para salvar em: {self.caminho_saida}")
            
            colunas = [
                'Movimentação', 'Código', 'Cliente/Fornecedor',
                'Documento', 'Valor', 'Forma de Pagamento'
            ]
            
            df_formatado = pd.DataFrame(self.dados_formatados, columns=colunas)
            print(f"Dados antes do agrupamento: {len(df_formatado)} linhas")
            
            # Agrupa por movimentação
            df_agrupado = df_formatado.groupby(['Movimentação']).agg({
                'Código': 'first',
                'Cliente/Fornecedor': 'first',
                'Documento': 'first',
                'Valor': 'sum',
                'Forma de Pagamento': 'first'
            }).reset_index()
            
            print(f"Dados após agrupamento: {len(df_agrupado)} linhas")
            
            # Adiciona coluna Filial e remove Documento
            df_agrupado['Filial'] = df_agrupado['Documento'].apply(extrair_loja)
            df_agrupado.drop(columns=['Documento'], inplace=True)
            
            # Reorganiza as colunas
            colunas_saida = [
                'Movimentação', 'Código', 'Cliente/Fornecedor',
                'Filial', 'Valor', 'Forma de Pagamento'
            ]
            
            print(f"Salvando arquivo em: {self.caminho_saida}")
            
            # Tenta criar um arquivo temporário primeiro
            temp_file = self.caminho_saida + '.temp'
            df_agrupado.to_excel(
                temp_file,
                index=False,
                columns=colunas_saida,
                engine='openpyxl'
            )
            
            # Se chegou aqui, o arquivo temporário foi criado com sucesso
            # Agora move para o arquivo final
            if os.path.exists(self.caminho_saida):
                os.remove(self.caminho_saida)
            os.rename(temp_file, self.caminho_saida)
            
            print(f"Arquivo salvo com sucesso!")
            print(f"Caminho completo: {os.path.abspath(self.caminho_saida)}")
            
            # Verifica se o arquivo realmente existe
            if os.path.exists(self.caminho_saida):
                tamanho = os.path.getsize(self.caminho_saida)
                print(f"Arquivo criado com {tamanho:,} bytes")
            else:
                raise FileNotFoundError("Arquivo não encontrado após salvamento")
                
        except Exception as e:
            print(f"\nErro ao salvar o arquivo de saída:")
            print(f"Tipo do erro: {type(e).__name__}")
            print(f"Descrição: {str(e)}")
            print(f"Caminho tentado: {self.caminho_saida}")
            raise

def transformar_planilha(caminho_entrada: str, caminho_saida: str) -> None:
    """
    Transforma a planilha HTML desformatada em um formato estruturado.
    
    Args:
        caminho_entrada: Caminho do arquivo de entrada
        caminho_saida: Caminho onde será salvo o arquivo processado
    """
    if not caminho_saida.lower().endswith(('.xls', '.xlsx')):
        caminho_saida += '.xlsx'
        print(f"Adicionada extensão .xlsx ao caminho de saída: {caminho_saida}")
    
    output_dir = os.path.dirname(caminho_saida)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Diretório criado: {output_dir}")

    try:
        df = pd.read_excel(caminho_entrada, header=None)
    except Exception as e:
        print(f"Erro ao ler o arquivo de entrada: {e}")
        return

    dados_formatados = []
    bloco_atual = []
    movimentacao_atual = None
    tipo_operacao = None
    forma_pagamento = None

    for i, row in df.iterrows():
        # Identifica nova movimentação
        if pd.notna(row[0]) and str(row[0]).strip().isdigit() and len(str(row[0]).strip()) == 6:
            if bloco_atual:
                for item in bloco_atual:
                    item['Forma de Pagamento'] = forma_pagamento if forma_pagamento else ''
                    dados_formatados.append(item)
                bloco_atual = []
                forma_pagamento = None

            movimentacao_atual = str(row[0]).strip()
            tipo_operacao = None
            continue

        # Captura tipo de operação
        if pd.notna(row[4]) and str(row[4]).strip() in ['Entrada', 'Saída']:
            tipo_operacao = str(row[4]).strip()
            continue

        # Captura forma de pagamento
        if pd.notna(row[0]) and str(row[0]).strip() in [
            'Dinheiro', 'Transferência Pix',
            'Cartão de Débito VISA/ MASTER',
            'Cartão de Crédito VISA / MASTER',
            'PIx Instantâneo Bradesco LJ02'
        ]:
            forma_pagamento = str(row[0]).strip()
            continue

        # Linhas de dados (códigos de até 5 dígitos)
        if pd.notna(row[0]) and str(row[0]).strip().isdigit() and len(str(row[0]).strip()) <= 5:
            valor_num = parse_valor(row[8])

            if tipo_operacao == 'Saída' and valor_num > 0:
                valor_num = -valor_num
            elif tipo_operacao == 'Entrada':
                valor_num = abs(valor_num)

            bloco_atual.append({
                'Movimentação': movimentacao_atual,
                'Código': str(row[0]).strip(),
                'Cliente/Fornecedor': str(row[1]).strip() if pd.notna(row[1]) else '',
                'Documento': str(row[5]).strip() if pd.notna(row[5]) else '',
                'Valor': valor_num,
                'Forma de Pagamento': None
            })

    if bloco_atual:
        for item in bloco_atual:
            item['Forma de Pagamento'] = forma_pagamento if forma_pagamento else ''
            dados_formatados.append(item)

    colunas = [
        'Movimentação', 'Código', 'Cliente/Fornecedor',
        'Documento', 'Valor', 'Forma de Pagamento'
    ]
    
    df_formatado = pd.DataFrame(dados_formatados, columns=colunas)

    df_agrupado = df_formatado.groupby(['Movimentação']).agg({
        'Código': 'first',
        'Cliente/Fornecedor': 'first',
        'Documento': 'first',
        'Valor': 'sum',
        'Forma de Pagamento': 'first'
    }).reset_index()

    # Cria a nova coluna "Filial" com o formato correto
    df_agrupado['Filial'] = df_agrupado['Documento'].apply(extrair_loja)
    # Elimina a coluna "Documento"
    df_agrupado.drop(columns=['Documento'], inplace=True)

    # Reorganize as colunas conforme desejar
    colunas_saida = [
        'Movimentação', 'Código', 'Cliente/Fornecedor',
        'Filial', 'Valor', 'Forma de Pagamento'
    ]

    try:
        df_agrupado.to_excel(caminho_saida, index=False, columns=colunas_saida, engine='openpyxl')
        print(f"Planilha formatada salva com sucesso em: {caminho_saida}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo de saída: {e}")