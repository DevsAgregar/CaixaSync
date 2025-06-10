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
        'Cartão de Débito ELO',
        'Cartão de Crédito ELO',
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
        # Alteração: Usar apenas o valor principal da movimentação
        valor_num = parse_valor(row[5])  # Mudando de row[8] para row[5] que é a coluna F
        
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
            # Extrai apenas os números da string
            apenas_numeros = ''.join(filter(str.isdigit, str(valor)))
            return len(apenas_numeros) == 6
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
            
            bloco_atual = []
            movimentacao_atual = None
            tipo_operacao = None
            forma_pagamento = None
            usuario_atual = None
            valor_movimentacao = None
            
            total_movimentacoes = 0
            total_linhas_dados = 0
            
            print("\nIniciando processamento linha a linha:")

            for idx, row in df.iterrows():
                # Verifica se a primeira coluna tem valor
                if pd.isna(row[0]):
                    continue
                    
                valor_col0 = str(row[0]).strip()
                
                # Nova movimentação (6 dígitos)
                if self._eh_movimentacao(valor_col0):
                    if bloco_atual:
                        # Só adiciona o bloco se não for uma operação de Saída
                        if tipo_operacao != 'Saída':
                            print(f"\nFinalizando bloco atual. Usuário: '{usuario_atual}'")
                            for item in bloco_atual:
                                item['Forma de Pagamento'] = forma_pagamento or ''
                                item['Usuario'] = usuario_atual or ''
                                item['Valor'] = valor_movimentacao or 0.0  # Usa o valor da movimentação
                                print(f"Adicionando item com usuário: '{item['Usuario']}' e valor: {item['Valor']}")
                                self.dados_formatados.append(item)
                    
                    bloco_atual = []
                    forma_pagamento = None
                    usuario_atual = None
                    valor_movimentacao = None
                    movimentacao_atual = valor_col0
                    tipo_operacao = None
                    total_movimentacoes += 1
                    print(f"\nNova movimentação encontrada: {movimentacao_atual}")
                    continue

                # Tipo de operação e usuário
                if pd.notna(row[4]):
                    valor_col4 = str(row[4]).strip()
                    if valor_col4 in self.TIPOS_OPERACAO:
                        tipo_operacao = valor_col4
                        # Captura o valor da movimentação da coluna F (índice 5)
                        if pd.notna(row[5]):
                            valor_str = str(row[5]).strip()
                            print(f"DEBUG - Valor original: {valor_str}")
                            
                            # Verifica se é um valor estornado (case insensitive)
                            if 'estornado' in valor_str.lower():
                                print(f"DEBUG - Valor estornado encontrado, convertendo para 0")
                                valor = 0.0
                            else:
                                # Remove caracteres especiais primeiro
                                valor_str = valor_str.replace('R$', '').strip()
                                
                                # Determina se o valor é negativo baseado nos parênteses
                                is_negativo = '(' in valor_str and ')' in valor_str
                                
                                # Remove parênteses após verificar se é negativo
                                valor_str = valor_str.replace('(', '').replace(')', '').strip()
                                
                                # Remove o + se existir
                                if valor_str.startswith('+'):
                                    valor_str = valor_str[1:]
                                
                                # Converte para float
                                valor = parse_valor(valor_str)
                                
                                # Ajusta o sinal baseado no tipo de operação
                                if tipo_operacao == 'Entrada':
                                    # Para entradas, sempre deve ser positivo
                                    valor = abs(valor)
                                else:  # Saída
                                    # Para saídas, sempre deve ser negativo
                                    valor = -abs(valor)
                            
                            print(f"Valor da movimentação capturado: {valor}")
                        
                        if pd.notna(row[6]):
                            usuario_atual = str(row[6]).strip()
                            print(f"DEBUG - Capturando usuário da linha. Valor encontrado: '{usuario_atual}'")
                        else:
                            print("DEBUG - Coluna do usuário está vazia!")
                        print(f"Tipo de operação definido para movimentação {movimentacao_atual}: {tipo_operacao}")
                    continue

                # Forma de pagamento
                if pd.notna(row[0]) and str(row[0]).strip() in self.FORMAS_PAGAMENTO_VALIDAS:
                    forma_pagamento = str(row[0]).strip()
                    print(f"Forma de pagamento definida para movimentação {movimentacao_atual}: {forma_pagamento}")
                    continue

                # Linhas de dados (códigos de até 5 dígitos)
                if self._eh_linha_dados(valor_col0):
                    print(f"\nCódigo de linha de dados encontrado: {valor_col0}")
                    if movimentacao_atual and tipo_operacao and tipo_operacao != 'Saída':
                        novo_item = {
                            'Movimentação': movimentacao_atual,
                            'Código': str(row[0]).strip(),
                            'Cliente/Fornecedor': str(row[1]).strip() if pd.notna(row[1]) else '',
                            'Documento': str(row[5]).strip() if pd.notna(row[5]) else '',
                            'Valor': 0.0,  # Será atualizado depois com valor_movimentacao
                            'Forma de Pagamento': None,
                            'Usuario': usuario_atual
                        }
                        print(f"DEBUG - Adicionando linha de dados com usuário: '{usuario_atual}'")
                        bloco_atual.append(novo_item)
                        total_linhas_dados += 1
                        print(f"Linha de dados processada. Total atual: {total_linhas_dados}")
                    else:
                        print(f"AVISO: Linha de dados ignorada - movimentação: {movimentacao_atual}, tipo_operacao: {tipo_operacao}")

            # Processa o último bloco
            if bloco_atual and tipo_operacao != 'Saída':
                print(f"\nProcessando último bloco. Usuário: '{usuario_atual}'")
                for item in bloco_atual:
                    item['Forma de Pagamento'] = forma_pagamento or ''
                    item['Usuario'] = usuario_atual or ''
                    item['Valor'] = valor_movimentacao or 0.0  # Usa o valor da movimentação
                    print(f"Adicionando último item com usuário: '{item['Usuario']}' e valor: {item['Valor']}")
                    self.dados_formatados.append(item)

            print(f"\nResumo do processamento:")
            print(f"Total de movimentações encontradas: {total_movimentacoes}")
            print(f"Total de linhas de dados processadas: {total_linhas_dados}")
            print(f"Total de registros formatados: {len(self.dados_formatados)}")

            # Debug: Mostra todos os registros formatados
            print("\nTodos os registros formatados:")
            for idx, item in enumerate(self.dados_formatados):
                print(f"Registro {idx}:")
                print(f"  Movimentação: {item['Movimentação']}")
                print(f"  Usuário: '{item['Usuario']}'")
                print(f"  Valor: {item['Valor']}")
                print(f"  Outros dados: {item}")

            self._salvar_resultado()
            
        except Exception as e:
            print(f"Erro ao ler o arquivo de entrada: {e}")
            return
    
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
                'Documento', 'Valor', 'Forma de Pagamento', 'Usuario'
            ]
            
            df_formatado = pd.DataFrame(self.dados_formatados, columns=colunas)
            print(f"Dados antes do agrupamento: {len(df_formatado)} linhas")
            print("\nPrimeiras linhas antes do agrupamento:")
            print(df_formatado.head())
            
            # Agrupa por movimentação
            df_agrupado = df_formatado.groupby(['Movimentação']).agg({
                'Código': 'first',
                'Cliente/Fornecedor': 'first',
                'Documento': 'first',
                'Valor': 'first',
                'Forma de Pagamento': 'first',
                'Usuario': 'first'
            }).reset_index()
            
            print(f"\nDados após agrupamento: {len(df_agrupado)} linhas")
            print("\nPrimeiras linhas após agrupamento:")
            print(df_agrupado.head())
            
            # Adiciona coluna Filial usando o usuário e remove Documento
            print("\nAplicando função extrair_loja para determinar a Filial:")
            df_agrupado['Filial'] = df_agrupado['Usuario'].apply(extrair_loja)
            print("\nPrimeiras linhas após determinar Filial:")
            print(df_agrupado[['Usuario', 'Filial']].head())
            
            df_agrupado.drop(columns=['Documento', 'Usuario'], inplace=True)
            
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
        print(f"\nIniciando processamento do arquivo: {caminho_entrada}")
        df = pd.read_excel(caminho_entrada, header=None)
        print(f"Arquivo lido com sucesso. Total de linhas: {len(df)}")
        
        # Primeiro, vamos mapear os usuários e valores para cada movimentação
        print("\nMapeando usuários e valores para cada movimentação...")
        usuarios_por_movimentacao = {}
        valores_por_movimentacao = {}
        tipos_por_movimentacao = {}
        movimentacao_atual = None
        
        for idx, row in df.iterrows():
            # Se é uma nova movimentação
            if pd.notna(row[0]):
                # Extrai apenas os números da string
                apenas_numeros = ''.join(filter(str.isdigit, str(row[0]).strip()))
                if len(apenas_numeros) == 6:
                    movimentacao_atual = apenas_numeros
                    
            # Se é uma linha de tipo de operação com usuário e valor
            if pd.notna(row[4]) and str(row[4]).strip() in ['Entrada', 'Saída']:
                if movimentacao_atual:
                    tipo_operacao = str(row[4]).strip()
                    tipos_por_movimentacao[movimentacao_atual] = tipo_operacao
                    
                    # Captura o usuário
                    if pd.notna(row[6]):
                        usuarios_por_movimentacao[movimentacao_atual] = str(row[6]).strip()
                    
                    # Captura o valor
                    if pd.notna(row[5]):
                        valor_str = str(row[5]).strip()
                        print(f"DEBUG - Valor original para movimentação {movimentacao_atual}: {valor_str}")
                        
                        # Verifica se é um valor estornado (case insensitive)
                        if 'estornado' in valor_str.lower():
                            print(f"DEBUG - Valor estornado encontrado, convertendo para 0")
                            valor = 0.0
                        else:
                            # Remove caracteres especiais primeiro
                            valor_str = valor_str.replace('R$', '').strip()
                            
                            # Determina se o valor é negativo baseado nos parênteses
                            is_negativo = '(' in valor_str and ')' in valor_str
                            
                            # Remove parênteses após verificar se é negativo
                            valor_str = valor_str.replace('(', '').replace(')', '').strip()
                            
                            # Remove o + se existir
                            if valor_str.startswith('+'):
                                valor_str = valor_str[1:]
                        
                            # Converte para float
                            valor = parse_valor(valor_str)
                        
                            # Ajusta o sinal baseado no tipo de operação
                            if tipo_operacao == 'Entrada':
                                # Para entradas, sempre deve ser positivo
                                valor = abs(valor)
                            else:  # Saída
                                # Para saídas, sempre deve ser negativo
                                valor = -abs(valor)
                        
                        valores_por_movimentacao[movimentacao_atual] = valor
                        print(f"Valor capturado para movimentação {movimentacao_atual}: {valor} (Tipo: {tipo_operacao})")

        dados_formatados = []
        bloco_atual = []
        movimentacao_atual = None
        tipo_operacao = None
        forma_pagamento = None

        for i, row in df.iterrows():
            # Identifica nova movimentação
            if pd.notna(row[0]) and str(row[0]).strip().isdigit() and len(str(row[0]).strip()) == 6:
                if bloco_atual:
                    # Só processa se não for uma operação de Saída
                    if tipos_por_movimentacao.get(movimentacao_atual) != 'Saída':
                        usuario_atual = usuarios_por_movimentacao.get(movimentacao_atual, '')
                        valor_atual = valores_por_movimentacao.get(movimentacao_atual, 0.0)
                        print(f"\nFinalizando bloco atual. Usuário: '{usuario_atual}', Valor: {valor_atual}")
                        for item in bloco_atual:
                            item['Forma de Pagamento'] = forma_pagamento if forma_pagamento else ''
                            item['Usuario'] = usuario_atual
                            item['Valor'] = valor_atual
                            print(f"Adicionando item com usuário: '{item['Usuario']}' e valor: {item['Valor']}")
                            dados_formatados.append(item)
                    bloco_atual = []
                    forma_pagamento = None

                movimentacao_atual = str(row[0]).strip()
                tipo_operacao = tipos_por_movimentacao.get(movimentacao_atual)
                print(f"\nNova movimentação encontrada: {movimentacao_atual}")
                continue

            # Captura forma de pagamento
            if pd.notna(row[0]) and str(row[0]).strip() in ProcessadorPlanilha.FORMAS_PAGAMENTO_VALIDAS:
                forma_pagamento = str(row[0]).strip()
                print(f"Forma de pagamento definida para movimentação {movimentacao_atual}: {forma_pagamento}")
                continue

            # Linhas de dados (códigos de até 5 dígitos)
            if pd.notna(row[0]) and str(row[0]).strip().isdigit() and len(str(row[0]).strip()) <= 5:
                # Só processa se não for uma operação de Saída
                if tipo_operacao != 'Saída':
                    novo_item = {
                        'Movimentação': movimentacao_atual,
                        'Código': str(row[0]).strip(),
                        'Cliente/Fornecedor': str(row[1]).strip() if pd.notna(row[1]) else '',
                        'Documento': str(row[5]).strip() if pd.notna(row[5]) else '',
                        'Valor': valores_por_movimentacao.get(movimentacao_atual, 0.0),
                        'Forma de Pagamento': None,
                        'Usuario': usuarios_por_movimentacao.get(movimentacao_atual, '')
                    }
                    print(f"DEBUG - Adicionando linha de dados com usuário: '{novo_item['Usuario']}' e valor: {novo_item['Valor']}")
                    bloco_atual.append(novo_item)

        # Processa o último bloco
        if bloco_atual and tipos_por_movimentacao.get(movimentacao_atual) != 'Saída':
            usuario_atual = usuarios_por_movimentacao.get(movimentacao_atual, '')
            valor_atual = valores_por_movimentacao.get(movimentacao_atual, 0.0)
            print(f"\nProcessando último bloco. Usuário: '{usuario_atual}', Valor: {valor_atual}")
            for item in bloco_atual:
                item['Forma de Pagamento'] = forma_pagamento if forma_pagamento else ''
                item['Usuario'] = usuario_atual
                item['Valor'] = valor_atual
                print(f"Adicionando último item com usuário: '{item['Usuario']}' e valor: {item['Valor']}")
                dados_formatados.append(item)

        colunas = [
            'Movimentação', 'Código', 'Cliente/Fornecedor',
            'Documento', 'Valor', 'Forma de Pagamento', 'Usuario'
        ]
        
        print(f"\nCriando DataFrame com {len(dados_formatados)} registros")
        df_formatado = pd.DataFrame(dados_formatados, columns=colunas)
        print("\nPrimeiras linhas antes do agrupamento:")
        print(df_formatado.head())

        df_agrupado = df_formatado.groupby(['Movimentação']).agg({
            'Código': 'first',
            'Cliente/Fornecedor': 'first',
            'Documento': 'first',
            'Valor': 'first',  # Mantém o primeiro valor já que todos são iguais
            'Forma de Pagamento': 'first',
            'Usuario': 'first'
        }).reset_index()

        print(f"\nDados após agrupamento: {len(df_agrupado)} linhas")
        print("\nPrimeiras linhas após agrupamento:")
        print(df_agrupado.head())
        
        # Cria a nova coluna "Filial" usando o usuário
        print("\nAplicando função extrair_loja para determinar a Filial:")
        df_agrupado['Filial'] = df_agrupado['Usuario'].apply(extrair_loja)
        print("\nPrimeiras linhas após determinar Filial:")
        print(df_agrupado[['Usuario', 'Filial']].head())
        
        # Elimina as colunas "Documento" e "Usuario"
        df_agrupado.drop(columns=['Documento', 'Usuario'], inplace=True)

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
    except Exception as e:
        print(f"Erro ao ler o arquivo de entrada: {e}")
        return