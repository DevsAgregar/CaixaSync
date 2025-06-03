import re
import pandas as pd
from typing import Union, Optional

def extrair_loja(usuario: str) -> str:
    """
    Determina a loja com base no usuário.
    
    Args:
        usuario: Nome do usuário que fez a movimentação
        
    Returns:
        String formatada 'Loja X' ou string vazia se não encontrar
    """
    print(f"\nExtraindo loja para usuário: '{usuario}'")
    if not usuario:
        print("Usuário vazio, retornando string vazia")
        return ''
    
    usuario = str(usuario).lower().strip()
    print(f"Usuário normalizado: '{usuario}'")
    
    if 'jozimara' in usuario:
        print("Usuário contém jozimara -> Loja 1")
        return 'Loja 1'
    elif 'neide' in usuario:
        print("Usuário contém neide -> Loja 1")
        return 'Loja 1'
    elif 'geizy' in usuario:
        print("Usuário contém geizy -> Loja 2")
        return 'Loja 2'
    
    print(f"Usuário não reconhecido: '{usuario}' -> retornando string vazia")
    return ''

def normalizar_filial(filial: str, padrao: str = r'Loja\s*0?(\d+)') -> str:
    """
    Normaliza o nome da filial para o formato padrão 'Loja X'.
    
    Args:
        filial: String contendo o nome da filial
        padrao: Padrão regex para extrair o número da loja
        
    Returns:
        String normalizada no formato 'Loja X' ou string vazia se não encontrar
    """
    if pd.isnull(filial):
        return ''
    match = re.search(padrao, str(filial), re.IGNORECASE)
    if match:
        return f'Loja {int(match.group(1))}'
    return ''

def parse_valor(valor: Union[str, int, float]) -> float:
    """
    Converte uma string de valor monetário para float.
    
    Args:
        valor: Valor a ser convertido (string, int ou float)
        
    Returns:
        Valor convertido para float
    """
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    
    valor_str = str(valor).replace('R$', '').replace(' ', '')
    
    # Se tiver vírgula, formato brasileiro
    if ',' in valor_str:
        valor_str = valor_str.replace('.', '').replace(',', '.')
        
    try:
        return float(valor_str)
    except Exception:
        return 0.0

def sanitizar_nome_arquivo(nome: str) -> str:
    """
    Remove caracteres inválidos de nomes de arquivo.
    
    Args:
        nome: Nome do arquivo a ser sanitizado
        
    Returns:
        Nome do arquivo sem caracteres inválidos
    """
    caracteres_invalidos = r'[\/:*?"<>|]'
    return re.sub(caracteres_invalidos, '_', nome)

def formatar_data(data: Union[str, pd.Timestamp]) -> str:
    """
    Formata uma data para o padrão dd/mm/yyyy.
    
    Args:
        data: Data a ser formatada
        
    Returns:
        Data formatada como string ou string vazia se inválida
    """
    if pd.isnull(data) or data == '':
        return ''
    return pd.to_datetime(data).strftime('%d/%m/%Y')

def obter_conta_bancaria(forma_pagamento: str, filial: str) -> str:
    """
    Determina a conta bancária com base na forma de pagamento e filial.
    
    Args:
        forma_pagamento: Forma de pagamento utilizada
        filial: Filial onde foi realizada a operação
        
    Returns:
        Nome da conta bancária correspondente
    """
    if filial == "Loja 1":
        if forma_pagamento == "Dinheiro":
            return "CAIXA 01"
        elif forma_pagamento == "Transferência Pix":
            return "SICOOB"
        elif forma_pagamento in ["Cartão de Débito VISA/ MASTER", "Cartão de Crédito VISA / MASTER"]:
            return "MAQUINETA ÚNICA"
            
    elif filial == "Loja 2":
        if forma_pagamento == "Dinheiro":
            return "CAIXA 02"
        elif forma_pagamento in ["Transferência Pix", "PIx Instantâneo Bradesco LJ02"]:
            return "BRADESCO C/C"
        elif forma_pagamento in ["Cartão de Débito VISA/ MASTER", "Cartão de Crédito VISA / MASTER"]:
            return "MAQUINETA ÚNICA"
            
    return ""

def obter_centro_custo(filial: str) -> str:
    """
    Retorna o centro de custo baseado na filial.
    
    Args:
        filial: Nome da filial
        
    Returns:
        Nome do centro de custo correspondente
    """
    return "Loja 01 - Petrolina" if filial == "Loja 1" else "Loja 02 - São Francisco" 