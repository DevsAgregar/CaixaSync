# CaixaSync

Sistema para processamento e sincronização de planilhas de movimentação financeira.

## Funcionalidades

1. **Transformação de Planilha HTML**
   - Processa planilhas HTML desformatadas
   - Extrai informações de movimentações
   - Organiza dados em formato padronizado

2. **Cruzamento de Movimentações**
   - Compara movimentações entre diferentes planilhas
   - Identifica lançamentos não relacionados
   - Gera relatórios separados por conta bancária

## Requisitos

- Python 3.8 ou superior
- Dependências listadas em `requirements.txt`

## Instalação

1. Clone o repositório
2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Uso

Execute o programa principal:
```bash
python main.py
```

A interface gráfica será aberta com duas etapas:

### Etapa 1: Transformação da Planilha HTML
1. Selecione a planilha HTML desformatada
2. Escolha a pasta de saída
3. Clique em "Rodar Transformação"

### Etapa 2: Cruzamento de Movimentações
1. Selecione a planilha de movimentações
2. Escolha a pasta para salvar as comparações
3. Clique em "Rodar Comparação"

## Estrutura do Projeto

- `main.py`: Ponto de entrada do programa
- `interface.py`: Interface gráfica do sistema
- `html_reader.py`: Processamento de planilhas HTML
- `compare_movements.py`: Comparação de movimentações
- `utils.py`: Funções utilitárias comuns

## Formatos de Arquivo

### Entrada
- Planilha HTML desformatada (.xlsx, .xls, .csv)
- Planilha de movimentações (.xlsx, .xls, .csv)

### Saída
- Planilha formatada (.xlsx)
- Relatórios por conta bancária (.xlsx)
- Relatório de não relacionados (.xlsx) 