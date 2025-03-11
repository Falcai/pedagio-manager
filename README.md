├── pedagio_auto/
│   ├── main.py  # Código principal
│   ├── pedagios.json  # Banco de dados JSON
│   ├── README.md  # Explicação do projeto
│   ├── requirements.txt  # Dependências do projeto
│   ├── .gitignore  # Arquivos a serem ignorados pelo Git

# Conteúdo do requirements.txt
openpyxl
python-docx
tkinter

# Conteúdo do README.md
# Automatização de Preenchimento de Tabela de Pedágios

Este projeto tem como objetivo automatizar o preenchimento de tabelas de pedágios em documentos Word.

## Funcionalidades
- Adiciona entradas em uma tabela existente.
- Gerencia informações de pedágios (CNPJ, Razão Social, Valor).
- Interface gráfica desenvolvida com Tkinter.
- Armazena dados de pedágios em um arquivo JSON.

## Como Usar
1. Instale as dependências com:
   ```bash
   pip install -r requirements.txt
   ```
2. Execute o script principal:
   ```bash
   python main.py
   ```
3. Utilize a interface gráfica para selecionar um arquivo Word e adicionar registros.

## Estrutura do Projeto
- `main.py`: Código principal da aplicação.
- `pedagios.json`: Banco de dados local para armazenar pedágios.
- `requirements.txt`: Lista de dependências necessárias.
- `.gitignore`: Define arquivos a serem ignorados pelo Git.

## Requisitos
- Python 3.8+
- Bibliotecas: `tkinter`, `python-docx`, `openpyxl`

## Licença
Este projeto está sob a licença MIT.
