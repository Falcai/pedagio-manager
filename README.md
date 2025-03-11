# Conteúdo do requirements.txt
openpyxl
python-docx
tkinter

# Automatização de Preenchimento de Tabelas de Pedágios

Este projeto é uma ferramenta de automação para preencher tabelas em documentos Word com dados de pedágios, como data, valor, local e cupom fiscal. Ele utiliza uma interface gráfica simples (GUI) para facilitar a inserção dos dados e gerenciar informações sobre os pedágios.

## Funcionalidades
- Adicionar entradas de pedágios em tabelas de documentos Word (.docx).
- Gerenciar informações de pedágios (local, CNPJ, razão social e valor).
- Formatação automática de células e bordas nas tabelas.
- Validação de dados e tratamento de erros.
  
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

Nome: Felipe Falcai

Email: [felipefalcai@hotmail.com]

GitHub: Falcai
