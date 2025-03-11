# Automatização de Preenchimento de Tabelas de Pedágios

Este projeto é uma ferramenta de automação para preencher tabelas em documentos Word com dados de pedágios, como data, valor, local e cupom fiscal. Ele utiliza uma interface gráfica simples (GUI) para facilitar a inserção dos dados e gerenciar informações sobre os pedágios.

## Funcionalidades
- Adicionar entradas de pedágios em tabelas de documentos Word (.docx).
- Gerenciar informações de pedágios (local, CNPJ, razão social e valor).
- Formatação automática de células e bordas nas tabelas.
- Validação de dados e tratamento de erros.

## Requisitos
- Python 3.8 ou superior.
- Bibliotecas listadas no arquivo `requirements.txt`.

## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/automatizacao-pedagios.git
Navegue até a pasta do projeto:

bash
Copy
cd automatizacao-pedagios
Crie um ambiente virtual (opcional, mas recomendado):

bash
Copy
python -m venv venv
Ative o ambiente virtual:

No Windows:

bash
Copy
venv\Scripts\activate
No Linux/Mac:

bash
Copy
source venv/bin/activate
Instale as dependências:

bash
Copy
pip install -r requirements.txt
Uso
Execute o script:

bash
Copy
python main.py
Na interface gráfica:

Selecione o documento Word (.docx) que contém a tabela.

Preencha os campos: data, valor, número do cupom fiscal.

Clique em "Adicionar Entrada" para inserir os dados na tabela.

Use o "Gerenciador de Pedágios" para adicionar, remover ou atualizar informações sobre os pedágios.

Estrutura do Projeto
main.py: Código principal da aplicação.

pedagios.json: Armazena os dados dos pedágios (CNPJ, razão social, valor, etc.).

README.md: Documentação do projeto.

requirements.txt: Lista de dependências do projeto.

Contribuição
Contribuições são bem-vindas! Siga os passos abaixo:

Faça um fork do repositório.

Crie uma branch para sua feature/correção:

bash
Copy
git checkout -b minha-feature
Faça commit das suas alterações:

bash
Copy
git commit -m "Adicionando nova funcionalidade"
Envie para o repositório remoto:

bash
Copy
git push origin minha-feature
Abra um Pull Request no GitHub.

Licença
Este projeto está licenciado sob a MIT License.

Contato
Se tiver dúvidas ou sugestões, entre em contato:

Nome: Felipe Falcai

Email: felipefalcai@hotmail.com

GitHub: Falcai
