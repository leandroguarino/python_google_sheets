# Como manipular planilhas do Google Sheets com Python?
Neste repositório, há um script que lê arquivos CSV de um diretório e atualiza uma planilha no Google Sheets.

## Versão do Python
Python 3.7.6

## Dependências
gspread: pacote necessário para uso da API do Google Sheets.
csv: usado para facilitar a leitura de arquivos CSV.
codecs: usado para ler um arquivo na codificação correta.

## Antes de executar o script
Para você conseguir manipular planilhas do Google Sheets, você precisa criar um projeto no [Google Developers Console] (https://console.developers.google.com/).
Em seguida, você precisará Ativar a API do Google Drive e Google Sheets no projeto.
Por fim, crie uma conta de serviço. Siga este [passo-a-passo](https://support.google.com/a/answer/7378726?hl=pt-BR).
No final do processo, você baixará um arquivo JSON.
Salve esse arquivo no mesmo diretório do script **import.py**.

O script está comentado linha-a-linha.
