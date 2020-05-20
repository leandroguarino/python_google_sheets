import csv
import gspread
import codecs
from os import listdir
from os.path import isfile, join
import os
import time

#define os parâmetros de conexão com o Google Sheets
scope = ['https://spreadsheets.google.com/feeds']
gc = gspread.service_account(filename='nome_do_arquivo_json_da_sua_conta_de_servico_no_google_api_console.json')

#Troque aqui o link da planilha no Google Drive
sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/asasdasdabdefasdbakbq97asjhbjadjdbashdas87asbbdashk/')
#seleciona a página da planilha pelo nome
worksheet = sheet.worksheet("Municipios_Estado_Brasil")

#seleciona todos os nomes de cidades para não ter que consultar a planilha toda vez
cities = worksheet.col_values(1)

#função para procurar uma cidade na lista de cidades
def findCity(city):
    for i in range(len(cities)):
        if (cities[i] == city):
            return i + 1
    return None

#caminho dos arquivos CSV
mypath = "./csv/"

#filtra apenas os arquivos do diretório
onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

#varre cada arquivo do diretório
for file in onlyfiles:
    filename = file.replace("Mensal-Município","").strip() #alguns arquivos começam com a string "Mensal-Municipio" seguida do nome do município, por isso a string é substituída por vazio
    if (filename[:3] == "OK-"): #arquivos que começam com "OK-" já foram lidos e inseridos na planilha, por isso são desconsiderados
        continue
    splitFilename = filename.split(".") #separa o nome do arquivo e a extensão para extrair o município
    city = splitFilename[0] #nome do município está no nome do arquivo. Faz o split por "." para separar nome e extensão do arquivo
    try:        
        try:
            rowCity = findCity(city) #procura a cidade na planilha, pegando o número da linha em que está
            if (rowCity is None): #se não encontrar a cidade, imprime uma mensagem
                print("Não encontrou "+city)
            else:
                openfile = codecs.open(mypath + file, 'rb', 'iso-8859-1') #abre o arquivo de acordo com a codificação original.
                csv_reader = csv.reader((line.replace('\0','') for line in openfile), delimiter=';') #o replace é necessário para substituir NUL bytes no arquivo
                line_count = 0 #contador de linhas por arquivo
                for row in csv_reader:
                    if (len(row) > 1): #uma linha válida tem mais de uma coluna no arquivo
                        descricao = row[0] #primeira coluna                    
                        if (descricao in ['HOMICÍDIO DOLOSO (2)']): #se for a linha de Homicídio Doloso, imprime o total do ano
                            total_homicidios = row[13] #pega a coluna 13 da linha do arquivo CSV
                            worksheet.update_cell(rowCity, 40, total_homicidios) #atualiza a coluna 40 da planilha no Google Sheets, na linha da cidade
                            print("city: ", city, "Homicídios dolosos: ", total_homicidios, " row: ", rowCity) #mensagem para acompanhar a execução
                        elif (descricao in ['FURTO - OUTROS']): #se for a linha de Furtos, imprime o total do ano
                            total_furtos = row[13].replace(".","") #pega o valor da coluna 13 e tira o caracter ".", se houver
                            worksheet.update_cell(rowCity, 43, total_furtos) #atualiza a coluna 40 da planilha no Google Sheets, na linha da cidade
                            print("city: ", city, "Furtos: ", total_furtos, " row: ", rowCity) #mensagem para acompanhar a execução
                    line_count += 1 #a cada linha, incrementa o contador de linhas
                os.rename(f"{mypath}{file}",f"{mypath}OK-{file}") #renomeia o arquivo para OK-nomedoarquivo, para marcá-lo como lido
        except:
            print("Erro na cidade: "+city) #caso ocorra alguma exceção, imprime uma mensagem
    except csv.Error:
        print('file %s, line %d: ', (mypath + file, csv_reader.line_num)) #se ocorrer erro na leitura do CSV, indica a linha
    time.sleep(2) #delay para atrasar a leitura de cada arquivo de cada cidade, devido ao limite de atualização da API do Google Sheets (100 updates a cada 100 segundos, por padrão)