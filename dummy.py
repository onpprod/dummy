# -*- coding: utf-8 -*-
#Created on Fri May  1 20:14:44 2020
#@author: Otavio Pessoa
#@onpprod

#iniciar com o nome DUMMY

print(" _______   __    __  .___  ___. .___  ___. ____    ____ ")
print("|       \ |  |  |  | |   \/   | |   \/   | \   \  /   / ")
print("|  .--.  ||  |  |  | |  \  /  | |  \  /  |  \   \/   /  ")
print("|  |  |  ||  |  |  | |  |\/|  | |  |\/|  |   \_    _/ ")
print("|  '--'  ||  `--'  | |  |  |  | |  |  |  |     |  |")     
print("|_______/  \______/  |__|  |__| |__|  |__|     |__|")
print("for PAS\n")
print("by:onpprod\n")

#------------------------------------------------------------------------------
#inicializando
#------------------------------------------------------------------------------
import os
import time
import numpy
import openpyxl
path_ini = os.getcwd()


#------------------------------------------------------------------------------
#funções
#------------------------------------------------------------------------------

def get_data (file_path):
    o_file = open(file_path,'r')
    file = o_file.read()
    o_file.close()
    return file
class Erro(Exception):
    pass


#------------------------------------------------------------------------------
#importando configuração
#------------------------------------------------------------------------------

conf_import = None
count = 0
while ((conf_import!='S') and (conf_import!='N')):
    conf_import = str(input('Importar configurações?(S/N)\nR:')).upper()
    if(not(conf_import in 'SN')):
        print('Responda com S ou N')
    if(count>5):
        print('Número excessivo de entradas inválidas')
        time.sleep('7')
        raise Erro('Número excessivo de entradas inválidas')
    count+=1

if conf_import == 'S':
    try:
        conf = get_data('CONFIG_DUM.txt')
        conf = conf.split('\n')
    except:
        conf_import = 'N'
        print('Configurações não encontradas\n',
              'Por favor insira manualmente os dados.\n')


#------------------------------------------------------------------------------
#gerenciamento de local de arquivo
#------------------------------------------------------------------------------

# a partir das configurações
if conf_import == 'S':
    path = conf[0]
    arquivo = str(input('Nome do arquivo:'))
    if('.txt' in arquivo):
        arquivo = arquivo.split('.')
        arquivo = arquivo[0]
# a partir de entradas novas
else:
    path = str(input('Digite o diretório do arquivo: '))+'\\' 
    print('-----Arquivo deve estar em TXT!?-----') 
    arquivo = str(input('Nome do arquivo:'))
    if('.txt' in arquivo):
        arquivo = arquivo.split('.')
        arquivo = arquivo[0]


#------------------------------------------------------------------------------
#verificação de path
#------------------------------------------------------------------------------

if(not(os.path.exists(path+arquivo+'.txt'))):
    print('Arquivo não encontrado')
    time.sleep(7)
    raise FileNotFoundError('Arquivo não encontrado.')
    
else:
    print("Arquivo encontrado...")
    time.sleep(1)
    os.chdir(path)


#------------------------------------------------------------------------------
# entradas de códigos
#------------------------------------------------------------------------------
print("Tratando as entradas...")
try:
# a partir das configurações
    if conf_import == 'S':
        numcods = len(conf)-1
        cods = list(numpy.zeros(numcods,dtype=str))
        for i in range(numcods):
            cods[i] = conf[i+1]
# a partir de entradas novas
    else:
        numcods = int(input('Quantidade de entradas(cod):'))
        cods = list(numpy.zeros(numcods,dtype=str))
        for i in range(numcods):
            q = 'digite o código {}:'.format(i+1)
            cods[i] = str(input(q)).upper()
except:
    print("Erro na entrada de códigos.")
    time.sleep(4)
    raise Erro("Erro na entrada de códigos")
    
print("Sucesso!")


#------------------------------------------------------------------------------
#abrir arquivo txt
#------------------------------------------------------------------------------

dados = get_data(arquivo+'.txt')



#------------------------------------------------------------------------------
#organizar formatação do arquivo
#------------------------------------------------------------------------------
print("Formatando arquivo...")
try:
    dados = dados.split('h ')
    dados_cabecalho = []
    for i in range(len(dados)):
        dados[i] = dados[i].split()
    data = dados[0][0] + dados[0][2] + dados[0][1] + dados[0][4]
    for i in range(8):
        dados_cabecalho.append(dados[0][0])
        del(dados[0][0])
except:
    print("Erro na formatação do arquivo.")
    time.sleep(5)
    raise Erro("Erro na formatação do arquivo.")
print("Sucesso!")

#------------------------------------------------------------------------------   
#gerador de chaves de acesso
#------------------------------------------------------------------------------
print("Gerando chaves de acesso...")
try:
    numkeys = numcods
    keys = list(numpy.zeros(numkeys,dtype=str))
    for i in range(len(keys)):
        for j in range(len(dados[0])):
            if dados[0][j] == cods[i]:
                keys[i] = j+1
                
        if(keys[i] == '0'):
            print('Código não encontrado: {}'.format(cods[i]))
            time.sleep(5)
            raise ValueError('Código não encontrado: {}'.format(cods[i]))
except:
    print("Erro ao gerar chaves.")
    time.sleep(5)
    raise Erro("Erro ao gerar chaves.")
print("Sucesso!")

#------------------------------------------------------------------------------
#modificando para linhas
#------------------------------------------------------------------------------

linha = numpy.zeros([len(dados),numcods+1,],dtype=str).tolist()
linha[0][0] = 'Horário'
try:
    for i in range(len(dados)):
        if i == 0:
            for j in range(numcods):
                linha[0][j+1] = cods[j]
        else:
            for j in range(numcods+1):
                if j == 0 :
                    linha[i][j] = dados[i][0]
                else:
                    try:
                        linha[i][j] = float(dados[i][keys[j-1]])
                    except:
                        linha[i][j] = dados[i][keys[j-1]]
except:
    print('Erro na manipulação de dados')
    time.sleep(5)
    raise Erro('Erro na manipulação de dados')


#------------------------------------------------------------------------------  
#salvando arquivo em excel
#------------------------------------------------------------------------------

arquivo_excel = openpyxl.Workbook()
planilha1 = arquivo_excel.active
planilha1.title = "{} {} {}".format(dados_cabecalho[0],
                                    dados_cabecalho[2],
                                    dados_cabecalho[1])
for i in linha:
    planilha1.append(i)
print(planilha1)
print('Salvo\n')
arquivo_excel.save('arquivo_{}_output.xlsx'.format(arquivo))


#------------------------------------------------------------------------------
#salvar configuração em arquivo txt 
#------------------------------------------------------------------------------

if conf_import == 'N':
    print('Configurações de diretório e códigos usados podem ser salvas.')
    q_ = str(input('Deseja salvar as configurações atuais?(S/N) \nR:')).upper()
    if q_ == 'S':
        os.chdir(path_ini)
        new_conf = open('CONFIG_DUM.txt','w')
        f_path = '{}\n'.format(path)
        new_conf.write(f_path)
        for i in range(len(cods)):
            if cods[i]==cods[-1]:
                f_cod = '{}'.format(cods[i])
                new_conf.write(f_cod)
            else:
                f_cod = '{}\n'.format(cods[i])
                new_conf.write(f_cod)
        new_conf.close()
        print('Configurações salvas.')
    else:
        print('Configurações descartadas.')

#------------------------------------------------------------------------------
print('Fim.')
time.sleep(7)


