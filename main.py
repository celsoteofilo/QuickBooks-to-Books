import os.path
import tkinter
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter import messagebox


import pandas as pd



co0 = "#fof3f5" # preto
co1 = "#feffff" # branco
co2 = "#3fb5a3" # verde
co3 = "#38576b" # valor
co4 = "#403d3d" # letra
co5 = "#0099ff" # azul claro


janela = tk.Tk()
janela.title("CONVERSOR DE DADOD QUICKBOOKS TO BOOKS ")
janela.geometry("600x700")
#janela.configure(background=co5)

foto1 = tk.PhotoImage(file="/Users/celsoteofilo/Desktop/QuickBooks_From_Books/quickbooks.png")
janela_foto = tkinter.Label(image=foto1)
janela_foto.place(x=250, y=40)


#----------------------- CONVERTE TABELA XLSX EM XLS-----------------------------

def pegadocumeto ():

    global pegadocumeto

    Pega_documeto =  filedialog.askopenfilename()
    documento_salvo = pd.read_excel(Pega_documeto)
    documento_salvo.dropna(how="all",inplace=True,  ) # Tratando a Tabela (- campo vazio)
    documento_salvo.to_excel(Pega_documeto + '.xls', index=True)# Tratandoa Tabela  (Numero e)
    messagebox.showinfo("Convertido com SUSCESSO PARA XLS")

BotaoCarregar = tkinter.Button(text='Converter XLSX PARA XLS', command=pegadocumeto,
                               bg="green", fg='black',width=25, height= 3,
                               font=('helvetica', 14, 'bold'))
BotaoCarregar.place(x= 10 , y=30)


#-----------Funçao - CARREGAR TABELA PARA ALTERAR CAMPO--------------------------------

def getDocumento ():   

    global tabela # Opcao de compartilhar a tabela entre funcoces

    Carregar_Doc = filedialog.askopenfilename() # Abrir tela para localizar xls
    tabela = pd.read_excel (Carregar_Doc) # carrega documnto xls
    messagebox.showinfo('Tabela Carregada', 'Tabela carregada com sucesso!') # Mensagem que foi carregado
    print(tabela.columns)  # Opcional do back and para saber se carregou as tabelas

    lista_tabela = tabela.values.tolist()  # Variavel que armazena a tabela carrega da em LIsta
    #print(lista_tabela)


BotaoCarregar = tkinter.Button(text='Carregar DOCUMENTO', command=getDocumento, #Botao para carregar documento
                               bg="green", fg='black',width=25, height= 3,
                               font=('helvetica', 14, 'bold'))
BotaoCarregar.place(x= 10 , y=150)

#-----------------------------ALTERAR TABELA EM MASSA---------------------------------------


#################-----------------------------

def getDocumento ():

    global tabela

    tabela =  filedialog.askopenfilename()
    #Ok_Up_documento = pd.read_excel(tabela)


    tabela_carregada = xlrd.open_workbook_xls(tabela)
    planilha = tabela_carregada.sheet_by_index(0)

    biblioteca ={}

    for linha in range(planilha.nrows):
        chave = planilha.cell_value(linha, 0)
        valor = planilha.cell_value(linha, 1)
        biblioteca[chave] = valor
    print(biblioteca)



####################-------------------------
def alterar_tabela_massa (): # Funcao para atualizar dados alterados na tabela em massa

    global tabelao_massa

    tabelao_massa =tabela # Peguei a tabela da funcao getdocumento e armazeneei em uma variavel
    tabelao_massa.rename(columns={'Data': 'ANOOOOOO'}, inplace=True)
    print(tabelao_massa) # print da  tabela alterada

    tabelao_massa.to_excel(r'/Users/celsoteofilo/Desktop/tabela_editada.xls', index=False) # Onde fica salva a nova tabela
    messagebox.showinfo('Conversor de Dados - Outsmart - V 1.0 ','ATUALIZADA COM SUCESSO!',)




#-----------------------------------BOTAO ESCOLHA (DE)---------------------------------------------------------
#def puxa_colunas_antigas():


escolha_2= tkinter.Label(text= "Escolha coluna DE :  ",
                                    fg='black',width=25, height= 3,
                                    font=('helvetica', 14, 'bold'))
escolha_2.place(x= 25, y =340)
escolha_2 = ttk.Combobox(width=25, height= 3,
                        font=('helvetica', 14, 'bold'),
                         values=["Coluna A","Coluna B","Coluna C","Coluna D","Coluna E","Coluna F","Coluna G",
                         "Coluna H","Coluna I","Coluna J","Coluna L ","Coluna M","Coluna N","Coluna O",])
escolha_2.place(x = 300, y = 360)

botao_atualizar = tkinter.Button(text='Up dados da tabela ',
                                   bg='green', fg='black',width=25, height= 3,
                                   font=('helvetica', 14, 'bold'))
botao_atualizar.place(x= 150, y=500)


#--------------------------------- BOTAO ESCOLHA (PARA)-----------------------------------------------------------

escolha_3 = tkinter.Label(text= "Escolha coluna PARA  :  ",
                                    fg='black',width=25, height= 3,
                                    font=('helvetica', 14, 'bold'))

escolha_3.place(x= 25, y =440)
escolha_3 = ttk.Combobox(width=25, height= 3,
                                    font=('helvetica', 14, 'bold'),values=["Coluna A","Coluna B","Coluna C","Coluna D","Coluna E","Coluna F","Coluna G",
                                           "Coluna H","Coluna I","Coluna J","Coluna L ","Coluna M","Coluna N","Coluna O",])
escolha_3.place(x = 300, y = 460)



#---------------------------------BOTAO ATUALIZAR _------------------------------------------------------------------

botao_atualizar = tkinter.Button(text='ATUALIZAR', command=alterar_tabela_massa,
                                   bg='green', fg='black',width=25, height= 3,
                                   font=('helvetica', 14, 'bold'))
botao_atualizar.place(x= 150, y=550)


janela .mainloop()


# --------------------------  BUSCAR NOME DAS COLUNAS ---------------------------------------
'''
def busca_coluna():

    e = tabelao_massa
    print(f'foi encontrdado {e}')

    nome_arquivo = filedialog.askopenfilenames()

    # Carregando o arquivo para um objeto pandas
    arquivo = pd.ExcelFile(nome_arquivo)

    # Pegando a primeira planilha do arquivo
    planilha = arquivo.sheet_names[0]

    # Carregando a primeira planilha para um objeto dataframe
    df = arquivo.parse(planilha)

    # Imprimindo os nomes das colunas
    print(df.columns)
'''


import xlrd

# Abrir arquivo xls
arquivo_xls = xlrd.open_workbook('tabela.xls')

# Pegar a primeira planilha do arquivo
planilha = arquivo_xls.sheet_by_index(0)

# Criar biblioteca com os dados da planilha
biblioteca = {}
for linha in range(planilha.nrows):  # Percorre as linhas da planilha
    chave = planilha.cell_value(linha, 0)  # Pega o valor da primeira coluna (chave)
    valor = planilha.cell_value(linha, 1)  # Pega o valor da segunda coluna (valor)
    biblioteca[chave] = valor  # Adiciona um item à biblioteca com a chave e o valor correspondente

print(biblioteca)