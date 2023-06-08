from tkinter import *
from tkinter import ttk
from tkcalendar import *
from babel import *
import cx_Oracle

# Parâmetros de conexão
host = '10.85.0.73'
servico = 'XE'
usuario = 'SYSTEM'
senha = 'CAIXA'

# Encontra o arquivo que aponta para o banco de dados
cx_Oracle.init_oracle_client(lib_dir="./instantclient_21_10")

# Faz a conexão ao banco de dados
conecta_banco = cx_Oracle.connect(usuario, senha, f'{host}/{servico}')

# Cria um cursor no banco para que seja possível fazer consultas e alterações no banco de dados
cursor = conecta_banco.cursor()

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Esta função executa a consulta que irá preencher a treeview com as informações sobre arequisição
def consultar():
    # Recebe os valores dos campos e os coloca em variáveis que serão posteriormente tratadas para se adequar ao formato do banco de dados (os campos de data têm essa peculiaridade)
    filial = campo_filial.get()
    num_req = campo_requisicao.get()
    data_ini = data_inicial.get_date()
    data_fin = data_final.get_date()

    consulta = 'SELECT NUMPREREQUISICAO, CODFILIAL, DATA, CODFUNCREQ, MOTIVO, SITUACAO, NUMTRANSVENDA FROM PCPREREQMATCONSUMOC WHERE CODFILIAL IN {}'.format(filial)
    
    # Estas condições irão concatenar com o valor da variável 'consulta' dependendo se será fornecido o número da requisição ou as datas inicial e final
    if filial == '':
        consulta += '(1, 3, 4, 5, 6, 7, 17 ,18 ,19 ,20 ,61 ,70)'
    if num_req:
        consulta += 'AND NUMPREREQUISICAO = {}'.format(num_req)
    if data_ini and data_fin:
        # Este trecho vai receber o valor das variáveis que receberam os valores da DateEntry e formatar de uma forma que seja possível fazer a consulta SQL no banco oracle
        data_ini = data_ini.strftime('%d-%b-%Y')
        data_fin = data_fin.strftime('%d-%b-%Y')
        consulta += " AND DATA BETWEEN TO_DATE('{}', 'DD-MON-YYYY') AND TO_DATE('{}', 'DD-MON-YYYY')".format(data_ini, data_fin)
    # Faz a verificação nos radios    
    if valor_radio.get() == 1:
        consulta += "AND SITUACAO = 'A'"
    if valor_radio.get() == 2:
        consulta += "AND SITUACAO = 'C'"
    if valor_radio.get() == 3:
        consulta += "AND SITUACAO = 'L'"
    
    consulta += 'ORDER BY DATA DESC'
        

    # Executa a consulta
    cursor.execute(consulta)

    # Limpa todos os dados que possam estar na treeview
    tree.delete(*tree.get_children())

    # Cria uma configuração externa à treeview que altera a cor das linhas
    tree.tag_configure("aprovados", foreground="blue")
    tree.tag_configure("cancelados", foreground="red")
    tree.tag_configure("pendentes", foreground="black")

    # Imprime linha a linha o resultado na treeview (da primeira até a ultima)
    """    for linha in cursor:
        tree.insert('','end', values=linha)"""
        
    # Faz uma verifica
    for linha in cursor:
        situacao = linha[5]  # Obtém o valor do campo 'SITUACAO' da linha
        if situacao == 'A':
            tree.insert("", "end", values=linha, tags=("aprovados",))
        elif situacao == 'C':
            tree.insert("", "end", values=linha, tags=("cancelados",))
        elif situacao == 'L':
            tree.insert("", "end", values=linha, tags=("pendentes",))
        else:
            tree.insert("", "end", values=linha)
    
#--------------------------------------------------------------------------------------------------------------------------------------------------------------

root = Tk()
# Esse trecho até o geometry() é um algorítmo que faz a janela iniciar bem no centro da tela, independente a resolução do monitor
largura = 800
altura = 500
largura_screen = root.winfo_screenwidth()
altura_screen = root.winfo_screenheight()
posx = largura_screen/2 - largura/2
posy = altura_screen/2 - altura/2
root.geometry('%dx%d+%d+%d' % (largura, altura, posx, posy))

root.title('Status Req. Mat. Consumo')
root.resizable(False,False)

#----------------------------------------------

# Frame que contém as informações iniciais necessárias para a consulta 
frame_cabecalho = Frame(root,
              width=640,
              height=110,
              relief=SOLID,
              bd=1,
              )
frame_cabecalho.pack(pady=(50,10))


#-------------------------------------------------------------

# Cria todos os widgets necessários para captar os dados
label_filial = Label(frame_cabecalho, text='Filial')
label_filial.grid(row=0, column=0, pady=(20,0))

campo_filial = ttk.Combobox(frame_cabecalho, width=4, values=['','1','3','4','5','6','7','17','18','19','20','61','70'])
# Mantém a filial 6 ao iniciar o programa
#campo_filial.current(4)
campo_filial.grid(row=1, column=0, pady=(0,20), padx=(20,0))

label_num_req = Label(frame_cabecalho, text='Nº Requisição')
label_num_req.grid(row=0, column=1, pady=(20,0))

campo_requisicao = Entry(frame_cabecalho)
campo_requisicao.grid(row=1, column=1, pady=(0,20), padx=(20,0))

label_data_ini = Label(frame_cabecalho, text='Data Inicial')
label_data_ini.grid(row=0, column=2, pady=(20,0))

data_inicial = DateEntry(frame_cabecalho, date_pattern='dd/mm/yyyy')
data_inicial.grid(row=1, column=2, pady=(0,20), padx=(20,0))

label_data_fin = Label(frame_cabecalho, text='Data Final')
label_data_fin.grid(row=0,column=3, pady=(20,0))

data_final = DateEntry(frame_cabecalho, date_pattern='dd/mm/yyyy')
data_final.grid(row=1,column=3,  pady=(0,20), padx=(20,20))

#-------------------------------------------------------------------

# Frame apenas para conter os radios
frame_radio = Frame(root)
frame_radio.pack()

valor_radio = IntVar()

radio_aprov = Radiobutton(frame_radio,text='Aprovados', variable = valor_radio, value=1, indicatoron=1)
radio_aprov.grid(row=0, column=0)

radio_cancel = Radiobutton(frame_radio,text='Cancelados', variable = valor_radio, value=2, indicatoron=1)
radio_cancel.grid(row=0,column=1)

radio_pend = Radiobutton(frame_radio,text='Pendentes', variable = valor_radio, value=3, indicatoron=1)
radio_pend.grid(row=0,column=2)

radio_todos = Radiobutton(frame_radio,text='Todos', variable = valor_radio, value=4, indicatoron=1)
radio_todos.grid(row=0, column=3)

radio_todos.select()

btn = Button(root, text='Pesquisar', command=consultar)
btn.pack(pady=(20,30))

#------------------------------------------------------------------------------------------------------------------------------------

# Define o a quantidade de colunas
tree = ttk.Treeview(root, columns=('coluna1','coluna2','coluna3','coluna4','coluna5','coluna6','coluna7'))

# Centraliza o conteúdo de todas as colunas
for coluna in ('coluna1','coluna2','coluna3','coluna4','coluna5','coluna6','coluna7'):
    tree.column(coluna, anchor='center')


# Nomeia o cabeçalho das colunas
tree.heading('coluna1',text='Nº REQ')
tree.heading('coluna2',text='FILIAL')
tree.heading('coluna3',text='DATA')
tree.heading('coluna4',text='CODFUNC')
tree.heading('coluna5',text='MOTIVO')
tree.heading('coluna6',text='STATUS')
tree.heading('coluna7',text='TRANSVENDA')

# Define uma largura pardrão para cada coluna, pode ser expandida normalmente clicando e arrastando
tree.column('coluna1', width=60)
tree.column('coluna2', width=40)
tree.column('coluna3', width=90)
tree.column('coluna4', width=90)
tree.column('coluna5', width=90)
tree.column('coluna6', width=70)
tree.column('coluna7', width=90)

# Por padrão é incluido uma coluna inicial 'obrigatória' (columnId) que pode ser ocultada com o comando abaixo
tree.column('#0',width=0)
tree.pack()

scrollbar = ttk.Scrollbar(root, orient='vertical', command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")

# Define a posição e as dimensões da treeview na janela
tree.place(x=50, y=250, width=700, height=200)

#-----------------------------------------------------------------------------------------------------------------------

frame_legenda = Frame(root)
frame_legenda.pack()
frame_legenda.place(x=50, y=450)

frame_azul = Frame(frame_legenda, width=10, height=10, bg='blue')
frame_azul.grid(row=0, column=0)

label_azul = Label(frame_legenda, text='Aprovado', foreground='blue')
label_azul.grid(row=0, column=1)

frame_vermelho = Frame(frame_legenda, width=10, height=10, bg='red')
frame_vermelho.grid(row=0, column=2)

label_vermelho = Label(frame_legenda, text='Cancelado', foreground='red')
label_vermelho.grid(row=0, column=3)

frame_preto = Frame(frame_legenda, width=10, height=10, bg='black')
frame_preto.grid(row=0, column=4)

label_preto = Label(frame_legenda, text='Pendente')
label_preto.grid(row=0, column=5)

#-----------------------------------------------------------------------------------------------------------

frame_credito = Frame(root)
frame_credito.pack()
frame_credito.place(x=690, y=450)

label_credito = Label(frame_credito, text="By Sipauba", fg='gray')
label_credito.pack()

root.mainloop()
