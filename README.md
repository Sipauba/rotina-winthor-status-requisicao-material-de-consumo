Se você tem interesse em ver a evolução do código do zero, este é o link que possui todas as versões do código: https://github.com/Sipauba/ESTUDO/tree/master/ensaio-nova-rotina-winthor

# Consulta em banco de dados Oracle usando a biblioteca Tkinter em Python

## Introdução

Esta aplicação foi desenvolvida principalmente para fins de aprendizado. Para esta finalidade existem outras ferramentas que também podem ter resultados igualmente satisfatórios e utilizando menos recursos.
A proposta dessa aplicação foi desenvolver uma interface gráfica utilizando a biblioteca Tkinter de Python que se conectasse a um banco de dados Oracle e fizesse uma consulta específica e retornasse esses dados para o usuário. A base de informações utilizadas é referente ao banco de dados WinThor da TOTVS.

## Problemática

No WinThor existe uma rotina que permite aos usuários fazer solicitação de material de consumo a partir do próprio estoque da empresa. Após a requisição, a mesma é avaliada e em seguida pode ser aprovada ou cancelada. Na própia rotina onde é feita a requisição é possível ver se a sua solicitação está pendente ou não, porém, por essa mesma rotina, se a solicitação não está lá, ou foi aprovada ou cancelada e é bem complicado de saber o destino da requisição.

## Solução

Tendo em vista o problema, eis a solução: criar uma aplicação com interface gráfica para o usuário de forma bem simples e que seja semelhante às outras rotinas do Winthor. Essa aplicação irá coletar algumas informações que não serão obrigatórias ao usuário para fazer a pesquisa, como o número da filial que solicitou, o numero da requisição, o intervalo de data e a situação da requisição, se aprovada, pendente, cancelada ou qualquer situação. Ao efetivar a consulta, os dados são exibidos em uma grid no centro da aplicação, simples assim.

# O processo de desenvolvimento
Neste tópico irei discorrer sobre alguns detalhes e particularidades do código e as principais dificuldades que encontrei durante o desenvolvimento.
## - As bibliotecas utilizadas


```bash
from tkinter import *
from tkinter import ttk
from tkcalendar import *
import cx_Oracle
```

Ao fazer o import de todos os médulos da biblioteca tkinter, nem todos os módulos são de fato incluídos. Por isso foi necessário um segundo import com tkinter buscando o módulo ttk, que será responsável por incluir o combobox no campo 'filial' da interface gráfica. A biblioteca cx_Oracle só foi possível instalar após instalar um pacote de ferramentas para desenvolvimento desktop do Visual Studio, é recomendado baixar essas ferramentas para esse tipo de aplicação.

## - Conexão com o banco de dados

Essa sem dúvidas foi a parte mais complicada do desenvolvimento desta aplicação. O banco de dados oracle tem certas peculiaridades das quais não conheço que exigiram paciência e resiliência na hora de ler a documentação da biblioteca utilizada. 

```bash
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
```

Uma atenção especial para a linha: 

```bash
cx_Oracle.init_oracle_client(lib_dir="./instantclient_21_10")
```
Este diretório incluido nesse método foi necessário para se conectar ao banco de dados, ainda não consegui fazer com que a aplicaçã apontasse para o 'tnsnames.ora' da máquina local para que fosse possível encontrar o caminho para o banco de dados. Sem esse arquivo, não seria possível a conexão com o banco. Posteriormente voltarei a comentar sobre esse diretório quando for falar do momento em que empacotei tudo para transformar a aplicação em um arquivo executável.

## - Função que executa a consulta

Um ponto importante é entender o tipo de valor de data que o usuário irá inserir com o tipo de data que o banco irá executar. Foi necessário fazer um breve tratamento do valor que é inserido pelo usuário no campo data e em seguida convertê-lo para um valor que fosse válido para o banco de dados realizar a consulta.

```bash
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
```


## - Interface da aplicação

O restante do código da aplicação é posicionando e formatando cada widget na interface

```bash

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
```

Essa é a aparência da aplicação após uma consulta

![App Screenshot](https://github.com/Sipauba/rotina-winthor-status-requisicao-material-de-consumo/blob/main/rotina.png?raw=true)

## - Utilizando pyinstaler para criar um arquivo executável (.exe)

Um dos desafios depois de escrever todo o código foi como transformar tudo em um arquivo único tendo em vista que nem todos os recursos da aplicação estão contidos no código. A pasta com o cliente oracle, que mencionei anteriormente que é responsável por possibilitar a conexão com o banco. Foi necessário incluir essa pasta juntamente com o código, o que tornou o arquivo final muito volumoso pra uma aplicação que executa comandos tão simples,o que pode acarretar lentidão em alguns computadores que rodarem este programa. Para concluir este tópico, este é o código utilizado para unir a pasta ao código e criar um arquivo executável sem a necessidade de abrir o console junto:

```bash
pyinstaller --add-data "./instantclient_21_10;./instantclient_21_10" --onefile --noconsole rotina.py
```

Depois de gerar o executável esperava-se que a aplicação funcionasse normalmente, já que no VSCode estava funcionando perfeitamente, porém, ao executá-la, ocorreu um erro não previsto:

![erro babel](https://github.com/Sipauba/rotina-winthor-status-requisicao-material-de-consumo/blob/main/erro-babel.png?raw=true)

Essas linhas referem-se a um erro na biblioteca babel que já está incluida no python e que é responsável por interpretar valores númericos. Nesse caso essa biblioteca não estava conseguindo interpretar os valores de data no código, como se a biblioteca babel não estivesse importada no programa. Nesse caso, foi necessário "forçar" a importação do módulo babel no ato de criar o arquivo executável. Foi necessário incluir mais uma tag no código responsável por gerar o executável: 

```bash
pyinstaller --add-data "./instantclient_21_10;./instantclient_21_10" --hidenimport --onefile --noconsole --hidden-import babel.numbers rotina.py
```

Feito isso, a aplicação funciona normalmente a partir do executável.
## Aprendizados

Este projeto, apesar de ter umaproposta bem simples, foi algo que desafiou meus conhecimentos em python e em banco de dados. Anteriormente havia feito outras aplicações muito simples utilizando poucos widgets e que não apresentavam problema algum. Tentei conexão com um banco de dados MySQL e foi muito simples fazer a conexão tanto quanto alterar os dados. O que tornou o projeto mais desafiador foi o fato de ter utilizado um banco de dado oracle a qual eu não tinha experiência em administrá-lo. Muitas vezes precisei recorrer ao chatGPT e a documentação das bibliotecas utilizadas, mas todos os obstáculos foram superados.

## Caso alguém queira utilizar

Caso seja necessário utilizar essa aplicação para outro banco de dados, é necessário buscar outra biblioteca específica para o banco em questão. Com relação ao código do programa, só será necessário substituir o trecho que trata a respeito da conexão com o banco de dados. Este é a pasta com o Oracle client que utilizei: https://drive.google.com/drive/folders/1eTUBgdArT-2DRLo5vJ-Dw295MU7-QpM0?usp=sharing
