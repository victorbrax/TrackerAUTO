from tkinter import PhotoImage, filedialog, Label, Tk, ttk, messagebox
import pandas as pd
import tkinter.ttk as ttk
pd.set_option('display.max_columns', None)
caminho = []
def mesclar():
    if len(caminho) > 2:

        df1 = pd.read_excel(caminho[1])
        df2 = pd.read_csv(caminho[0], encoding='UTF-16 LE', sep='	')
        df3 = df1.merge(df2, left_on='Ticket WHD', right_on='No.', how='outer')


        df3.drop(['Comentários', 'No.', 'Client', 'Location', 'Request Type', 
                  'Priority', 'Data de Criação WHD', 'Alert Level', 'Status_y', 'Tech', 'Subject', 'Request Detail', 'Unnamed: 13', 'Updated'], axis = 1, inplace = True)
        df3.rename(columns={'Notes': 'Comentários', 'Date': 'Data de Criação WHD', 'Status_x': 'Status'}, inplace=True)

        df3 = df3 [['Ticket 4BIZ', 'Ticket WHD', 'Data de Criação 4Biz', 'Data de Criação WHD',
        'Titulo', 'Descrição', 'Modulo', 'Priorização', 'Horas Aptdas', 'Impacto',
        'Urgencia', 'Tipo', 'Solicitante', 'Consultor', 'Status', 'Ticket Fornecedor',
        'Vencimento SLA', 'Ultima Atualização', 'Data Resolvido', 'Comentários']]
        
        df3['Data de Criação 4Biz'] = pd.to_datetime(df3['Data de Criação 4Biz'])

        caminho[2] = caminho[2] + '/Tracker Ninecon.xlsx'
        df3.to_excel(caminho[2], index=False) 
        messagebox.showinfo('We won once again Sencineer!', 'Tracker generated successfully. :)')
        janela.destroy()

def dir_select():
    global path2
    path2 = filedialog.askdirectory(initialdir='/Desktop')
    caminho.append(path2)
    mesclar()

def files_select_csv():
    global path
    path = filedialog.askopenfilename(initialdir='/Desktop',
                                    filetypes=(("CSV files .csv", "*.csv"), 
                                                ("All filess", "*.*")))
    caminho.append(path)

def files_select_xlsx():
    global path
    path = filedialog.askopenfilename(initialdir='/Desktop',
                                    filetypes=(("XLSX files .xlsx", "*.xlsx"), 
                                                ("All files", "*.*")))
    caminho.append(path)

def descricao(texto):
    texto = Label(janela, text=texto, bg='darkslategray', fg='white',pady=10)
    texto.pack()    

janela = Tk()
img = PhotoImage(file=r"C:\Users\joaog\Área de Trabalho\FinalVersionPF\iconenativo.png")
janela.iconphoto(False, img)
janela.title('TRACKER GENERATOR')
janela['background'] = 'darkslategray'
janela.geometry('330x300+600+200')

messagebox.showinfo('WARNING!', 'Please, follow the path selection sequence.')

descricao('Select the WHD Tickets   ')
botao1 = ttk.Button(text='WHD_Tickets.csv', command=files_select_csv)
botao1.pack()

descricao('Select the Ninecon Tickets')
botao2 = ttk.Button(text='Tracker Ninecon XX/XX/XX.xlsx', command=files_select_xlsx)
botao2.pack()

descricao('Select the destination folder:')
botao3 = ttk.Button(text='TRACKER!', command=dir_select)
botao3.pack()


by = Label(janela, text='By: Joao Gomes Intern :)', bg='darkslategray', fg='white', height=19, width=25)
by.pack()

# Recursos
janela.mainloop()   