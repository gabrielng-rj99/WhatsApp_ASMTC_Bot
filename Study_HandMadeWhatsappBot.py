"""     WhatsApp bot handmade

This Code is written in Pt-Br mixed with English because was
not the original idea up to GitHub. And I did for studying.

the idea behind of this code is to get a excel of clients data
with name and number for contact and will dispair messages to
all clients in this DataFrame. First Column was names and second
was numbers. 

1 - The code is made in many steps like:
2 - Browse Chrome.exe and WhatsApp.exe
3 - Browse xlsx file (the dataframe)
4 - Treat Data
5 - Execute Bot

"""


import mouse, keyboard, time, screeninfo, subprocess, configparser, os
import numpy as np
import tkinter as tk
from tkinter import messagebox, filedialog
import pygetwindow as gw
import pandas as pd

## Definindo o Tamanho da Tela e Parâmetros Futuros
screen_size = screeninfo.get_monitors()[0].width, screeninfo.get_monitors()[0].height

## Get the directory of the script
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

#_________________________________________________________________________________________________________

## Definindo MessageBoxes Forçadas
def show_info(info_box):
    root = tk.Tk()
    root.withdraw()  ## Hide the main Tkinter window

    ## janela temporária forçada
    top = tk.Toplevel(root)
    top.overrideredirect(True)
    geometry_string = f"1x1+1+1"
    top.geometry(geometry_string)
    
    top.attributes("-topmost", True)
    top.focus_force()
    top.lift()  ## Raise the window to the top

    ## Show the message box
    messagebox.showinfo(info_box[0], info_box[1], parent=top)

    top.destroy()

def show_warning(info_box):
    root = tk.Tk()
    root.withdraw()  ## Hide the main Tkinter window

    ## janela temporária forçada
    top = tk.Toplevel(root)
    top.overrideredirect(True)
    geometry_string = f"1x1+1+1"
    top.geometry(geometry_string)
    
    top.attributes("-topmost", True)
    top.focus_force()
    top.lift()  ## Raise the window to the top

    ## Show the message box
    messagebox.showwarning(info_box[0], info_box[1], parent=top)

    top.destroy()

def show_error(info_box):
    root = tk.Tk()
    root.withdraw()  ## Hide the main Tkinter window

    ## janela temporária forçada
    top = tk.Toplevel(root)
    top.overrideredirect(True)
    geometry_string = f"1x1+1+1"
    top.geometry(geometry_string)
    
    top.attributes("-topmost", True)
    top.focus_force()
    top.lift()  ## Raise the window to the top

    ## Show the message box
    messagebox.showerror(info_box[0], info_box[1], parent=top)

    top.destroy()
    
#_________________________________________________________________________________________________________
## Definindo Buscar Chrome Diretório
def browse_chrome():
    show_info(("Chrome_Directory", r'''1 – USE SOMENTE SE NÃO ESTIVER SELECIONADO AINDA
2 – Se já estiver selecionado, pode cancelar
3 – Selecione o diretório do chrome.exe

Obs: Normalmente fica em:
C:\Program Files\Google\Chrome\Application''' ))
    
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window

    ## janela temporária forçada
    top = tk.Toplevel(root)
    top.overrideredirect(True)
    geometry_string = "1x1+1+1"
    top.geometry(geometry_string)

    ## Raise the temporary window to the top and give it focus
    top.attributes('-topmost', True)
    top.focus_force()

    file_path = filedialog.askopenfilename(parent=top)
    
    top.destroy()

    if file_path[-10:] == "chrome.exe":
        
        try:
            subprocess.Popen(file_path)
        except:
            show_error(('Erro!', 'O Arquivo chrome.exe foi selecionado incorretamente ou não foi selecionado'))
            return None
            
        time.sleep(.5)
        
        if 'google chrome' in gw.getActiveWindowTitle().lower():

            ## Save the file path in a configuration file
            config = configparser.ConfigParser()
            config.read("config.ini")

            ## Create the 'Paths' section if it doesn't exist
            if "Paths" not in config:
                config["Paths"] = {}

            ## Store the file path under the 'chrome_path' key
            config["Paths"]["chrome_path"] = file_path

            ## Write the updated configuration to the file
            with open("config.ini", "w") as config_file:
                config.write(config_file)


            keyboard.press('alt')
            keyboard.press_and_release('f4')
            keyboard.release('alt')
            return (repr(file_path)[1:-1])
        
        else:
            show_error(('Erro!', 'O arquivo chrome.exe foi selecionado incorretamente'))
            return None

    elif not file_path:
        show_warning(("Aviso!", "Nenhum Arquivo Selecionado"))
        return None
    
    else:
        show_error(('Erro!', 'O Arquivo chrome.exe foi selecionado incorretamente'))
        return None

def get_chrome_path():
    config = configparser.ConfigParser()
    config.read("config.ini")

    if "Paths" in config and "chrome_path" in config["Paths"]:
        return repr(config["Paths"]["chrome_path"])[1:-1]
    else:
        return None

## Definindo Buscar WhatsApp Diretório
def browse_WhatsApp():
    show_info(("WhatsApp_Directory", r'''1 - USE SOMENTE SE NÃO ESTIVER SELECIONADO AINDA
2 - Se já estiver selecionado, pode cancelar
3 - Selecione o diretório do WhatsApp.exe

Obs: Normalmente fica em (pasta oculta):
C:\Program Files\WindowsApps\
"Pesquise WhatsAppDesktop"\

Obs: Se não souber achar, busque ajuda no google/youtube''' ))
    
    root = tk.Tk()
    root.withdraw()  ## Hide the main Tkinter window

    ## janela temporária forçada
    top = tk.Toplevel(root)
    top.overrideredirect(True)
    geometry_string = "1x1+1+1"
    top.geometry(geometry_string)

    ## Raise the temporary window to the top and give it focus
    top.attributes('-topmost', True)
    top.focus_force()

    file_path = filedialog.askopenfilename(parent=top)
    
    top.destroy()

    if file_path[-12:] == "WhatsApp.exe":
        
        try:
            subprocess.Popen(file_path)
        except:
            show_error(('Erro!', 'O Arquivo WhatsApp.exe foi selecionado incorretamente ou não foi selecionado'))
            return None
            
        time.sleep(0.5)
        
        if 'WhatsApp' in gw.getActiveWindowTitle():

            # Save the file path in a configuration file
            config = configparser.ConfigParser()
            config.read("config.ini")

            # Create the 'Paths' section if it doesn't exist
            if "Paths" not in config:
                config["Paths"] = {}

            # Store the file path under the 'chrome_path' key
            config["Paths"]["WhatsApp_path"] = file_path

            # Write the updated configuration to the file
            with open("config.ini", "w") as config_file:
                config.write(config_file)


            keyboard.press('alt')
            keyboard.press_and_release('f4')
            keyboard.release('alt')
            return (repr(file_path)[1:-1])
        
        else:
            show_error(('Erro!', 'O arquivo WhatsApp.exe foi selecionado incorretamente'))
            return None

    elif not file_path:
        show_warning(("Aviso!", "Nenhum Arquivo Selecionado"))
        return None
    
    else:
        show_error(('Erro!', 'O Arquivo WhatsApp.exe foi selecionado incorretamente'))
        return None

def get_WhatsApp_path():
    config = configparser.ConfigParser()
    config.read("config.ini")

    if "Paths" in config and "WhatsApp_path" in config["Paths"]:
        return repr(config["Paths"]["WhatsApp_path"])[1:-1]
    else:
        return None  

##_________________________________________________________________________________________________________
## Definindo buscar arquivo
def browse_xlsx_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window

    ## janela temporária forçada
    top = tk.Toplevel(root)
    top.overrideredirect(True)
    geometry_string = f"1x1+1+1"
    top.geometry(geometry_string)

    ## Raise the temporary window to the top and give it focus
    top.attributes('-topmost', True)
    top.focus_force()

    file_path = filedialog.askopenfilename(parent=top)

    top.destroy()
    
    if file_path[-5:] == '.xlsx':
        return (repr(file_path)[1:-1])
    
    elif file_path and file_path[-5:] != '.xlsx':
        show_error(("Erro!", 'Formato do arquivo não aceito, deve-se colocar um arquivo tipo ".xlsx".'))
        return None 
    else:
        show_warning(("Aviso!", "Nenhum Arquivo Selecionado"))
        return None

##_________________________________________________________________________________________________________
## Definindo Tratamento dos Dados Para Execução Posterior
def treat_data(xlsx_file):
    ## PNC_DF -> Potential New Client Data Frame
    ## Ler Arquivo Excel
    PNC_DF = pd.read_excel(xlsx_file, header=None)

    ## Coverter a planilha com dados (Data Frame) em uma numpy array (matriz)
    PNC_array = PNC_DF.values


    ## Deixar os nomes limpos e prontos pra uso
    for i in range(len(PNC_array)):
        ## Lista Vazia Para Uso
        index_list = []

        if (str(PNC_array[i][0]).lower()) == 'nan':
            pass
        
        ## Trasformar a Célula em Editável        
        cell = list(str(PNC_array[i][0]).upper())

        ## Buscar Índices de Characteres "Errados"
        for char_i in range(len(cell)): # Índice do character na célula
            if cell[char_i].isalpha() or cell[char_i].isspace():
                pass
            else:
                index_list += [char_i]
            
        ## Tratar Dados de Trás pra Frente para evitar "jump" erros
        index_list.reverse()
        for bad_char_i in index_list:
            del cell[bad_char_i]

        if len(cell) == 0:
            PNC_array[i][0] = ''.join(cell)
            pass
        
        else:
            if not cell[-1].isalpha():
                del cell[-1]
        
            if not cell[0].isalpha():
                del cell[0]

        
        ## Arrumando a Célula Depois do Tratamento
        PNC_array[i][0] = ''.join(cell)


    ## Deixar os números limpos e prontos pra uso
    for i in range(len(PNC_array)):
        ## Lista Vazia Para Uso
        index_list = []

        if (str(PNC_array[i][1]).lower()) == 'nan':
            pass
        
        ## Trasformar a Célula em Editável
        cell = list(str(PNC_array[i][1]))

        ## Buscar Índices de Characteres "Errados"
        for char_i in range(len(cell)):     # Índice do character na célula
            if cell[char_i].isdigit():
                pass
            else:
                index_list += [char_i]
            
            
        ## Tratar Dados de Trás pra Frente para evitar "jump" erros
        index_list.reverse()
        for bad_char_i in index_list:
            del cell[bad_char_i]

        if len(cell) == 0:
            PNC_array[i][1] = ''.join(cell)
            pass
        
        else:
            if not cell[-1].isdigit():
                del cell[-1]
        
            if not cell[0].isdigit():
                del cell[0]

        if len(cell) == 13:
            cell[0:2] = []
            
            if cell[0:2] == ['5', '5']:
                cell[0:2] = ['2', '1']

    
        ## Arrumando a Célula Depois do Tratamento
        PNC_array[i][1] = ''.join(cell)


    ## Deletando Celulas Vazias
    index_list = []
    for i in range(len(PNC_array)):
        if PNC_array[i][0].lower() == 'nan' or PNC_array[i][1] == '':
            index_list += [i]

    index_list.reverse()
    
    for i in index_list:
        PNC_array = np.delete(PNC_array, i, 0)

    return PNC_array

##_________________________________________________________________________________________________________   
## Função do Botão da Primeira Janela de Parâmetros (abaixo)
def process_input(user_input, company_input, message_text, window):
    global User, Company, Message
    User = user_input
    Company = company_input
    Message = message_text.replace('{User}', User)
    Message = Message.replace('{Company}', Company)

    window.quit()
    window.destroy()

## Primeira Janela de Parâmetros    
def firstwindow():
   
    bg_color = '#393939' ## Cores Base
    fg_color = '#ffffff'
    win_size = int(0.2225*screen_size[0]), int(0.35*screen_size[1])
    
    ## Tkinter window
    window = tk.Tk()
    window.geometry(f"{win_size[0]}x{win_size[1]}+{(screen_size[0]-520)//2}+{int((screen_size[1]-370)/2.5)}")
    window.resizable(False, False)
    window.overrideredirect(True)
    window.attributes("-topmost", True)
    window.focus_force()
    window.lift()
    window.config(bg = bg_color)

    ## Título + Browser
    label_title = tk.Label(window, text=" Auto Enviar Mensagens WhatsApp                          ",  font=('Bahnschrift Semibold', 18), bg = '#012E26', fg = fg_color)
    label_title.place(x=0, y=-5)
    label_browser = tk.Label(window, text="Browse:", font=('Bahnschrift Semibold', 12), bg = bg_color, fg = fg_color)
    label_browser.place(x = win_size[0] - 150 ,  y=30)
    
    ## Label and Entry for User input
    label_user = tk.Label(window, text="User:", font=('Bahnschrift Semibold', 12), bg = bg_color, fg = fg_color)
    label_user.place(x=12, y=52)
    entry_user = tk.Entry(window, width=50)
    entry_user.place(x=15, y=75)
    entry_user.insert(tk.END, "Fulano")

    ## Label and Entry for Company input
    label_company = tk.Label(window, text="Company:", font=('Bahnschrift Semibold', 12), bg = bg_color, fg = fg_color)
    label_company.place(x=12, y=102)
    entry_company = tk.Entry(window, width=50)
    entry_company.place(x=15, y=125)
    entry_company.insert(tk.END, "YOUR COMPANY")
    
    ## Label and text for message text
    label_message = tk.Label(window, text="Message Text:", font=('Bahnschrift Semibold', 12), bg = bg_color, fg = fg_color)
    label_message.place(x=12, y=152)
    entry_message = tk.Text(window)
    entry_message.place(x=15, y=175, width=(win_size[0]-30), height=(win_size[1]-175-75))
    entry_message.insert(tk.END, """Boa Noite senhor(a) {Cliente},
Meu nome é {User} e sou corretor da {Company}
Consta em nosso sistema que o(a) senhor(a) fez uma pesquisa sobre planos de saúde
Gostaríamos de saber se o(a) senhor(a) já foi atendido(a) e se posso/podemos ajudá-lo(a)?""")

    ## Botões + Função
    button_size = 120, 40
    buttons_x = win_size[0]- button_size[0] - 15

    canvas = tk.Canvas(window, width=button_size[0], height=button_size[1], borderwidth=0, highlightthickness=0)
    canvas.place( x = buttons_x ,  y=55,  width=button_size[0]-2, height=button_size[1]-2)
    button_browse_chrome = tk.Button(window, text="chrome.exe", font=('Bahnschrift Semibold', 12), command=lambda: browse_chrome(), bg='#327DF0', fg='#ffffff')
    button_browse_chrome.place( x = buttons_x + 5 ,  y=55+5 ,  width=button_size[0]-10, height=button_size[1]-10)
    canvas.create_rectangle( 0 , 0,  40, 40,  fill = "#2D9B4B" )    ## green
    canvas.create_rectangle( 40, 0,  80, 40,  fill = "#E64132" )    ## red
    canvas.create_rectangle( 80, 0, 120, 40,  fill = "#FAC319" )    ## yelow


    button_browse_WhatsApp = tk.Button(window, text="WhatsApp.exe", font=('Bahnschrift Semibold', 12), command=lambda: browse_WhatsApp(), bg='#025C4C', fg='white')
    button_browse_WhatsApp.place(x=buttons_x, y=105, width=button_size[0], height=button_size[1])

    button_send = tk.Button(window, text="Enviar Excel", font=('Bahnschrift Semibold', 12), command=lambda: process_input(entry_user.get(), entry_company.get(), entry_message.get("1.0", tk.END), window))
    button_send.place(x=buttons_x, y=(win_size[1]-button_size[1]-15), width=button_size[0], height=button_size[1])

    button_cancel = tk.Button(window, text="Cancelar", font=('Bahnschrift Semibold', 12), command=lambda: quit_app() )
    button_cancel.place( x = 15,  y = (win_size[1]-button_size[1]-15) ,  width = button_size[0] ,  height = button_size[1] )

    ## Botão Fechar
    def quit_app():
        window.destroy()
        window.quit()
        raise SystemExit      

    button_quit_app = tk.Button(window, text="×", font=("Arial", 20), command=lambda: quit_app(), bg='red', fg='white')
    button_quit_app.place(x=win_size[0]-40, y=0, width=40, height=30)
    
    
    # Start the tkinter event loop
    window.mainloop()
    
##__________________________________________________________________________________________________________
## Função Principal 

## Antes, a hotkey para quebrar o cliclo
exit_flag = False
def exit_command():
    global exit_flag
    exit_flag = True

def auto_send_message():
    global exit_flag
    
    ## Aviso anti-lag
    show_warning(("Aviso!", 'Feche tudo que puder para evitar problemas de lentidão que comprometem o uso do programa"'))

    ## Escolhendo os Primeiros Parâmetros
    firstwindow()

    ## Diretório do Chorme e do WhatsApp
    chrome_path = get_chrome_path()
    WhatsApp_path = get_WhatsApp_path()
    
    ## Parâmetros escolhidos, agora vamos ao arquivo Excel
    ## Listando Nomes e Números
    xlsx_file = browse_xlsx_file()
    temp_list = treat_data(xlsx_file)
    
    ## Lista de Nomes e Números JÁ TRATADOS
    List_Names  = list(temp_list.T[0])
    List_Number = list(temp_list.T[1])
    List = [List_Names, List_Number]
    
    ## Condições de Execução

    
    
    if 'WhatsApp' in gw.getAllTitles():
        for wpp in gw.getWindowsWithTitle('WhatsApp'):
            wpp.close()


    ## Programa vai Executar
    show_info(("O Programa Será Executado!", '''O Programa será executado quando der ok, não mexa o mouse ou digite no teclado pois provocará erro.
Para interromper a execução, apenas aperte "End" e espere pois será a última execução.
Precionar "Esc" fará o programa ser interrompido no ato, aperte somente se achar necessário ou em caso de erro.'''))

    ## Abrir o Chrome + Espera
    subprocess.Popen(chrome_path + " --incognito")
    time.sleep(1)
    gw.getWindowsWithTitle('Nova guia anônima - Google Chrome')[0].resizeTo(screen_size[0]//2, screen_size[1])
    gw.getWindowsWithTitle('Nova guia anônima - Google Chrome')[0].moveTo(1, 1)

    ## Abrir o WhatsApp + Espera
    subprocess.Popen(r"C:\Program Files\WindowsApps\5319275A.WhatsAppDesktop_2.2321.4.0_x64__cv1g1gvanyjgm\WhatsApp.exe")
    time.sleep(1)
    gw.getWindowsWithTitle('WhatsApp')[0].resizeTo(screen_size[0]//2, int(screen_size[1]/1.5))
    gw.getWindowsWithTitle('WhatsApp')[0].moveTo(screen_size[0]//12, screen_size[1]//8)
    time.sleep(0.25)

    ## Selecionar uma Conversa para Evitar Bugs
    keyboard.press_and_release('tab')
    time.sleep(0.1)
    keyboard.press_and_release('down')
    time.sleep(0.1)
    keyboard.press_and_release('enter')
    time.sleep(0.1)
    keyboard.press('ctrl')
    keyboard.press_and_release('a')
    keyboard.release('ctrl')
    keyboard.press_and_release('backspace')
    time.sleep(0.2)
    

    ## Definindo Contador, Quebra e Repetição
    n = 0
    keyboard.add_hotkey('end', exit_command, timeout=0)
    
    while n < len(List[0]):
        ## Definindo Nome e Número por coordenada
        name = List[0][n]
        number = List[1][n]
        
        ## Nova Aba
        mouse.move(30,30)
        mouse.click('left')
        keyboard.press('ctrl')
        keyboard.press_and_release('t')
        keyboard.release('ctrl')
        time.sleep(0.25)

        ## Abrindo Nova Conversa
        keyboard.write('wa.me/55'+number)
        keyboard.press_and_release('enter')
        time.sleep(1.5)

        ## Enviando Mensagem
        ready_message = Message.replace('{Cliente}', name)
        keyboard.write(ready_message)
        time.sleep(2)

        ## Conferir Numeros inexistentes
        keyboard.press_and_release('enter')

        ## Fechar Aba
        mouse.move(30, screeninfo.get_monitors()[0].height/3)
        mouse.click('left')
        keyboard.press('ctrl')
        keyboard.press_and_release('w')
        keyboard.release('ctrl')

        ## Interromper Repetição
        if exit_flag:
            break
        
        ## Contador
        n += 1

    ## Fechando o q foi aberto
    mouse.move(30, screeninfo.get_monitors()[0].height/3)
    mouse.click('left')
    keyboard.press('ctrl')
    keyboard.press_and_release('w')
    keyboard.release('ctrl')
    time.sleep(0.5)
    keyboard.press('alt')
    keyboard.press_and_release('f4')
    keyboard.release('alt')

##__________________________________________________________________________________________________________
## executar

auto_send_message()