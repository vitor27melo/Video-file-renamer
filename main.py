import tkinter.ttk
from tkinter import *
from tkcalendar import *
from datetime import datetime, timedelta
import win32com.client
import os
import sys
import babel.numbers

sh=win32com.client.gencache.EnsureDispatch('Shell.Application',0)

ws = Tk()
ws.title("Organizador")
ws.geometry("500x450")
ws.config(bg="#cd950c")

hour_string = StringVar()
min_string = StringVar()
last_value_sec = ""
last_value = ""
f = ('Times', 20)

currentDay = datetime.now().day
currentMonth = datetime.now().month
currentYear = datetime.now().year

def display_msg():
    try:
        actionBtn['state'] = DISABLED
        ws.update()
        date = cal.get_date()
        h = min_sb.get()
        m = sec_hour.get()
        s = sec.get()

        data_escolhida = datetime(year=int(date.split('/')[0]),month=int(date.split('/')[1]),day=int(date.split('/')[2]),hour=int(h),minute=int(m),second=int(s))
        lista = []

        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        elif __file__:
            application_path = os.path.dirname(__file__)

        application_path = application_path.replace("/", "\\")

        ns = sh.NameSpace(application_path)
        colnum = 0
        columns = []
        while True:
            colname = ns.GetDetailsOf(None, colnum)
            if not colname:
                break
            columns.append(colname)
            colnum += 1

        cont = 0
        for item in ns.Items():
            if (item.Path[-3:] == 'mp4'):
                cont += 1

        if (cont == 0):
            msg_display.config(text="Não há arquivos para se ordenar.")
            return None

        step = 100/cont

        for item in ns.Items():
            if (item.Path[-3:] == 'mp4'):
                for colnum in range(len(columns)):
                    colval = ns.GetDetailsOf(item, colnum)
                    if colval:
                        if (columns[colnum] == 'Length'):
                            lista.append({
                                "path": item.Path,
                                "length": colval
                            })
                            pb['value'] += step
                            ws.update()

        for index in range(len(lista)):
            if index == 0:
                lista[index]['data'] = data_escolhida
            else:
                tamanho_lista=lista[index]['length'].split(':')
                time_change = timedelta(hours=int(tamanho_lista[0]), minutes=int(tamanho_lista[1]), seconds=int(tamanho_lista[2]))
                data_escolhida = data_escolhida + time_change
                lista[index]['data'] = data_escolhida

        for index in range(len(lista)):
            da = lista[index]['data']
            lista[index]['novo_nome'] = str(da.year) + str(da.month).rjust(2, '0') + str(da.day).rjust(2, '0') + '_' + str(da.hour).rjust(2, '0') + str(da.minute).rjust(2, '0') + str(da.second).rjust(2, '0') + '.mp4'

        for index in range(len(lista)):
            os.rename(lista[index]['path'], './' + lista[index]['novo_nome'])

        with open('log.txt', 'w') as f:
            for index in range(len(lista)):
                f.write('"' + lista[index]['path'].split('\\')[-1] + '" renamed to "' + lista[index]['novo_nome'] + '"\n')

        ws.destroy()
    except Exception as e:
        with open('log.txt', 'w') as f:
            f.write("ERRO FATAL => ")
            f.write(e.args[0])
        ws.destroy()


if last_value == "59" and min_string.get() == "0":
    hour_string.set(int(hour_string.get()) + 1 if hour_string.get() != "23" else 0)
    last_value = min_string.get()

if last_value_sec == "59" and sec_hour.get() == "0":
    min_string.set(int(min_string.get()) + 1 if min_string.get() != "59" else 0)
if last_value == "59":
    hour_string.set(int(hour_string.get()) + 1 if hour_string.get() != "23" else 0)
    last_value_sec = sec_hour.get()

fone = Frame(ws)
ftwo = Frame(ws)

fone.pack(pady=10)
ftwo.pack(pady=10)

cal = Calendar(
    fone,
    date_pattern="y/mm/dd",
    selectmode="day",
    year=currentYear,
    month=currentMonth,
    day=currentDay
)
cal.pack()

min_sb = Spinbox(
    ftwo,
    from_=0,
    to=23,
    wrap=True,
    textvariable=hour_string,
    width=2,
    state="readonly",
    font=f,
    justify=CENTER
)
sec_hour = Spinbox(
    ftwo,
    from_=0,
    to=59,
    wrap=True,
    textvariable=min_string,
    font=f,
    width=2,
    justify=CENTER
)

sec = Spinbox(
    ftwo,
    from_=0,
    to=59,
    wrap=True,
    textvariable=sec_hour,
    width=2,
    font=f,
    justify=CENTER
)

min_sb.pack(side=LEFT, fill=X, expand=True)
sec_hour.pack(side=LEFT, fill=X, expand=True)
sec.pack(side=LEFT, fill=X, expand=True)

msg = Label(
    ws,
    text="Hora  Minuto  Segundo",
    font=("Times", 12),
    bg="#cd950c"
)
msg.pack(side=TOP)

actionBtn = Button(
    ws,
    text="Renomear",
    padx=10,
    pady=10,
    command=display_msg
)
actionBtn.pack(pady=10)

pb = tkinter.ttk.Progressbar(
    ws,
    orient='horizontal',
    mode='determinate',
    length=280
)
# place the progressbar
pb.pack( padx=10, pady=20)

msg_display = Label(
    ws,
    text="",
    bg="#cd950c"
)
msg_display.pack(pady=10)

ws.mainloop()
