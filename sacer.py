import os
import tkinter as tk
import tkcalendar as tc
from tkinter import messagebox
import pandas as pd
import numpy as np
import datetime
import matplotlib.pyplot as plt
import warnings
from dotenv import load_dotenv

load_dotenv()

# Select server and download routes
server_route = os.getenv('server_route')
download_route = os.getenv('download_route')

# Name of the files with the data for suspension authorizations and the broadcasting stations
file_aut_sus = os.getenv('file_aut_sus')
file_estaciones = os.getenv('file_estaciones')

# Columns to be selected in the data files
columnasFM = ['Tiempo', 'Frecuencia (Hz)', 'Level (dBµV/m)', 'Offset (Hz)', 'FM (Hz)', 'Bandwidth (Hz)']
columnasTV = ['Tiempo', 'Frecuencia (Hz)', 'Level (dBµV/m)', 'Offset (Hz)', 'AM (%)', 'Bandwidth (Hz)']
columnasAM = ['Tiempo', 'Frecuencia (Hz)', 'Level (dBµV/m)', 'Offset (Hz)', 'AM (%)', 'Bandwidth (Hz)']
columnasAUT = ['No. INGRESO ARCOTEL', 'FECHA INGRESO', 'NOMBRE ESTACIÓN', 'M/R', 'FREC / CANAL',
               'CIUDAD PRINCIPAL COBERTURA', 'DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'No. OFICIO ARCOTEL',
               'FECHA OFICIO', 'FECHA INICIO SUSPENSION', 'DIAS', 'ZONAL']

# This code only produce a warning that pop-up when using matplotlib for annotations, that is the reason why it is
# disable on purpose. In case a change is made in the code, comment this line to see any other new warning
warnings.filterwarnings('ignore')


class SacerApp(tk.Frame):
    """Create the tkinter application class"""

    def __init__(self, master=None):
        """This is a constructor method for a class in Python. It initializes the object's internal state and is
        automatically called when an object is created."""

        # Use the built-in constructor method of the tkinter module in Python. The super() function is used to call the
        # constructor of the parent class of the current class, and __init__() is the method that is called when
        # creating a new instance of the class. In this case, master is the parameter being passed to the constructor
        # method, which is the main window object that's being created. This line of code is used to initialize an
        # instance of a tkinter widget or frame.
        super().__init__(master)

        # Initialize all the variables
        self.master = master
        self.master.title("Reportes SACER")
        self.master.geometry("800x450")

        self.list_of_cities = ["Ambato", "Cañar", "Cuenca", "Esmeraldas", "Guayaquil", "Ibarra", "Loja", "Macas",
                               "Machala", "Manta", "Nueva Loja", "Puyo", "Quevedo", "Quito", "Riobamba", "Santa Cruz",
                               "Santo Domingo", "Taura", "Tulcan", "Zamora"]

        self.AM_freq = [570000, 590000, 610000, 640000, 670000, 690000, 720000, 740000, 760000, 780000, 800000, 820000,
                        840000, 860000, 880000, 900000, 920000, 940000, 990000, 1020000, 1070000, 1090000, 1110000,
                        1140000, 1180000, 1200000, 1220000, 1240000, 1260000, 1280000, 1310000, 1330000, 1360000,
                        1380000, 1410000, 1430000, 1450000, 1470000, 1490000, 1510000, 1540000, 1580000, 1590000]

        self.FM_freq = [87700000, 87900000, 88100000, 88300000, 88500000, 88700000, 88900000, 89100000, 89300000,
                        89500000, 89700000, 89900000, 90100000, 90300000, 90500000, 90700000, 90900000, 91100000,
                        91300000, 91500000, 91700000, 91900000, 92100000, 92300000, 92500000, 92700000, 92900000,
                        93100000, 93300000, 93500000, 93700000, 93900000, 94100000, 94300000, 94500000, 94700000,
                        94900000, 95100000, 95300000, 95500000, 95700000, 95900000, 96100000, 96300000, 96500000,
                        96700000, 96900000, 97100000, 97300000, 97500000, 97700000, 97900000, 98100000, 98300000,
                        98500000, 98700000, 98900000, 99100000, 99300000, 99500000, 99700000, 99900000, 100100000,
                        100300000, 100500000, 100700000, 100900000, 101100000, 101300000, 101500000, 101700000,
                        101900000, 102100000, 102300000, 102500000, 102700000, 102900000, 103100000, 103300000,
                        103500000, 103700000, 103900000, 104100000, 104300000, 104500000, 104700000, 104900000,
                        105100000, 105300000, 105500000, 105700000, 105900000, 106100000, 106300000, 106500000,
                        106700000, 106900000, 107100000, 107300000, 107500000, 107700000, 107900000]

        self.TV_ch = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32,
                      33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51]

        self.RepGen = tk.BooleanVar()
        self.Ciudad = tk.StringVar()
        self.Ciudad.set(self.list_of_cities[0])
        self.Ocupacion = tk.BooleanVar()
        self.AM_Reporte_individual = tk.BooleanVar()
        self.Frecuencia_AM = tk.IntVar()
        self.Frecuencia_AM.set(self.AM_freq[0])
        self.FM_Reporte_individual = tk.BooleanVar()
        self.Frecuencia_FM = tk.IntVar()
        self.Frecuencia_FM.set(self.FM_freq[0])
        self.TV_Reporte_individual = tk.BooleanVar()
        self.Canal_TV = tk.IntVar()
        self.Canal_TV.set(self.TV_ch[0])
        self.Seleccionar = tk.BooleanVar()
        self.Autorizaciones = tk.BooleanVar()
        self.create_widgets()
        self.program_is_running = False

    def create_widgets(self):
        """Create the widges to be used with tkinter and tkcalendar"""

        self.lbl_2 = tk.Label(self.master, text="Ciudad:", width=20, font=("bold", 11))
        self.lbl_2.grid(row=0, column=0, sticky=tk.W)

        self.option_menu1 = tk.OptionMenu(self.master, self.Ciudad, *self.list_of_cities)
        self.option_menu1.grid(row=0, column=0, sticky=tk.W, padx=130)

        self.lbl_3 = tk.Label(self.master, text="Fecha inicio:", width=10, font=("bold", 11))
        self.lbl_3.grid(row=0, column=0, sticky=tk.W, padx=250)

        self.fecha_inicio = tc.DateEntry(self.master, selectmode='day', date_pattern='yyyy-mm-dd')
        self.fecha_inicio.grid(row=0, column=0, sticky=tk.W, padx=355)

        self.lbl_4 = tk.Label(self.master, text="Fecha fin:", width=10, font=("bold", 11))
        self.lbl_4.grid(row=0, column=0, sticky=tk.W, padx=450)

        self.fecha_fin = tc.DateEntry(self.master, selectmode='day', date_pattern='yyyy-mm-dd')
        self.fecha_fin.grid(row=0, column=0, sticky=tk.W, padx=545)

        self.button0 = tk.Checkbutton(self.master, text="Reporte General.", variable=self.RepGen,
                                      command=lambda: self.toggle_button0_state())
        self.button0.grid(row=1, column=0, sticky=tk.W, padx=30, pady=5)

        self.button1 = tk.Checkbutton(self.master,
                                      text="Ocupación del canal de frecuencias (FCO). Escoger los valores de Umbral.",
                                      variable=self.Ocupacion,
                                      command=lambda: self.toggle_button1_state())
        self.button1.grid(row=2, column=0, sticky=tk.W, padx=30, pady=10)

        self.lbl_6 = tk.Label(self.master, text="Umbral AM:", width=10, font=("bold", 10))
        self.lbl_6.grid(row=3, column=0, sticky=tk.W, padx=60)

        self.Umbral_AM = tk.Entry(self.master)
        self.Umbral_AM.grid(row=3, column=0, sticky=tk.W, padx=160)

        self.lbl_7 = tk.Label(self.master, text="Umbral FM:", width=10, font=("bold", 10))
        self.lbl_7.grid(row=3, column=0, sticky=tk.W, padx=260)

        self.Umbral_FM = tk.Entry(self.master)
        self.Umbral_FM.grid(row=3, column=0, sticky=tk.W, padx=360)

        self.lbl_8 = tk.Label(self.master, text="Umbral TV:", width=10, font=("bold", 10))
        self.lbl_8.grid(row=3, column=0, sticky=tk.W, padx=460)

        self.Umbral_TV = tk.Entry(self.master)
        self.Umbral_TV.grid(row=3, column=0, sticky=tk.W, padx=560)

        self.lbl_9 = tk.Label(self.master, text="Reporte Individual.", width=20, font=('bold', 10))
        self.lbl_9.grid(row=4, column=0, sticky=tk.W, padx=5, pady=10)

        self.button2 = tk.Checkbutton(self.master, text="AM", variable=self.AM_Reporte_individual,
                                      command=lambda: self.toggle_button2_state())
        self.button2.grid(row=5, column=0, sticky=tk.W, padx=30)

        self.option_menu2 = tk.OptionMenu(self.master, self.Frecuencia_AM, self.AM_freq[0], *self.AM_freq)
        self.option_menu2.grid(row=5, column=0, sticky=tk.W, padx=100)

        self.button3 = tk.Checkbutton(self.master, text="FM", variable=self.FM_Reporte_individual,
                                      command=lambda: self.toggle_button3_state())
        self.button3.grid(row=6, column=0, sticky=tk.W, padx=30)

        self.option_menu3 = tk.OptionMenu(self.master, self.Frecuencia_FM, self.FM_freq[0], *self.FM_freq)
        self.option_menu3.grid(row=6, column=0, sticky=tk.W, padx=100)

        self.button4 = tk.Checkbutton(self.master, text="TV", variable=self.TV_Reporte_individual,
                                      command=lambda: self.toggle_button4_state())
        self.button4.grid(row=7, column=0, sticky=tk.W, padx=30)

        self.option_menu4 = tk.OptionMenu(self.master, self.Canal_TV, self.TV_ch[0], *self.TV_ch)
        self.option_menu4.grid(row=7, column=0, sticky=tk.W, padx=100)

        self.button5 = tk.Checkbutton(self.master,
                                      text="Activa la visualización de los periodos sin medición del SACER.",
                                      variable=self.Seleccionar)
        self.button5.grid(row=8, column=0, sticky=tk.W, padx=30, pady=10)

        self.button6 = tk.Checkbutton(self.master,
                                      text="Activa la visualización de los periodos en los que las estaciones de RTV disponen de autorización para suspender emisiones.",
                                      variable=self.Autorizaciones)
        self.button6.grid(row=11, column=0, sticky=tk.W, padx=30, pady=10)

    def toggle_button0_state(self):
        """Define button0 state conditions"""
        if self.RepGen.get():
            self.button1.config(state=tk.DISABLED)
            self.Umbral_AM.config(state=tk.DISABLED)
            self.Umbral_FM.config(state=tk.DISABLED)
            self.Umbral_TV.config(state=tk.DISABLED)
            self.button2.config(state=tk.DISABLED)
            self.option_menu2.config(state=tk.DISABLED)
            self.button3.config(state=tk.DISABLED)
            self.option_menu3.config(state=tk.DISABLED)
            self.button4.config(state=tk.DISABLED)
            self.option_menu4.config(state=tk.DISABLED)
        else:
            self.button1.config(state=tk.NORMAL)
            self.Umbral_AM.config(state=tk.NORMAL)
            self.Umbral_FM.config(state=tk.NORMAL)
            self.Umbral_TV.config(state=tk.NORMAL)
            self.button2.config(state=tk.NORMAL)
            self.option_menu2.config(state=tk.NORMAL)
            self.button3.config(state=tk.NORMAL)
            self.option_menu3.config(state=tk.NORMAL)
            self.button4.config(state=tk.NORMAL)
            self.option_menu4.config(state=tk.NORMAL)

    def toggle_button1_state(self):
        """Define button1 state conditions"""
        if self.Ocupacion.get():
            self.button0.config(state=tk.DISABLED)
            self.button2.config(state=tk.DISABLED)
            self.option_menu2.config(state=tk.DISABLED)
            self.button3.config(state=tk.DISABLED)
            self.option_menu3.config(state=tk.DISABLED)
            self.button4.config(state=tk.DISABLED)
            self.option_menu4.config(state=tk.DISABLED)
            self.button5.config(state=tk.DISABLED)
            self.button6.config(state=tk.DISABLED)
        else:
            self.button0.config(state=tk.NORMAL)
            self.button2.config(state=tk.NORMAL)
            self.option_menu2.config(state=tk.NORMAL)
            self.button3.config(state=tk.NORMAL)
            self.option_menu3.config(state=tk.NORMAL)
            self.button4.config(state=tk.NORMAL)
            self.option_menu4.config(state=tk.NORMAL)
            self.button5.config(state=tk.NORMAL)
            self.button6.config(state=tk.NORMAL)

    def toggle_button2_state(self):
        """Define button2 state conditions"""
        if self.AM_Reporte_individual.get():
            self.button0.config(state=tk.DISABLED)
            self.button1.config(state=tk.DISABLED)
            self.Umbral_AM.config(state=tk.DISABLED)
            self.Umbral_FM.config(state=tk.DISABLED)
            self.Umbral_TV.config(state=tk.DISABLED)
            self.button3.config(state=tk.DISABLED)
            self.option_menu3.config(state=tk.DISABLED)
            self.button4.config(state=tk.DISABLED)
            self.option_menu4.config(state=tk.DISABLED)
        else:
            self.button0.config(state=tk.NORMAL)
            self.button1.config(state=tk.NORMAL)
            self.Umbral_AM.config(state=tk.NORMAL)
            self.Umbral_FM.config(state=tk.NORMAL)
            self.Umbral_TV.config(state=tk.NORMAL)
            self.button3.config(state=tk.NORMAL)
            self.option_menu3.config(state=tk.NORMAL)
            self.button4.config(state=tk.NORMAL)
            self.option_menu4.config(state=tk.NORMAL)

    def toggle_button3_state(self):
        """Define button3 state conditions"""
        if self.FM_Reporte_individual.get():
            self.button0.config(state=tk.DISABLED)
            self.button1.config(state=tk.DISABLED)
            self.Umbral_AM.config(state=tk.DISABLED)
            self.Umbral_FM.config(state=tk.DISABLED)
            self.Umbral_TV.config(state=tk.DISABLED)
            self.button2.config(state=tk.DISABLED)
            self.option_menu2.config(state=tk.DISABLED)
            self.button4.config(state=tk.DISABLED)
            self.option_menu4.config(state=tk.DISABLED)
        else:
            self.button0.config(state=tk.NORMAL)
            self.button1.config(state=tk.NORMAL)
            self.Umbral_AM.config(state=tk.NORMAL)
            self.Umbral_FM.config(state=tk.NORMAL)
            self.Umbral_TV.config(state=tk.NORMAL)
            self.button2.config(state=tk.NORMAL)
            self.option_menu2.config(state=tk.NORMAL)
            self.button4.config(state=tk.NORMAL)
            self.option_menu4.config(state=tk.NORMAL)

    def toggle_button4_state(self):
        """Define button4 state conditions"""
        if self.TV_Reporte_individual.get():
            self.button0.config(state=tk.DISABLED)
            self.button1.config(state=tk.DISABLED)
            self.Umbral_AM.config(state=tk.DISABLED)
            self.Umbral_FM.config(state=tk.DISABLED)
            self.Umbral_TV.config(state=tk.DISABLED)
            self.button2.config(state=tk.DISABLED)
            self.option_menu2.config(state=tk.DISABLED)
            self.button3.config(state=tk.DISABLED)
            self.option_menu3.config(state=tk.DISABLED)
        else:
            self.button0.config(state=tk.NORMAL)
            self.button1.config(state=tk.NORMAL)
            self.Umbral_AM.config(state=tk.NORMAL)
            self.Umbral_FM.config(state=tk.NORMAL)
            self.Umbral_TV.config(state=tk.NORMAL)
            self.button2.config(state=tk.NORMAL)
            self.option_menu2.config(state=tk.NORMAL)
            self.button3.config(state=tk.NORMAL)
            self.option_menu3.config(state=tk.NORMAL)

    def start(self):
        """Define start button actions"""
        if start_button['text'] == 'Iniciar':
            self.program_is_running = True
            start_button['text'] = "Detener"
            self.program()

        else:
            self.program_is_running = False
            start_button['text'] = "Iniciar"

    def quit(self):
        """Define quit button actions"""
        really_quit = messagebox.askyesno("Cerrar?", "Desea cerrar el programa?")
        if really_quit:
            self.master.destroy()

    def program(self):
        """Define base program actions. All data analysis is presented in this section of the code."""
        Ciudad = self.Ciudad.get()
        fecha_inicio = self.fecha_inicio.get_date().strftime("%Y-%m-%d")
        fecha_fin = self.fecha_fin.get_date().strftime("%Y-%m-%d")

        Ocupacion = self.Ocupacion.get()
        if Ocupacion == True:
            if self.Umbral_AM.get() == '':
                Umbral_AM = float(0)
            else:
                Umbral_AM = float(self.Umbral_AM.get())

            if self.Umbral_FM.get() == '':
                Umbral_FM = float(0)
            else:
                Umbral_FM = float(self.Umbral_FM.get())

            if self.Umbral_TV.get() == '':
                Umbral_TV = float(0)
            else:
                Umbral_TV = float(self.Umbral_TV.get())

        AM_Reporte_individual = self.AM_Reporte_individual.get()
        Frecuencia_AM = self.Frecuencia_AM.get()
        FM_Reporte_individual = self.FM_Reporte_individual.get()
        Frecuencia_FM = self.Frecuencia_FM.get()
        TV_Reporte_individual = self.TV_Reporte_individual.get()
        Canal_TV = self.Canal_TV.get()

        Seleccionar = self.Seleccionar.get()

        Autorizaciones = self.Autorizaciones.get()

        # Set the figure size to be used
        plt.rcParams["figure.figsize"] = (20, 10)

        # Declare new variables based in the initial ones
        if Ciudad == 'Tulcan':
            ciu = 'TUL'
            autori = 'TULCÁN'
            sheet_name1 = 'scn-l01FM'
            sheet_name2 = 'scn-l01TV'
        elif Ciudad == 'Ibarra':
            ciu = 'IBA'
            autori = 'IBARRA'
            sheet_name1 = 'scn-l02FM'
            sheet_name2 = 'scn-l02TV'
        elif Ciudad == 'Esmeraldas':
            ciu = 'ESM'
            autori = 'ESMERALDAS'
            sheet_name1 = 'scn-l03FM'
            sheet_name2 = 'scn-l03TV'
        elif Ciudad == 'Nueva Loja':
            ciu = 'NL'
            autori = 'NUEVA LOJA'
            sheet_name1 = 'scn-l05FM'
            sheet_name2 = 'scn-l05TV'
        elif Ciudad == 'Quito':
            ciu = 'UIO'
            autori = 'QUITO'
            sheet_name1 = 'scn-l06FM'
            sheet_name2 = 'scn-l06TV'
            sheet_name3 = 'scn-l06AM'
        elif Ciudad == 'Guayaquil':
            ciu = 'GYE'
            autori = 'GUAYAQUIL'
            sheet_name1 = 'scc-l02FM'
            sheet_name2 = 'scc-l02TV'
            sheet_name3 = 'scc-l02AM'
        elif Ciudad == 'Quevedo':
            ciu = 'QUE'
            autori = 'QUEVEDO'
            sheet_name1 = 'scc-l03FM'
            sheet_name2 = 'scc-l03TV'
        elif Ciudad == 'Machala':
            ciu = 'MACH'
            autori = 'MACHALA'
            sheet_name1 = 'scc-l04FM'
            sheet_name2 = 'scc-l04TV'
        elif Ciudad == 'Taura':
            ciu = 'TAU'
            autori = 'TAURA'
            sheet_name1 = 'scc-l05FM'
            sheet_name2 = 'scc-l05TV'
        elif Ciudad == 'Zamora':
            ciu = 'ZAM'
            autori = 'ZAMORA'
            sheet_name1 = 'scs-l01FM'
            sheet_name2 = 'scs-l01TV'
        elif Ciudad == 'Loja':
            ciu = 'LOJ'
            autori = 'LOJA'
            sheet_name1 = 'scs-l02FM'
            sheet_name2 = 'scs-l02TV'
        elif Ciudad == 'Cañar':
            ciu = 'CAÑ'
            autori = 'CAÑAR'
            sheet_name1 = 'scs-l03FM'
            sheet_name2 = 'scs-l03TV'
        elif Ciudad == 'Macas':
            ciu = 'MAC'
            autori = 'MAC'
            sheet_name1 = 'scs-l04FM'
            sheet_name2 = 'scs-l04TV'
        elif Ciudad == 'Cuenca':
            ciu = 'CUE'
            autori = 'CUE'
            sheet_name1 = 'scs-l05FM'
            sheet_name2 = 'scs-l05TV'
            sheet_name3 = 'scs-l05AM'
        elif Ciudad == 'Riobamba':
            ciu = 'RIO'
            autori = 'RIOBAMBA'
            sheet_name1 = 'scd-l01FM'
            sheet_name2 = 'scd-l01TV'
        elif Ciudad == 'Ambato':
            ciu = 'AMB'
            autori = 'AMBATO'
            sheet_name1 = 'scd-l02FM'
            sheet_name2 = 'scd-l02TV'
        elif Ciudad == 'Puyo':
            ciu = 'PUY'
            autori = 'PUYO'
            sheet_name1 = 'scd-l03FM'
            sheet_name2 = 'scd-l03TV'
        elif Ciudad == 'Manta':
            ciu = 'MAN'
            autori = 'MANTA'
            sheet_name1 = 'scm-l01FM'
            sheet_name2 = 'scm-l01TV'
        elif Ciudad == 'Santo Domingo':
            ciu = 'STO'
            autori = 'SANTO DOMINGO'
            sheet_name1 = 'scm-l02FM'
            sheet_name2 = 'scm-l02TV'
        elif Ciudad == 'Santa Cruz':
            ciu = 'STC'
            autori = 'SANTA CRUZ'
            sheet_name1 = 'erm-l01FM'
            sheet_name2 = 'erm-l01TV'

        # Get the names of the months and the years from de input dates
        Mes_inicio = datetime.datetime.strptime(fecha_inicio, '%Y-%m-%d').strftime("%B")
        Mes_fin = datetime.datetime.strptime(fecha_fin, '%Y-%m-%d').strftime("%B")
        Year1 = datetime.datetime.strptime(fecha_inicio, '%Y-%m-%d').year
        Year2 = datetime.datetime.strptime(fecha_fin, '%Y-%m-%d').year

        # Translate the names of the months to Spanish for the initial date
        if Mes_inicio == 'January':
            Mes_inicio = 'Enero'
        elif Mes_inicio == 'February':
            Mes_inicio = 'Febrero'
        elif Mes_inicio == 'March':
            Mes_inicio = 'Marzo'
        elif Mes_inicio == 'April':
            Mes_inicio = 'Abril'
        elif Mes_inicio == 'May':
            Mes_inicio = 'Mayo'
        elif Mes_inicio == 'June':
            Mes_inicio = 'Junio'
        elif Mes_inicio == 'July':
            Mes_inicio = 'Julio'
        elif Mes_inicio == 'August':
            Mes_inicio = 'Agosto'
        elif Mes_inicio == 'September':
            Mes_inicio = 'Septiembre'
        elif Mes_inicio == 'October':
            Mes_inicio = 'Octubre'
        elif Mes_inicio == 'November':
            Mes_inicio = 'Noviembre'
        elif Mes_inicio == 'December':
            Mes_inicio = 'Diciembre'

        # Translate the names of the months to Spanish for the final date
        if Mes_fin == 'January':
            Mes_fin = 'Enero'
        elif Mes_fin == 'February':
            Mes_fin = 'Febrero'
        elif Mes_fin == 'March':
            Mes_fin = 'Marzo'
        elif Mes_fin == 'April':
            Mes_fin = 'Abril'
        elif Mes_fin == 'May':
            Mes_fin = 'Mayo'
        elif Mes_fin == 'June':
            Mes_fin = 'Junio'
        elif Mes_fin == 'July':
            Mes_fin = 'Julio'
        elif Mes_fin == 'August':
            Mes_fin = 'Agosto'
        elif Mes_fin == 'September':
            Mes_fin = 'Septiembre'
        elif Mes_fin == 'October':
            Mes_fin = 'Octubre'
        elif Mes_fin == 'November':
            Mes_fin = 'Noviembre'
        elif Mes_fin == 'December':
            Mes_fin = 'Diciembre'

        # Create a vector with format "Enero_2021" for all the years in evaluation
        vector = []
        for year in range(int(Year1), int(Year2 + 1)):
            meses = [f"Enero_{year}", f"Febrero_{year}", f"Marzo_{year}", f"Abril_{year}", f"Mayo_{year}",
                     f"Junio_{year}",
                     f"Julio_{year}", f"Agosto_{year}", f"Septiembre_{year}", f"Octubre_{year}", f"Noviembre_{year}",
                     f"Diciembre_{year}"]
            vector.append(meses)
            month_year = [num for elem in vector for num in elem]

        # DATA READING: FM and TV broadcasting cases
        # Specifies the names of the data columns to be used
        df_d1 = []
        df_d2 = []
        # "_".join((Mes_inicio, str(Year1))) = "Mes_inicio_Year1", the content for join must be strings that is why
        # str(Year1) is used
        m = int(month_year.index("_".join((Mes_inicio,
                                           str(Year1)))))
        n = int(month_year.index("_".join((Mes_fin, str(Year2)))))
        for mes in month_year[m:n + 1]:
            # u: read the csv file, used usecols, to list the data and pass it to a numpy array
            try:
                u = pd.read_csv(f'{server_route}/{ciu}/FM_{ciu}_{mes}.csv', engine='python',
                                skipinitialspace=True, usecols=columnasFM, encoding='unicode_escape').to_numpy()
            except IOError:
                # Raise if file does not exist
                u = np.full([1, 6], np.nan)

            # df_d1: append(u) adds all the elements in the u lists generated by the for loop
            df_d1.append(u)

            try:
                v = pd.read_csv(f'{server_route}/{ciu}/TV_{ciu}_{mes}.csv', engine='python',
                                skipinitialspace=True, usecols=columnasTV, encoding='unicode_escape').to_numpy()
            except IOError:
                # raise if file does not exist
                v = np.full([1, 6], np.nan)

            df_d2.append(v)

        # Join all the sequences of arrays in the previous df to get only one array of data
        df_d1 = np.concatenate(df_d1)
        df_d2 = np.concatenate(df_d2)
        # df_original1: convert numpy array df_d1 to pandas dataframe and add header
        df_original1 = pd.DataFrame(df_d1,
                                    columns=['Tiempo', 'Frecuencia (Hz)', 'Level (dBµV/m)', 'Offset (Hz)', 'FM (Hz)',
                                             'Bandwidth (Hz)'])
        df_original2 = pd.DataFrame(df_d2,
                                    columns=['Tiempo', 'Frecuencia (Hz)', 'Level (dBµV/m)', 'Offset (Hz)', 'AM (%)',
                                             'Bandwidth (Hz)'])

        # df7: read the TX.xlsx data, convert it to a pandas dataframe and fill na with -
        df7 = pd.read_excel(f'{server_route}/{file_estaciones}', sheet_name=sheet_name1)
        df7 = df7.fillna('-')
        df8 = pd.read_excel(f'{server_route}/{file_estaciones}', sheet_name=sheet_name2)
        df8 = df8.fillna('-')

        # dfau1: read the data in SUSPENSIÓN EMISIONES-VERIFICACIÓN REINICIO OPERACIÓN.xlsx file and convert
        # it to a pandas dataframe
        dfau1 = pd.read_excel(
            f'{server_route}/{file_aut_sus}',
            skiprows=1, usecols=columnasAUT)
        dfau1 = dfau1.fillna('-')
        dfau1 = dfau1.rename(
            columns={'FECHA INGRESO': 'Fecha_ingreso', 'FREC / CANAL': 'freq1', 'CIUDAD PRINCIPAL COBERTURA': 'ciu',
                     'No. OFICIO ARCOTEL': 'Oficio', 'NOMBRE ESTACIÓN': 'est',
                     'FECHA INICIO SUSPENSION': 'Fecha_inicio',
                     'FECHA OFICIO': 'Fecha_oficio', 'DIAS': 'Plazo'})
        dfau1['Tipo'] = pd.Series(['S' for x in range(len(dfau1.index))])
        dfau1 = dfau1[dfau1.Oficio != '-']
        dfau1 = dfau1[dfau1.Fecha_inicio != '-']
        dfau1['Fecha_ingreso'] = pd.to_datetime(dfau1['Fecha_ingreso'])
        dfau1['Fecha_oficio'] = pd.to_datetime(dfau1['Fecha_oficio'])
        dfau1['Fecha_inicio'] = pd.to_datetime(dfau1['Fecha_inicio'])
        dfau1['Fecha_fin'] = dfau1['Fecha_inicio'] + pd.to_timedelta(dfau1['Plazo'] - 1, unit='d')
        dfau1['freq1'] = dfau1['freq1'].replace('-', np.nan)
        dfau1['freq1'] = pd.to_numeric(dfau1['freq1'])

        def freq(row):
            """function to modify the values in freq1 column to present all in Hz, except if is a TV channel number"""
            if row['freq1'] >= 570 and row['freq1'] <= 1590:
                return row['freq1'] * 1000
            elif row['freq1'] >= 88 and row['freq1'] <= 108:
                return row['freq1'] * 1000000
            else:
                return row['freq1']

        # Create a new column in the dfau1 dataframe by using the last function def freq(row)
        dfau1['freq'] = dfau1.apply(lambda row: freq(row), axis=1)
        dfau1 = dfau1.drop(columns=['freq1'])

        # DATA CLEANING
        # convert column "Tiempo" to datetime in an especific format
        df_original1["Tiempo"] = pd.to_datetime(df_original1["Tiempo"], format='%d/%m/%Y %H:%M:%S.%f')
        df_original2["Tiempo"] = pd.to_datetime(df_original2["Tiempo"], format='%d/%m/%Y %H:%M:%S.%f')

        # DATA ANALYSIS
        df1 = df_original1
        df2 = df_original2

        # DATA READING: For the locations where AM Broadcasting can be measured.
        if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
            df_d3 = []
            for mes in month_year[m:n + 1]:
                # w: read the csv file, used usecols, to list the data and pass it to a numpy array
                try:
                    w = pd.read_csv(f'{server_route}/{ciu}/AM_{ciu}_{mes}.csv', engine='python',
                                    skipinitialspace=True, usecols=columnasAM, encoding='unicode_escape').to_numpy()
                except IOError:
                    # raise if file does not exist
                    w = np.full([1, 6], np.nan)

                # df_d3: append(w) adds all the elements in the w lists generated by the for loop
                df_d3.append(w)

            # Join the sequence of numpy arrays in the previous df to get only one array of data
            df_d3 = np.concatenate(df_d3)
            # df_original3: convert df_d3 to pandas dataframe and add header
            df_original3 = pd.DataFrame(df_d3,
                                        columns=['Tiempo', 'Frecuencia (Hz)', 'Level (dBµV/m)', 'Offset (Hz)', 'AM (%)',
                                                 'Bandwidth (Hz)'])

            # df13: read the TX.xlsx data, convert it to a pandas dataframe and fill na with -
            df13 = pd.read_excel(f'{server_route}/{file_estaciones}', sheet_name=sheet_name3)
            df13 = df13.fillna('-')

            # DATA CLEANING
            # convert column "Tiempo" to datetime in an especific format
            df_original3["Tiempo"] = pd.to_datetime(df_original3["Tiempo"], format='%d/%m/%Y %H:%M:%S.%f')
            df14 = df_original3

        # Add HH:MM:SS to fecha_inicio and fecha_fin so the range of dates we want to show in the report is correct
        # (reminder: the dates we enter in the form is just a string with format 2022-01-12, that is why we made this
        # change to not lose information when we present the data at the end)
        add_string1 = ' 00:00:01'
        add_string2 = ' 23:59:59'
        fecha_inicio += add_string1
        fecha_fin += add_string2

        def convert(date_time):
            """function to convert a string to datetime object"""
            format = '%Y-%m-%d %H:%M:%S'  # The format
            datetime_str = datetime.datetime.strptime(date_time, format)
            return datetime_str

        # Convert fecha_inicio and fecha_fin to datetime object
        fecha_inicio = convert(fecha_inicio)
        fecha_fin = convert(fecha_fin)

        # Create an empty dataframe with columns "Tiempo" and "Frecuencia (Hz)" for the dates of the month in
        # evaluation
        df3 = []
        for t in pd.date_range(start=fecha_inicio, end=fecha_fin):
            for f in df7['Frecuencia (Hz)'].tolist():
                df3.append((t, f))
        df3 = pd.DataFrame(df3, columns=('Tiempo', 'Frecuencia (Hz)'))

        df4 = []
        for t in pd.date_range(start=fecha_inicio, end=fecha_fin):
            for f in df8['Frecuencia (Hz)'].tolist():
                df4.append((t, f))
        df4 = pd.DataFrame(df4, columns=('Tiempo', 'Frecuencia (Hz)'))

        # Concatenate real dataframe with empty dataframe to fill missing dates
        df5 = pd.concat([df3, df1])

        df6 = pd.concat([df4, df2])

        # Merge df5 with TX dataframe to fill TX by frequency, df9 is the dataframe that contains all the information
        # for FM broadcasting
        df9 = df5.merge(df7, how='right', on='Frecuencia (Hz)')
        # Select data between fecha_inicio and fecha_fin and fill missing data with 0
        df9 = df9[(df9.Tiempo >= fecha_inicio) & (df9.Tiempo <= fecha_fin)]
        df9 = df9.fillna(0)

        # Merge df6 with TX dataframe to fill TX by frequency, df10 is the dataframe that contains all the information
        # for TV broadcasting
        df10 = df6.merge(df8, how='right', on='Frecuencia (Hz)')
        # Select data between fecha_inicio and fecha_fin and fill missing data with 0
        df10 = df10[(df10.Tiempo >= fecha_inicio) & (df10.Tiempo <= fecha_fin)]
        df10 = df10.fillna(0)

        if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
            # For the AM broadcasting case the procedure from line 616 to 647 is recreated
            df15 = []
            for t in pd.date_range(start=fecha_inicio, end=fecha_fin):
                for f in df13['Frecuencia (Hz)'].tolist():
                    df15.append((t, f))
            df15 = pd.DataFrame(df15, columns=('Tiempo', 'Frecuencia (Hz)'))
            df16 = pd.concat([df15, df14])

            # Merge df16 with TX dataframe to fill TX by frequency, df17 is the dataframe that contains all the
            # information for AM broadcasting
            df17 = df16.merge(df13, how='right', on='Frecuencia (Hz)')
            df17 = df17[(df17.Tiempo >= fecha_inicio) & (df17.Tiempo <= fecha_fin)]
            df17 = df17.fillna(0)

        if Ocupacion == False and AM_Reporte_individual == False and FM_Reporte_individual == False and TV_Reporte_individual == False and Autorizaciones == False:
            # REPORTE GENERAL
            # Group the information in df9 according to the requirements for the final report: max Level and average
            # Bandwidth per Frequency and Day (FM broadcasting)
            df11 = df9.groupby(by=[
                pd.Grouper(key='Tiempo', freq='D'),
                pd.Grouper(key='Frecuencia (Hz)'),
                pd.Grouper(key='Estación'),
                pd.Grouper(key='Potencia'),
                pd.Grouper(key='BW Asignado')
            ]).agg({
                'Level (dBµV/m)': 'max',
                'Bandwidth (Hz)': 'mean'
            }).reset_index()

            # Group the information in df10 according to the requirements for the final report: max Level per Frequency
            # and Day (TV broadcasting)
            df12 = df10.groupby(by=[
                pd.Grouper(key='Tiempo', freq='D'),
                pd.Grouper(key='Frecuencia (Hz)'),
                pd.Grouper(key='Estación'),
                pd.Grouper(key='Canal (Número)'),
                pd.Grouper(key='Analógico/Digital'),
            ]).agg({
                'Level (dBµV/m)': 'max'
            }).reset_index()

            # Make the pivot tables with the data structured in the way we want to show in the report
            # (FM broadcasting)
            df_final1 = pd.pivot_table(df11,
                                       index=[pd.Grouper(key='Tiempo')],
                                       values=['Level (dBµV/m)', 'Bandwidth (Hz)'],
                                       columns=['Frecuencia (Hz)', 'Estación', 'Potencia', 'BW Asignado'],
                                       aggfunc={'Level (dBµV/m)': max, 'Bandwidth (Hz)': np.average}).round(2)
            df_final1 = df_final1.T
            df_final3 = df_final1.replace(0, '-')
            # Reset the index (unstack) to have the columns 'Potencia', 'BW Asignado' and 'Tiempo' in the index so we
            # can use the flags in the columns 'Potencia' and 'BW Asignado'
            df_final3 = df_final3.reset_index()

            # Sorter first by 'Level (dBµV/m)' and after by 'Bandwidth (Hz)' and rename the column header as
            # 'Param'
            sorter = ['Level (dBµV/m)', 'Bandwidth (Hz)']
            df_final3.level_0 = df_final3.level_0.astype("category")
            df_final3.level_0 = df_final3.level_0.cat.set_categories(sorter)
            df_final3 = df_final3.sort_values(['level_0', 'Frecuencia (Hz)'])
            df_final5 = df_final3.rename(columns={'level_0': 'Param'})

            if Year1 == Year2 and Mes_inicio == Mes_fin:
                # If evaluation period is just one month get the average of the values in another column named
                # 'Promedio' and create a last column named 'Observaciones'
                df_final5['Promedio'] = df_final5.drop(
                    ['Param', 'Frecuencia (Hz)', 'Estación', 'Potencia', 'BW Asignado'],
                    axis=1).replace('-', np.NaN).apply(lambda x: x.mean(), axis=1).round(
                    2)
                df_final5['Observaciones'] = ''
            else:
                # Only create a last column named 'Observaciones'
                df_final5['Observaciones'] = ''
                df_final5
            df_final5 = df_final5.rename(columns={'Param': 'Parámetro'}).set_index('Parámetro')

            # Make the pivot tables with the data structured in the way we want to show in the report
            # (TV broadcasting)
            df_final2 = pd.pivot_table(df12,
                                       index=[pd.Grouper(key='Tiempo')],
                                       values=['Level (dBµV/m)'],
                                       columns=['Frecuencia (Hz)', 'Estación', 'Canal (Número)', 'Analógico/Digital'],
                                       aggfunc={'Level (dBµV/m)': max}).round(2)
            df_final2 = df_final2.T
            df_final4 = df_final2.replace(0, '-')
            df_final4 = df_final4.reset_index()

            # Sorter first by 'Level (dBµV/m)' and rename the column header as 'Param'
            sorter1 = ['Level (dBµV/m)']
            df_final4.level_0 = df_final4.level_0.astype("category")
            df_final4.level_0 = df_final4.level_0.cat.set_categories(sorter1)
            df_final4 = df_final4.sort_values(['level_0', 'Frecuencia (Hz)'])
            df_final6 = df_final4.rename(columns={'level_0': 'Param'})

            if Year1 == Year2 and Mes_inicio == Mes_fin:
                # If evaluation period is just one month get the average of the values in another column named
                # 'Promedio' and create a last column named 'Observaciones'
                df_final6['Promedio'] = df_final6.drop(
                    ['Param', 'Frecuencia (Hz)', 'Estación', 'Canal (Número)', 'Analógico/Digital'], axis=1).replace(
                    '-',
                    np.NaN).apply(
                    lambda x: x.mean(), axis=1).round(2)
                df_final6['Observaciones'] = ''
            else:
                # Only create a last column named 'Observaciones'
                df_final6['Observaciones'] = ''
                df_final6
            df_final6 = df_final6.rename(columns={'Param': 'Parámetro'}).set_index('Parámetro')

            if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                # Group the information in df17 according to the requirements for the final report: max Level and
                # average Bandwidth per Frequency and Day (AM broadcasting)
                df18 = df17.groupby(by=[
                    pd.Grouper(key='Tiempo', freq='D'),
                    pd.Grouper(key='Frecuencia (Hz)'),
                    pd.Grouper(key='Estación')
                ]).agg({
                    'Level (dBµV/m)': 'max',
                    'Bandwidth (Hz)': 'mean'
                }).reset_index()

                # Make the pivot tables with the data structured in the way we want to show in the report
                # (AM broadcasting)
                df_final7 = pd.pivot_table(df18,
                                           index=[pd.Grouper(key='Tiempo')],
                                           values=['Level (dBµV/m)', 'Bandwidth (Hz)'],
                                           columns=['Frecuencia (Hz)', 'Estación'],
                                           aggfunc={'Level (dBµV/m)': max, 'Bandwidth (Hz)': np.average}).round(2)
                df_final7 = df_final7.T
                df_final8 = df_final7.replace(0, '-')
                # Reset the index (unstack) to have the columns 'Potencia', 'BW Asignado' and 'Tiempo' in the index so
                # we can use the flags in the columns 'Potencia' and 'BW Asignado'
                df_final8 = df_final8.reset_index()

                # Sorter first by 'Level (dBµV/m)' and after by 'Bandwidth (Hz)' and rename the column header as
                # 'Param'
                sorter = ['Level (dBµV/m)', 'Bandwidth (Hz)']
                df_final8.level_0 = df_final8.level_0.astype("category")
                df_final8.level_0 = df_final8.level_0.cat.set_categories(sorter)
                df_final8 = df_final8.sort_values(['level_0', 'Frecuencia (Hz)'])
                df_final9 = df_final8.rename(columns={'level_0': 'Param'})

                if Year1 == Year2 and Mes_inicio == Mes_fin:
                    # If evaluation period is just one month get the average of the values in another column named
                    # 'Promedio' and create a last column named 'Observaciones'
                    df_final9['Promedio'] = df_final9.drop(['Param', 'Frecuencia (Hz)', 'Estación'], axis=1).replace(
                        '-',
                        np.NaN).apply(
                        lambda x: x.mean(), axis=1).round(2)
                    df_final9['Observaciones'] = ''
                else:
                    # Only create a last column named 'Observaciones'
                    df_final9['Observaciones'] = ''
                    df_final9
                df_final9 = df_final9.rename(columns={'Param': 'Parámetro'}).set_index('Parámetro')

            # EXCEL REPORT CREATION: REPORTE GENERAL
            # create, write and save
            with pd.ExcelWriter(f'{download_route}/RTV_Verificación de parámetros.xlsx') as writer:
                df_final5.to_excel(writer, sheet_name='Radiodifusión FM')
                df_final6.to_excel(writer, sheet_name='Televisión')
                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    df_final9.to_excel(writer, sheet_name='Radiodifusión AM')

                if Year1 == Year2 and Mes_inicio == Mes_fin:
                    df_original1.to_excel(writer, sheet_name='Mediciones FM')
                    df_original2.to_excel(writer, sheet_name='Mediciones TV')
                    worksheet2 = writer.sheets['Mediciones FM']
                    worksheet3 = writer.sheets['Mediciones TV']
                    if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                        df_original3.to_excel(writer, sheet_name='Mediciones AM')
                        worksheet5 = writer.sheets['Mediciones AM']

                # Get the xlsxwriter workbook and worksheet objects.
                workbook = writer.book
                worksheet = writer.sheets['Radiodifusión FM']
                worksheet1 = writer.sheets['Televisión']
                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    worksheet4 = writer.sheets['Radiodifusión AM']

                # Add a format.
                format1 = workbook.add_format({'bg_color': '#C6EFCE',
                                               'font_color': '#006100'})
                format2 = workbook.add_format({'bg_color': '#FFC7CE',
                                               'font_color': '#9C0006'})
                format3 = workbook.add_format({'bg_color': '#FFDC47',
                                               'font_color': '#9C6500'})
                format4 = workbook.add_format({'num_format': 'dd/mm/yy'})

                format5 = workbook.add_format({'border': 1, 'border_color': 'black'})

                format6 = workbook.add_format({'bg_color': '#99CCFF',
                                               'font_color': '#0066FF'})

                # Get the dimensions of the dataframe (FM broadcasting).
                (max_row, max_col) = df_final5.drop(['Observaciones'], axis=1).shape

                # Apply a conditional format to the required cell range.
                worksheet.conditional_format(0, 5, 0, int(max_col + 1),
                                             {'type': 'no_errors',
                                              'format': format4})
                worksheet.conditional_format(0, 0, int(max_row), int(max_col + 1),
                                             {'type': 'no_errors',
                                              'format': format5})
                worksheet.autofilter(0, 0, 0, int(max_col + 1))

                a = df_final3.groupby('level_0')['level_0'].count()
                max_row = a[0]

                r = 1
                for r in range(1, int(max_row + 1)):
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(F2="-")',
                                                  'format': format6})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(F2<30)',
                                                  'format': format2})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="BAJA",F2>=43)',
                                                  'format': format1})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="BAJA",F2>=30,F2<43)',
                                                  'format': format3})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2="-",F2>=54)',
                                                  'format': format1})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2="-",F2>=30,F2<54)',
                                                  'format': format3})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2=200,F2>=54)',
                                                  'format': format1})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2=200,F2>=30,F2<54)',
                                                  'format': format3})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2=180,F2>=48)',
                                                  'format': format1})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2=180,F2>=30,F2<48)',
                                                  'format': format3})

                s = int(max_row + 1)
                for s in range(int(max_row + 1), int((2 * max_row) + 1)):
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(F104="-")',
                                                  'format': format6})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104=180,F104<=180000)',
                                                  'format': format1})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104=180,F104>180000)',
                                                  'format': format2})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104=200,F104<=200000)',
                                                  'format': format1})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104=200,F104>200000)',
                                                  'format': format2})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104="-",F104<=220000)',
                                                  'format': format1})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104="-",F104>220000)',
                                                  'format': format2})

                worksheet.conditional_format('A208', {'type': 'no_errors',
                                                      'format': format1})
                worksheet.conditional_format('A209', {'type': 'no_errors',
                                                      'format': format3})
                worksheet.conditional_format('A210', {'type': 'no_errors',
                                                      'format': format2})
                worksheet.conditional_format('A211', {'type': 'no_errors',
                                                      'format': format6})
                worksheet.write('B207',
                                '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                worksheet.write('B208',
                                '- Color VERDE: los valores de campo eléctrico diario superan el valor del borde de área de cobertura (Estereofónico: >=54 dBuV/m, Monofónico: >=48 dBuV/m o Baja Potencia: >=43 dBuV/m).')
                worksheet.write('B209',
                                '- Color AMARILLO: los valores de campo eléctrico diario se encuentran entre el valor del borde de área de protección y el valor del borde de área de cobertura (Estereofónico: entre 30 y 54 dBuV/m, Monofónico: entre 30 y 48 dBuV/m o Baja Potencia: entre 30 y 43 dBuV/m).')
                worksheet.write('B210',
                                '- Color ROJO: los valores de campo eléctrico diario son inferiores al valor del borde de área de protección (<30 dBuV/m).')
                worksheet.write('B211', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                worksheet.write('B212',
                                '- Para todos los casos el valor de ancho de banda corresponde a 220 kHz, excepto aquellos en los que se especifica que el valor es de 180 kHz o 200 kHz.')
                worksheet.write('A214', 'CONSIDERACIONES GENERALES:')
                worksheet.write('B215',
                                f'- Nota 1.- En apego a la Resolución ST-2014-0257 referida en los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), se considera que un control individual a una estación RTV es pertinente cuando esta ha suspendido emisiones por más de 8 días.')
                worksheet.write('B216',
                                '- Nota 2.- Las mediciones de Ancho de Banda obtenidas con el SACER son únicamente referenciales. La Unión Internacional de Telecomunicaciones – UIT, a través del Manual COMPROBACIÓN TÉCNICA DEL ESPECTRO RADIOELÉCTRICO (numeral 4.5.3) define las condiciones que deben tenerse en cuenta al medir la anchura de banda, señalando: “(…) a causa de las imprecisiones debidas a las razones expuestas, estas mediciones a distancia sólo son útiles a título indicativo. Cuando se requiera una mayor precisión, es conveniente que las mediciones se realicen en las inmediaciones del transmisor.”.')
                worksheet.write('B217',
                                f'- Nota 3.- De acuerdo a los numerales 5.2.4 y 5.2.5 de los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), para cada observación detectada que contenga datos inconsistentes que no se encuentren dentro del rango autorizado, dependiendo del caso y de los resultados, corresponde programar una verificación en sitio. Dicha programación puede estar sujeta a los resultados de las mediciones del siguiente mes.')

                # Get the dimensions of the dataframe (TV broadcasting).
                (max_row1, max_col1) = df_final6.drop(['Observaciones'], axis=1).shape

                # Apply a conditional format to the required cell range.
                worksheet1.conditional_format(0, 5, 0, int(max_col1 + 1),
                                              {'type': 'no_errors',
                                               'format': format4})
                worksheet1.conditional_format(0, 0, int(max_row1), int(max_col1 + 1),
                                              {'type': 'no_errors',
                                               'format': format5})
                worksheet1.autofilter(0, 0, 0, int(max_col1 + 1))

                b = df_final4.groupby('level_0')['level_0'].count()
                max_row1 = b[0]

                i = 1
                for i in range(1, int(max_row1 + 1)):
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND(F2="-")',
                                                   'format': format6})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="D",F2>=51)',
                                                   'format': format1})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="D",F2<51)',
                                                   'format': format2})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=54000000,$B2<=88000000,F2>=68)',
                                                   'format': format1})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=54000000,$B2<=88000000,F2>=47,F2<68)',
                                                   'format': format3})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=54000000,$B2<=88000000,F2<47)',
                                                   'format': format2})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=174000000,$B2<=216000000,F2>=71)',
                                                   'format': format1})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=174000000,$B2<=216000000,F2>=56,F2<71)',
                                                   'format': format3})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=174000000,$B2<=216000000,F2<56)',
                                                   'format': format2})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=470000000,$B2<=880000000,F2>=74)',
                                                   'format': format1})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=470000000,$B2<=880000000,F2>=64,F2<74)',
                                                   'format': format3})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=470000000,$B2<=880000000,F2<64)',
                                                   'format': format2})

                worksheet1.conditional_format('A49', {'type': 'no_errors',
                                                      'format': format1})
                worksheet1.conditional_format('A50', {'type': 'no_errors',
                                                      'format': format3})
                worksheet1.conditional_format('A51', {'type': 'no_errors',
                                                      'format': format2})
                worksheet1.conditional_format('A52', {'type': 'no_errors',
                                                      'format': format6})
                worksheet1.write('B48',
                                 '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                worksheet1.write('B49',
                                 '- Color VERDE significa que los valores de campo eléctrico diario superan el límite del área de cobertura primaria. (Canales 2 al 6: >=68 dBuV/m, Canales 7 al 13: >=71 dBuV/m, Canales 14 al 51: >=74 dBuV/m, Canales digitales: >=51 dBuV/m).')
                worksheet1.write('B50',
                                 '- Color AMARILLO significa que los valores de campo eléctrico diario superan el límite del área de cobertura secundario pero son inferiores al límite del área de cobertura principal establecidos para cada rango. (Canales 2 al 6: entre 47 y 68 dBuV/m, Canales 7 al 13: entre 56 y 71 dBuV/m, Canales 14 al 51: entre 64 y 74 dBuV/m.')
                worksheet1.write('B51',
                                 '- Color ROJO significa que los valores de campo eléctrico diario son inferiores al límite de área de cobertura secundario. (Canales 2 al 6: <47 dBuV/m, Canales 7 al 13: <56 dBuV/m, Canales 14 al 51: <64 dBuV/m, Canales digitales: <51 dBuV/m).')
                worksheet1.write('B52', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                worksheet1.write('A54', 'CONSIDERACIONES GENERALES:')
                worksheet1.write('B55',
                                 f'- Nota 1.- En apego a la Resolución ST-2014-0257 referida en los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), se considera que un control individual a una estación RTV es pertinente cuando esta ha suspendido emisiones por más de 8 días.')
                worksheet1.write('B56',
                                 f'- Nota 2.- De acuerdo a los numerales 5.2.4 y 5.2.5 de los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), para cada observación detectada que contenga datos inconsistentes que no se encuentren dentro del rango autorizado, dependiendo del caso y de los resultados, corresponde programar una verificación en sitio. Dicha programación puede estar sujeta a los resultados de las mediciones del siguiente mes.')

                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    # Get the dimensions of the dataframe (AM broadcasting).
                    (max_row4, max_col4) = df_final9.drop(['Observaciones'], axis=1).shape

                    # Apply a conditional format to the required cell range.
                    worksheet4.conditional_format(0, 3, 0, int(max_col4 + 1),
                                                  {'type': 'no_errors',
                                                   'format': format4})
                    worksheet4.conditional_format(0, 0, int(max_row4), int(max_col4 + 1),
                                                  {'type': 'no_errors',
                                                   'format': format5})
                    worksheet4.autofilter(0, 0, 0, int(max_col4 + 1))

                    c = df_final8.groupby('level_0')['level_0'].count()
                    max_row4 = c[0]

                    r = 1
                    for r in range(1, int(max_row4 + 1)):
                        worksheet4.conditional_format(r, 3, int(max_row4), int(max_col4),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D2="-")',
                                                       'format': format6})
                        worksheet4.conditional_format(r, 3, int(max_row4), int(max_col4),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D2<62)',
                                                       'format': format2})
                        worksheet4.conditional_format(r, 3, int(max_row4), int(max_col4),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D2>=62)',
                                                       'format': format1})

                    s = int(max_row4 + 1)
                    for s in range(int(max_row4 + 1), int((2 * max_row4) + 1)):
                        worksheet4.conditional_format(s, 3, int(2 * max_row4), int(max_col4),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D45="-")',
                                                       'format': format6})
                        worksheet4.conditional_format(s, 3, int(2 * max_row4), int(max_col4),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D45<=15000)',
                                                       'format': format1})
                        worksheet4.conditional_format(s, 3, int(2 * max_row4), int(max_col4),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D45>15000)',
                                                       'format': format2})

                    worksheet4.conditional_format('A90', {'type': 'no_errors',
                                                          'format': format1})
                    worksheet4.conditional_format('A91', {'type': 'no_errors',
                                                          'format': format2})
                    worksheet4.conditional_format('A92', {'type': 'no_errors',
                                                          'format': format6})
                    worksheet4.write('B89',
                                     '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet4.write('B90',
                                     '- Color VERDE: los valores de campo eléctrico diario superan el valor del borde de área de cobertura (>=62 dBuV/m).')
                    worksheet4.write('B91',
                                     '- Color ROJO: los valores de campo eléctrico diario son inferiores al valor del borde de área de protección (<62 dBuV/m).')
                    worksheet4.write('B92', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                    worksheet4.write('B93', '- Para todos los casos el valor de ancho de banda corresponde a 15 kHz.')
                    worksheet4.write('A95', 'CONSIDERACIONES GENERALES:')
                    worksheet4.write('B96',
                                     f'- Nota 1.- En apego a la Resolución ST-2014-0257 referida en los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), se considera que un control individual a una estación RTV es pertinente cuando esta ha suspendido emisiones por más de 8 días.')
                    worksheet4.write('B97',
                                     '- Nota 2.- Las mediciones de Ancho de Banda obtenidas con el SACER son únicamente referenciales. La Unión Internacional de Telecomunicaciones – UIT, a través del Manual COMPROBACIÓN TÉCNICA DEL ESPECTRO RADIOELÉCTRICO (numeral 4.5.3) define las condiciones que deben tenerse en cuenta al medir la anchura de banda, señalando: “(…) a causa de las imprecisiones debidas a las razones expuestas, estas mediciones a distancia sólo son útiles a título indicativo. Cuando se requiera una mayor precisión, es conveniente que las mediciones se realicen en las inmediaciones del transmisor.”.')
                    worksheet4.write('B98',
                                     f'- Nota 3.- De acuerdo a los numerales 5.2.4 y 5.2.5 de los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), para cada observación detectada que contenga datos inconsistentes que no se encuentren dentro del rango autorizado, dependiendo del caso y de los resultados, corresponde programar una verificación en sitio. Dicha programación puede estar sujeta a los resultados de las mediciones del siguiente mes.')

                    if Year1 == Year2 and Mes_inicio == Mes_fin:
                        # Get the dimensions of the dataframe (FM broadcasting).
                        (max_row2, max_col2) = df_original1.shape

                        # Apply a conditional format to the required cell range.
                        worksheet2.conditional_format(1, 1, int(max_row2), int(max_col2),
                                                      {'type': 'no_errors',
                                                       'format': format5})
                        worksheet2.autofilter(0, 0, 0, int(max_col2))

                        # Get the dimensions of the dataframe (TV broadcasting).
                        (max_row3, max_col3) = df_original2.shape

                        # Apply a conditional format to the required cell range.
                        worksheet3.conditional_format(1, 1, int(max_row3), int(max_col3),
                                                      {'type': 'no_errors',
                                                       'format': format5})
                        worksheet3.autofilter(0, 0, 0, int(max_col3))

                        if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                            # Get the dimensions of the dataframe (AM broadcasting).
                            (max_row5, max_col5) = df_original3.shape

                            # Apply a conditional format to the required cell range.
                            worksheet5.conditional_format(1, 1, int(max_row5), int(max_col5),
                                                          {'type': 'no_errors',
                                                           'format': format5})
                            worksheet5.autofilter(0, 0, 0, int(max_col5))

            # Change the name of the file
            old_name = 'RTV_Verificación de parámetros.xlsx'
            if Year1 == Year2 and Mes_inicio == Mes_fin:
                new_name = 'RTV_Verificación de parámetros_{}_{}_{}.xlsx'.format(Ciudad, Mes_inicio, Year1)
            else:
                new_name = 'RTV_Verificación de parámetros_{}_{}{}_{}{}.xlsx'.format(Ciudad, Mes_inicio, Year1, Mes_fin,
                                                                                     Year2)

            # Remove the previous file if already exist
            if os.path.exists(f'{download_route}/{new_name}'):
                os.remove(f'{download_route}/{new_name}')

            # Rename the file
            os.rename(f'{download_route}/{old_name}',
                      f'{download_route}/{new_name}')

        elif Ocupacion == False and AM_Reporte_individual == False and FM_Reporte_individual == False and TV_Reporte_individual == False and Autorizaciones == True:
            # REPORTE GENERAL AUTORIZACIONES
            # Filter dfau2 to get FM frequencies per City
            dfau2 = dfau1[(dfau1.freq > 87700000) & (dfau1.freq < 108100000)]
            dfau2 = dfau2.rename(columns={'freq': 'Frecuencia (Hz)', 'Fecha_inicio': 'Tiempo'})
            dfau2 = dfau2.loc[dfau2['ciu'] == autori]
            dfau6 = dfau2
            dfau2 = dfau2.drop(columns=['est'])

            # Filter dfau3 to get TV channels per City
            dfau3 = dfau1[(dfau1.freq >= 2) & (dfau1.freq <= 51)]
            dfau3 = dfau3.rename(columns={'freq': 'Canal (Número)', 'Fecha_inicio': 'Tiempo'})
            dfau3 = dfau3.loc[dfau3['ciu'] == autori]
            dfau7 = dfau3
            dfau3 = dfau3.drop(columns=['est'])

            # Filter dfau8 to get AM frequencies per City
            dfau8 = dfau1[(dfau1.freq >= 570000) & (dfau1.freq <= 1590000)]
            dfau8 = dfau8.rename(columns={'freq': 'Frecuencia (Hz)', 'Fecha_inicio': 'Tiempo'})
            dfau8 = dfau8.loc[dfau8['ciu'] == autori]
            dfau10 = dfau8
            dfau8 = dfau8.drop(columns=['est'])

            # Group the information in df9 according to the requirements for the final report: max Level and average
            # Bandwidth per Frequency and Day (FM broadcasting)
            df11 = df9.groupby(by=[
                pd.Grouper(key='Tiempo', freq='D'),
                pd.Grouper(key='Frecuencia (Hz)'),
                pd.Grouper(key='Estación'),
                pd.Grouper(key='Potencia'),
                pd.Grouper(key='BW Asignado')
            ]).agg({
                'Level (dBµV/m)': 'max',
                'Bandwidth (Hz)': 'mean'
            }).reset_index()

            # Group the information in df10 according to the requirements for the final report: max Level per Frequency
            # and Day (TV broadcasting)
            df12 = df10.groupby(by=[
                pd.Grouper(key='Tiempo', freq='D'),
                pd.Grouper(key='Frecuencia (Hz)'),
                pd.Grouper(key='Estación'),
                pd.Grouper(key='Canal (Número)'),
                pd.Grouper(key='Analógico/Digital'),
            ]).agg({
                'Level (dBµV/m)': 'max'
            }).reset_index()

            # Merge dfau2 with df11 dataframe to add the autorization df11 is the dataframe that contains all the
            # information for FM
            dfau2 = dfau2.rename(columns={'Fecha_inicio': 'Tiempo'})
            dfau2 = dfau2.drop(columns=['ciu'])
            dfau4 = []
            for index, row in dfau2.iterrows():
                for t in pd.date_range(start=row['Tiempo'], end=row['Fecha_fin']):
                    dfau4.append(
                        (row['Frecuencia (Hz)'], row['Tipo'], row['Plazo'], t, row['Oficio'], row['Fecha_oficio'],
                         row['Fecha_fin']))
            dfau4 = pd.DataFrame(dfau4,
                                 columns=(
                                     'Frecuencia (Hz)', 'Tipo', 'Plazo', 'Tiempo', 'Oficio', 'Fecha_oficio',
                                     'Fecha_fin'))
            df11 = dfau4.merge(df11, how='right', on=['Tiempo', 'Frecuencia (Hz)'])
            df11 = df11.fillna('-')

            # Merge dfau3 with df12 dataframe to add the autorization df10 is the dataframe that contains all the
            # information for TV
            dfau3 = dfau3.rename(columns={'Fecha_inicio': 'Tiempo'})
            dfau3 = dfau3.drop(columns=['ciu'])
            dfau5 = []
            for index, row in dfau3.iterrows():
                for t in pd.date_range(start=row['Tiempo'], end=row['Fecha_fin']):
                    dfau5.append(
                        (row['Canal (Número)'], row['Tipo'], row['Plazo'], t, row['Oficio'], row['Fecha_oficio'],
                         row['Fecha_fin']))
            dfau5 = pd.DataFrame(dfau5,
                                 columns=(
                                     'Canal (Número)', 'Tipo', 'Plazo', 'Tiempo', 'Oficio', 'Fecha_oficio',
                                     'Fecha_fin'))
            df12 = dfau5.merge(df12, how='right', on=['Tiempo', 'Canal (Número)'])
            df12 = df12.fillna('-')

            # Make the pivot tables with the data structured in the way we want to show in the report
            # (FM broadcasting)
            df_final1 = pd.pivot_table(df11,
                                       index=[pd.Grouper(key='Tiempo')],
                                       values=['Level (dBµV/m)', 'Bandwidth (Hz)', 'Fecha_fin'],
                                       columns=['Frecuencia (Hz)', 'Estación', 'Potencia', 'BW Asignado'],
                                       aggfunc={'Level (dBµV/m)': max, 'Bandwidth (Hz)': np.average,
                                                'Fecha_fin': max}).round(2)
            df_final1 = df_final1.rename(columns={'Fecha_fin': 'Fin de Autorización'})
            df_final1 = df_final1.T
            df_final3 = df_final1.replace(0, '-')
            """Reset the index (unstack) to have the columns 'Potencia', 'BW Asignado' and 'Tiempo' in the index so we
            can use the flags in the columns 'Potencia' and 'BW Asignado' """
            df_final3 = df_final3.reset_index()

            """sorter first by 'Level (dBµV/m)', after by 'Bandwidth (Hz)' and finally by 'Fin de Autorización' and
            rename the column header as 'Param' """
            sorter = ['Level (dBµV/m)', 'Bandwidth (Hz)', 'Fin de Autorización']
            df_final3.level_0 = df_final3.level_0.astype("category")
            df_final3.level_0 = df_final3.level_0.cat.set_categories(sorter)
            df_final3 = df_final3.sort_values(['level_0', 'Frecuencia (Hz)'])
            df_final5 = df_final3.rename(columns={'level_0': 'Param'})

            if Year1 == Year2 and Mes_inicio == Mes_fin:
                """If evaluation period is just one month get the average of the values in another column named 
                'Promedio' and create a last column named 'Observaciones' """
                df_final5['Promedio'] = df_final5[(df_final5.Param != 'Fin de Autorización')].drop(
                    ['Param', 'Frecuencia (Hz)', 'Estación', 'Potencia', 'BW Asignado'], axis=1).replace('-',
                                                                                                         np.NaN).apply(
                    lambda x: x.mean(), axis=1).round(2)
                df_final5['Observaciones'] = ''
            else:
                """Only create a last column named 'Observaciones' """
                df_final5['Observaciones'] = ''
                df_final5
            df_final5 = df_final5.rename(columns={'Param': 'Parámetro'}).set_index('Parámetro')

            """Make the pivot tables with the data structured in the way we want to show in the report 
            (TV broadcasting)"""
            df_final2 = pd.pivot_table(df12,
                                       index=[pd.Grouper(key='Tiempo')],
                                       values=['Level (dBµV/m)', 'Fecha_fin'],
                                       columns=['Frecuencia (Hz)', 'Estación', 'Canal (Número)', 'Analógico/Digital'],
                                       aggfunc={'Level (dBµV/m)': max, 'Fecha_fin': max}).round(2)
            df_final2 = df_final2.rename(columns={'Fecha_fin': 'Fin de Autorización'})
            df_final2 = df_final2.T
            df_final4 = df_final2.replace(0, '-')
            df_final4 = df_final4.reset_index()

            """sorter first by 'Level (dBµV/m)', after by 'Fin de Autorización' and rename the column header as 
            'Param' """
            sorter1 = ['Level (dBµV/m)', 'Fin de Autorización']
            df_final4.level_0 = df_final4.level_0.astype("category")
            df_final4.level_0 = df_final4.level_0.cat.set_categories(sorter1)
            df_final4 = df_final4.sort_values(['level_0', 'Frecuencia (Hz)'])
            df_final6 = df_final4.rename(columns={'level_0': 'Param'})

            if Year1 == Year2 and Mes_inicio == Mes_fin:
                """If evaluation period is just one month get the average of the values in another column named
                'Promedio' and create a last column named 'Observaciones' """
                df_final6['Promedio'] = df_final6[(df_final6.Param != 'Fin de Autorización')].drop(
                    ['Param', 'Frecuencia (Hz)', 'Estación', 'Canal (Número)', 'Analógico/Digital'], axis=1).replace(
                    '-',
                    np.NaN).apply(
                    lambda x: x.mean(), axis=1).round(2)
                df_final6['Observaciones'] = ''
            else:
                """Only create a last column named 'Observaciones' """
                df_final6['Observaciones'] = ''
                df_final6
            df_final6 = df_final6.rename(columns={'Param': 'Parámetro'}).set_index('Parámetro')

            if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                """Group the information in df17 according to the requirements for the final report: max Level and
                average Bandwidth per Frequency and Day (AM broadcasting)"""
                df18 = df17.groupby(by=[
                    pd.Grouper(key='Tiempo', freq='D'),
                    pd.Grouper(key='Frecuencia (Hz)'),
                    pd.Grouper(key='Estación')
                ]).agg({
                    'Level (dBµV/m)': 'max',
                    'Bandwidth (Hz)': 'mean'
                }).reset_index()

                """Merge dfau9 with df11 dataframe to add the autorization df18 is the dataframe that contains all the
                 information for AM."""
                dfau8 = dfau8.rename(columns={'Fecha_inicio': 'Tiempo'})
                dfau8 = dfau8.drop(columns=['ciu'])
                dfau9 = []
                for index, row in dfau8.iterrows():
                    for t in pd.date_range(start=row['Tiempo'], end=row['Fecha_fin']):
                        dfau9.append(
                            (row['Frecuencia (Hz)'], row['Tipo'], row['Plazo'], t, row['Oficio'], row['Fecha_oficio'],
                             row['Fecha_fin']))
                dfau9 = pd.DataFrame(dfau9, columns=(
                    'Frecuencia (Hz)', 'Tipo', 'Plazo', 'Tiempo', 'Oficio', 'Fecha_oficio', 'Fecha_fin'))
                df18 = dfau9.merge(df18, how='right', on=['Tiempo', 'Frecuencia (Hz)'])
                df18 = df18.fillna('-')

                """Make the pivot tables with the data structured in the way we want to show in the report
                (AM broadcasting)"""
                df_final7 = pd.pivot_table(df18,
                                           index=[pd.Grouper(key='Tiempo')],
                                           values=['Level (dBµV/m)', 'Bandwidth (Hz)', 'Fecha_fin'],
                                           columns=['Frecuencia (Hz)', 'Estación'],
                                           aggfunc={'Level (dBµV/m)': max, 'Bandwidth (Hz)': np.average,
                                                    'Fecha_fin': max}).round(2)
                df_final7 = df_final7.rename(columns={'Fecha_fin': 'Fin de Autorización'})
                df_final7 = df_final7.T
                df_final8 = df_final7.replace(0, '-')
                """Reset the index (unstack) to have the columns 'Potencia', 'BW Asignado' and 'Tiempo' in the index so
                we can use the flags in the columns 'Potencia' and 'BW Asignado' """
                df_final8 = df_final8.reset_index()

                """sorter first by 'Level (dBµV/m)' and after by 'Bandwidth (Hz)' and rename the column header as
                'Param' """
                sorter = ['Level (dBµV/m)', 'Bandwidth (Hz)', 'Fin de Autorización']
                df_final8.level_0 = df_final8.level_0.astype("category")
                df_final8.level_0 = df_final8.level_0.cat.set_categories(sorter)
                df_final8 = df_final8.sort_values(['level_0', 'Frecuencia (Hz)'])
                df_final9 = df_final8.rename(columns={'level_0': 'Param'})

                if Year1 == Year2 and Mes_inicio == Mes_fin:
                    """If evaluation period is just one month get the average of the values in another column named
                    'Promedio' and create a last column named 'Observaciones' """
                    df_final9['Promedio'] = df_final9[(df_final9.Param != 'Fin de Autorización')].drop(
                        ['Param', 'Frecuencia (Hz)', 'Estación'], axis=1).replace('-', np.NaN).apply(lambda x: x.mean(),
                                                                                                     axis=1).round(2)
                    df_final9['Observaciones'] = ''
                else:
                    """Only create a last column named 'Observaciones' """
                    df_final9['Observaciones'] = ''
                    df_final9
                df_final9 = df_final9.rename(columns={'Param': 'Parámetro'}).set_index('Parámetro')

            """EXCEL REPORT CREATION: REPORTE GENERAL AUTORIZACIONES"""
            """create, write and save"""
            with pd.ExcelWriter(f'{download_route}/RTV_Verificación de parámetros.xlsx') as writer:
                df_final5.to_excel(writer, sheet_name='Radiodifusión FM')
                df_final6.to_excel(writer, sheet_name='Televisión')
                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    df_final9.to_excel(writer, sheet_name='Radiodifusión AM')

                """copy the column 'Tiempo' to a list and save it in col1"""
                col1 = dfau6['Tiempo'].copy().tolist()
                """insert a column in the index 13 with the name 'Fecha_ini' with the col1 information"""
                dfau6.insert(13, 'Fecha_ini', col1)
                dfau6 = dfau6.rename(
                    columns={'freq': 'Frecuencia (Hz)', 'Fecha_ini': 'Fecha_inicio'})
                dfau6.set_index('Frecuencia (Hz)').rename(
                    columns={'est': 'Estación', 'Fecha_inicio': 'Inicio Autorización',
                             'Fecha_fin': 'Fin Autorización'}).drop(
                    columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).to_excel(writer,
                                                                                         sheet_name='Autorizaciones FM')

                """copy the column 'Tiempo' to a list and save it in col1"""
                col2 = dfau7['Tiempo'].copy().tolist()
                """insert a column in the index 13 with the name 'Fecha_ini' with the col1 information"""
                dfau7.insert(13, 'Fecha_ini', col2)
                dfau7 = dfau7.rename(
                    columns={'freq': 'Canal (Número)', 'Fecha_ini': 'Fecha_inicio'})
                dfau7.set_index('Canal (Número)').rename(
                    columns={'est': 'Estación', 'Fecha_inicio': 'Inicio Autorización',
                             'Fecha_fin': 'Fin Autorización'}).drop(
                    columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).to_excel(writer,
                                                                                         sheet_name='Autorizaciones TV')

                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    """copy the column 'Tiempo' to a list and save it in col1"""
                    col3 = dfau10['Tiempo'].copy().tolist()
                    """insert a column in the index 13 with the name 'Fecha_ini' with the col1 information"""
                    dfau10.insert(13, 'Fecha_ini', col3)
                    dfau10 = dfau10.rename(
                        columns={'freq': 'Frecuencia (Hz)', 'Fecha_ini': 'Fecha_inicio'})
                    dfau10.set_index('Frecuencia (Hz)').rename(
                        columns={'est': 'Estación', 'Fecha_inicio': 'Inicio Autorización',
                                 'Fecha_fin': 'Fin Autorización'}).drop(
                        columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).to_excel(writer,
                                                                                             sheet_name='Autorizaciones AM')

                if Year1 == Year2 and Mes_inicio == Mes_fin:
                    df_original1.to_excel(writer, sheet_name='Mediciones FM')
                    df_original2.to_excel(writer, sheet_name='Mediciones TV')
                    worksheet2 = writer.sheets['Mediciones FM']
                    worksheet3 = writer.sheets['Mediciones TV']
                    if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                        df_original3.to_excel(writer, sheet_name='Mediciones AM')
                        worksheet8 = writer.sheets['Mediciones AM']

                """Get the xlsxwriter workbook and worksheet objects."""
                workbook = writer.book
                worksheet = writer.sheets['Radiodifusión FM']
                worksheet1 = writer.sheets['Televisión']
                worksheet4 = writer.sheets['Autorizaciones FM']
                worksheet5 = writer.sheets['Autorizaciones TV']
                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    worksheet6 = writer.sheets['Radiodifusión AM']
                    worksheet7 = writer.sheets['Autorizaciones AM']

                """Add a format."""
                format1 = workbook.add_format({'bg_color': '#C6EFCE',
                                               'font_color': '#006100'})
                format2 = workbook.add_format({'bg_color': '#FFC7CE',
                                               'font_color': '#9C0006'})
                format3 = workbook.add_format({'bg_color': '#FFDC47',
                                               'font_color': '#9C6500'})
                format4 = workbook.add_format({'num_format': 'dd/mm/yy'})

                format5 = workbook.add_format({'border': 1, 'border_color': 'black'})

                format6 = workbook.add_format({'bg_color': '#99CCFF',
                                               'font_color': '#0066FF'})
                format7 = workbook.add_format({'bg_color': '#C0C0C0',
                                               'font_color': '#000000'})

                """Get the dimensions of the dataframe (FM broadcasting)."""
                (max_row4, max_col4) = dfau6.drop(
                    columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).set_index(
                    'Frecuencia (Hz)').shape

                """Apply a conditional format to the required cell range."""
                worksheet4.conditional_format(1, 1, int(max_row4), int(max_col4),
                                              {'type': 'no_errors',
                                               'format': format5})
                worksheet4.autofilter(0, 0, 0, int(max_col4))

                """Get the dimensions of the dataframe (TV broadcasting)."""
                (max_row5, max_col5) = dfau7.drop(
                    columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).set_index(
                    'Canal (Número)').shape

                """Apply a conditional format to the required cell range."""
                worksheet5.conditional_format(1, 1, int(max_row5), int(max_col5),
                                              {'type': 'no_errors',
                                               'format': format5})
                worksheet5.autofilter(0, 0, 0, int(max_col5))

                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    """Get the dimensions of the dataframe (AM broadcasting)."""
                    (max_row6, max_col6) = dfau10.drop(
                        columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).set_index(
                        'Frecuencia (Hz)').shape

                    """Apply a conditional format to the required cell range."""
                    worksheet7.conditional_format(1, 1, int(max_row6), int(max_col6),
                                                  {'type': 'no_errors',
                                                   'format': format5})
                    worksheet7.autofilter(0, 0, 0, int(max_col6))

                """Get the dimensions of the dataframe (FM broadcasting)."""
                (max_row, max_col) = df_final5.drop(['Observaciones'], axis=1).shape

                """Apply a conditional format to the required cell range."""
                worksheet.conditional_format(0, 5, 0, int(max_col + 1),
                                             {'type': 'no_errors',
                                              'format': format4})
                worksheet.conditional_format(0, 0, int(max_row), int(max_col + 1),
                                             {'type': 'no_errors',
                                              'format': format5})
                worksheet.autofilter(0, 0, 0, int(max_col + 1))

                a = df_final3.groupby('level_0')['level_0'].count()
                max_row = a[0]

                r = 1
                for r in range(1, int(max_row + 1)):
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(F206<>"-",F206<>"")',
                                                  'format': format7})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(F2="-")',
                                                  'format': format6})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(F2<30)',
                                                  'format': format2})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="BAJA",F2>=43)',
                                                  'format': format1})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="BAJA",F2>=30,F2<43)',
                                                  'format': format3})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2="-",F2>=54)',
                                                  'format': format1})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2="-",F2>=30,F2<54)',
                                                  'format': format3})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2=200,F2>=54)',
                                                  'format': format1})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2=200,F2>=30,F2<54)',
                                                  'format': format3})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2=180,F2>=48)',
                                                  'format': format1})
                    worksheet.conditional_format(r, 5, int(max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($D2="-",$E2=180,F2>=30,F2<48)',
                                                  'format': format3})

                s = int(max_row + 1)
                for s in range(int(max_row + 1), int((2 * max_row) + 1)):
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(F206<>"-",F206<>"")',
                                                  'format': format7})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(F104="-")',
                                                  'format': format6})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104=180,F104<=180000)',
                                                  'format': format1})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104=180,F104>180000)',
                                                  'format': format2})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104=200,F104<=200000)',
                                                  'format': format1})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104=200,F104>200000)',
                                                  'format': format2})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104="-",F104<=220000)',
                                                  'format': format1})
                    worksheet.conditional_format(s, 5, int(2 * max_row), int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND($E104="-",F104>220000)',
                                                  'format': format2})

                worksheet.conditional_format('A310', {'type': 'no_errors',
                                                      'format': format1})
                worksheet.conditional_format('A311', {'type': 'no_errors',
                                                      'format': format3})
                worksheet.conditional_format('A312', {'type': 'no_errors',
                                                      'format': format2})
                worksheet.conditional_format('A313', {'type': 'no_errors',
                                                      'format': format6})
                worksheet.conditional_format('A314', {'type': 'no_errors',
                                                      'format': format7})

                worksheet.write('B309',
                                '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                worksheet.write('B310',
                                '- Color VERDE: los valores de campo eléctrico diario superan el valor del borde de área de cobertura (Estereofónico: >=54 dBuV/m, Monofónico: >=48 dBuV/m o Baja Potencia: >=43 dBuV/m).')
                worksheet.write('B311',
                                '- Color AMARILLO: los valores de campo eléctrico diario se encuentran entre el valor del borde de área de protección y el valor del borde de área de cobertura (Estereofónico: entre 30 y 54 dBuV/m, Monofónico: entre 30 y 48 dBuV/m o Baja Potencia: entre 30 y 43 dBuV/m).')
                worksheet.write('B312',
                                '- Color ROJO: los valores de campo eléctrico diario son inferiores al valor del borde de área de protección (<30 dBuV/m).')
                worksheet.write('B313', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                worksheet.write('B314',
                                '- Color PLOMO: La estación Color GRIS: Dispone de autorización para suspensión de emisiones.')
                worksheet.write('B315',
                                '- Para todos los casos el valor de ancho de banda corresponde a 220 kHz, excepto aquellos en los que se especifica que el valor es de 180 kHz o 200 kHz.')
                worksheet.write('A316', 'CONSIDERACIONES GENERALES:')
                worksheet.write('B317',
                                f'- Nota 1.- En apego a la Resolución ST-2014-0257 referida en los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), se considera que un control individual a una estación RTV es pertinente cuando esta ha suspendido emisiones por más de 8 días.')
                worksheet.write('B318',
                                '- Nota 2.- Las mediciones de Ancho de Banda obtenidas con el SACER son únicamente referenciales. La Unión Internacional de Telecomunicaciones – UIT, a través del Manual COMPROBACIÓN TÉCNICA DEL ESPECTRO RADIOELÉCTRICO (numeral 4.5.3) define las condiciones que deben tenerse en cuenta al medir la anchura de banda, señalando: “(…) a causa de las imprecisiones debidas a las razones expuestas, estas mediciones a distancia sólo son útiles a título indicativo. Cuando se requiera una mayor precisión, es conveniente que las mediciones se realicen en las inmediaciones del transmisor.”.')
                worksheet.write('B319',
                                f'- Nota 3.- De acuerdo a los numerales 5.2.4 y 5.2.5 de los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), para cada observación detectada que contenga datos inconsistentes que no se encuentren dentro del rango autorizado, dependiendo del caso y de los resultados, corresponde programar una verificación en sitio. Dicha programación puede estar sujeta a los resultados de las mediciones del siguiente mes.')

                for row in range(int((max_row * 2) + 1), int((max_row * 3) + 1)):
                    worksheet.set_row(row, None, None, {'hidden': True})

                """Get the dimensions of the dataframe (TV broadcasting)."""
                (max_row1, max_col1) = df_final6.drop(['Observaciones'], axis=1).shape

                """Apply a conditional format to the required cell range."""
                worksheet1.conditional_format(0, 5, 0, int(max_col1 + 1),
                                              {'type': 'no_errors',
                                               'format': format4})
                worksheet1.conditional_format(0, 0, int(max_row1), int(max_col1 + 1),
                                              {'type': 'no_errors',
                                               'format': format5})
                worksheet1.autofilter(0, 0, 0, int(max_col1 + 1))

                b = df_final4.groupby('level_0')['level_0'].count()
                max_row1 = b[0]

                i = 1
                for i in range(1, int(max_row1 + 1)):
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND(F47<>"-", F47<>"")',
                                                   'format': format7})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND(F2="-")',
                                                   'format': format6})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="D",F2>=51)',
                                                   'format': format1})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="D",F2<51)',
                                                   'format': format2})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=54000000,$B2<=88000000,F2>=68)',
                                                   'format': format1})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=54000000,$B2<=88000000,F2>=47,F2<68)',
                                                   'format': format3})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=54000000,$B2<=88000000,F2<47)',
                                                   'format': format2})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=174000000,$B2<=216000000,F2>=71)',
                                                   'format': format1})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=174000000,$B2<=216000000,F2>=56,F2<71)',
                                                   'format': format3})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=174000000,$B2<=216000000,F2<56)',
                                                   'format': format2})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=470000000,$B2<=880000000,F2>=74)',
                                                   'format': format1})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=470000000,$B2<=880000000,F2>=64,F2<74)',
                                                   'format': format3})
                    worksheet1.conditional_format(i, 5, int(max_row1), int(max_col1),
                                                  {'type': 'formula',
                                                   'criteria': '=AND($E2="-",$B2>=470000000,$B2<=880000000,F2<64)',
                                                   'format': format2})

                worksheet1.conditional_format('A94', {'type': 'no_errors',
                                                      'format': format1})
                worksheet1.conditional_format('A95', {'type': 'no_errors',
                                                      'format': format3})
                worksheet1.conditional_format('A96', {'type': 'no_errors',
                                                      'format': format2})
                worksheet1.conditional_format('A97', {'type': 'no_errors',
                                                      'format': format6})
                worksheet1.conditional_format('A98', {'type': 'no_errors',
                                                      'format': format7})
                worksheet1.write('B93',
                                 '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                worksheet1.write('B94',
                                 '- Color VERDE significa que los valores de campo eléctrico diario superan el límite del área de cobertura primaria. (Canales 2 al 6: >=68 dBuV/m, Canales 7 al 13: >=71 dBuV/m, Canales 14 al 51: >=74 dBuV/m, Canales digitales: >=51 dBuV/m).')
                worksheet1.write('B95',
                                 '- Color AMARILLO significa que los valores de campo eléctrico diario superan el límite del área de cobertura secundario pero son inferiores al límite del área de cobertura principal establecidos para cada rango. (Canales 2 al 6: entre 47 y 68 dBuV/m, Canales 7 al 13: entre 56 y 71 dBuV/m, Canales 14 al 51: entre 64 y 74 dBuV/m.')
                worksheet1.write('B96',
                                 '- Color ROJO significa que los valores de campo eléctrico diario son inferiores al límite de área de cobertura secundario. (Canales 2 al 6: <47 dBuV/m, Canales 7 al 13: <56 dBuV/m, Canales 14 al 51: <64 dBuV/m, Canales digitales: <51 dBuV/m).')
                worksheet1.write('B97', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                worksheet1.write('B98',
                                 '- Color PLOMO: La estación Color GRIS: Dispone de autorización para suspensión de emisiones.')
                worksheet1.write('A99', 'CONSIDERACIONES GENERALES:')
                worksheet1.write('B100',
                                 f'- Nota 1.- En apego a la Resolución ST-2014-0257 referida en los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), se considera que un control individual a una estación RTV es pertinente cuando esta ha suspendido emisiones por más de 8 días.')
                worksheet1.write('B101',
                                 f'- Nota 2.- De acuerdo a los numerales 5.2.4 y 5.2.5 de los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), para cada observación detectada que contenga datos inconsistentes que no se encuentren dentro del rango autorizado, dependiendo del caso y de los resultados, corresponde programar una verificación en sitio. Dicha programación puede estar sujeta a los resultados de las mediciones del siguiente mes.')

                for row1 in range(int(max_row1 + 1), int((max_row1 * 2) + 1)):
                    worksheet1.set_row(row1, None, None, {'hidden': True})

                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    """Get the dimensions of the dataframe (AM broadcasting)."""
                    (max_row7, max_col7) = df_final9.drop(['Observaciones'], axis=1).shape

                    """Apply a conditional format to the required cell range."""
                    worksheet6.conditional_format(0, 3, 0, int(max_col7 + 1),
                                                  {'type': 'no_errors',
                                                   'format': format4})
                    worksheet6.conditional_format(0, 0, int(max_row7), int(max_col7 + 1),
                                                  {'type': 'no_errors',
                                                   'format': format5})
                    worksheet6.autofilter(0, 0, 0, int(max_col7 + 1))

                    c = df_final8.groupby('level_0')['level_0'].count()
                    max_row7 = c[0]

                    r = 1
                    for r in range(1, int(max_row7 + 1)):
                        worksheet6.conditional_format(r, 3, int(max_row7), int(max_col7),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D88<>"-",D88<>"")',
                                                       'format': format7})
                        worksheet6.conditional_format(r, 3, int(max_row7), int(max_col7),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D2="-")',
                                                       'format': format6})
                        worksheet6.conditional_format(r, 3, int(max_row7), int(max_col7),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D2<62)',
                                                       'format': format2})
                        worksheet6.conditional_format(r, 3, int(max_row7), int(max_col7),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D2>=62)',
                                                       'format': format1})

                    s = int(max_row7 + 1)
                    for s in range(int(max_row7 + 1), int((2 * max_row7) + 1)):
                        worksheet6.conditional_format(s, 3, int(2 * max_row7), int(max_col7),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D88<>"-",D88<>"")',
                                                       'format': format7})
                        worksheet6.conditional_format(s, 3, int(2 * max_row7), int(max_col7),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D45="-")',
                                                       'format': format6})
                        worksheet6.conditional_format(s, 3, int(2 * max_row7), int(max_col7),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D45<=15000)',
                                                       'format': format1})
                        worksheet6.conditional_format(s, 3, int(2 * max_row7), int(max_col7),
                                                      {'type': 'formula',
                                                       'criteria': '=AND(D45>15000)',
                                                       'format': format2})

                    worksheet6.conditional_format('A133', {'type': 'no_errors',
                                                           'format': format1})
                    worksheet6.conditional_format('A134', {'type': 'no_errors',
                                                           'format': format2})
                    worksheet6.conditional_format('A135', {'type': 'no_errors',
                                                           'format': format6})
                    worksheet6.write('B132',
                                     '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet6.write('B133',
                                     '- Color VERDE: los valores de campo eléctrico diario superan el valor del borde de área de cobertura (>=62 dBuV/m).')
                    worksheet6.write('B134',
                                     '- Color ROJO: los valores de campo eléctrico diario son inferiores al valor del borde de área de protección (<62 dBuV/m).')
                    worksheet6.write('B135', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                    worksheet6.write('B136', '- Para todos los casos el valor de ancho de banda corresponde a 15 kHz.')
                    worksheet6.write('A138', 'CONSIDERACIONES GENERALES:')
                    worksheet6.write('B139',
                                     f'- Nota 1.- En apego a la Resolución ST-2014-0257 referida en los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), se considera que un control individual a una estación RTV es pertinente cuando esta ha suspendido emisiones por más de 8 días.')
                    worksheet6.write('B140',
                                     '- Nota 2.- Las mediciones de Ancho de Banda obtenidas con el SACER son únicamente referenciales. La Unión Internacional de Telecomunicaciones – UIT, a través del Manual COMPROBACIÓN TÉCNICA DEL ESPECTRO RADIOELÉCTRICO (numeral 4.5.3) define las condiciones que deben tenerse en cuenta al medir la anchura de banda, señalando: “(…) a causa de las imprecisiones debidas a las razones expuestas, estas mediciones a distancia sólo son útiles a título indicativo. Cuando se requiera una mayor precisión, es conveniente que las mediciones se realicen en las inmediaciones del transmisor.”.')
                    worksheet6.write('B141',
                                     f'- Nota 3.- De acuerdo a los numerales 5.2.4 y 5.2.5 de los LINEAMIENTOS PARA EL CONTROL Y MONITOREO DE PARÁMETROS TÉCNICOS RTV CON EL SACER (CCDE-01, PACT-{Year2}), para cada observación detectada que contenga datos inconsistentes que no se encuentren dentro del rango autorizado, dependiendo del caso y de los resultados, corresponde programar una verificación en sitio. Dicha programación puede estar sujeta a los resultados de las mediciones del siguiente mes.')

                    for row2 in range(int((max_row7 * 2) + 1), int((max_row7 * 3) + 1)):
                        worksheet6.set_row(row2, None, None, {'hidden': True})

                if Year1 == Year2 and Mes_inicio == Mes_fin:
                    """Get the dimensions of the dataframe (FM broadcasting)."""
                    (max_row2, max_col2) = df_original1.shape

                    """Apply a conditional format to the required cell range."""
                    worksheet2.conditional_format(1, 1, int(max_row2), int(max_col2),
                                                  {'type': 'no_errors',
                                                   'format': format5})
                    worksheet2.autofilter(0, 0, 0, int(max_col2))

                    """Get the dimensions of the dataframe (TV broadcasting)."""
                    (max_row3, max_col3) = df_original2.shape

                    """Apply a conditional format to the required cell range."""
                    worksheet3.conditional_format(1, 1, int(max_row3), int(max_col3),
                                                  {'type': 'no_errors',
                                                   'format': format5})
                    worksheet3.autofilter(0, 0, 0, int(max_col3))

                    if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                        """Get the dimensions of the dataframe (AM broadcasting)."""
                        (max_row8, max_col8) = df_original3.shape

                        """Apply a conditional format to the required cell range."""
                        worksheet8.conditional_format(1, 1, int(max_row8), int(max_col8),
                                                      {'type': 'no_errors',
                                                       'format': format5})
                        worksheet8.autofilter(0, 0, 0, int(max_col8))

            """Change the name of the file"""
            old_name = 'RTV_Verificación de parámetros.xlsx'
            if Year1 == Year2 and Mes_inicio == Mes_fin:
                new_name = 'RTV_Verificación de parámetros_{}_{}_{}.xlsx'.format(Ciudad, Mes_inicio, Year1)
            else:
                new_name = 'RTV_Verificación de parámetros_{}_{}{}_{}{}.xlsx'.format(Ciudad, Mes_inicio, Year1, Mes_fin,
                                                                                     Year2)

            """Remove the previous file if already exist"""
            if os.path.exists(f'{download_route}/{new_name}'):
                os.remove(f'{download_route}/{new_name}')

            """Rename the file"""
            os.rename(f'{download_route}/{old_name}',
                      f'{download_route}/{new_name}')

        elif Ocupacion == True and AM_Reporte_individual == False and FM_Reporte_individual == False and TV_Reporte_individual == False:
            """REPORTE OCUPACIÓN"""

            """Drop columns in df9 according to the requirements for the final report (FM broadcasting)"""
            df11 = df9.drop(columns=['Offset (Hz)', 'FM (Hz)', 'Bandwidth (Hz)', 'Estación', 'Potencia', 'BW Asignado'])
            df11 = df11.rename(columns={'Tiempo': 'tiempo', 'Frecuencia (Hz)': 'freq', 'Level (dBµV/m)': 'level'})

            """Drop columns in df9 according to the requirements for the final report (TV broadcasting)"""
            df12 = df10.drop(
                columns=['Offset (Hz)', 'AM (%)', 'Bandwidth (Hz)', 'Estación', 'Canal (Número)', 'Analógico/Digital'])
            df12 = df12.rename(columns={'Tiempo': 'tiempo', 'Frecuencia (Hz)': 'freq', 'Level (dBµV/m)': 'level'})

            if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                """Drop columns in df9 according to the requirements for the final report (AM broadcasting)"""
                df18 = df17.drop(columns=['Offset (Hz)', 'AM (%)', 'Bandwidth (Hz)', 'Estación'])
                df18 = df18.rename(columns={'Tiempo': 'tiempo', 'Frecuencia (Hz)': 'freq', 'Level (dBµV/m)': 'level'})

            df11["tiempo"] = pd.to_datetime(df11["tiempo"], format='%d/%m/%Y %H:%M:%S.%f')
            df19 = df11[(df11.tiempo >= fecha_inicio) & (df11.tiempo <= fecha_fin)]

            df12["tiempo"] = pd.to_datetime(df12["tiempo"], format='%d/%m/%Y %H:%M:%S.%f')
            df20 = df12[(df12.tiempo >= fecha_inicio) & (df12.tiempo <= fecha_fin)]

            if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                df18["tiempo"] = pd.to_datetime(df18["tiempo"], format='%d/%m/%Y %H:%M:%S.%f')
                df21 = df18[(df18.tiempo >= fecha_inicio) & (df18.tiempo <= fecha_fin)]

            def umb1(row):
                """function to return level value if is >= to the selected Umbral_FM"""
                if row['level'] >= Umbral_FM:
                    return row['level']
                return

            def umb2(row):
                """function to return level value if is >= to the selected Umbral_TV"""
                if row['level'] >= Umbral_TV:
                    return row['level']
                return

            def umb3(row):
                """function to return level value if is >= to the selected Umbral_AM"""
                if row['level'] >= Umbral_AM:
                    return row['level']
                return

            fecha_init = fecha_inicio.strftime('%Y-%m-%d')
            fecha_end = fecha_fin.strftime('%Y-%m-%d')

            """contar4: dataframe with the occupation by frequency (FM broadcasting)"""
            series1 = df19.drop(columns=['tiempo'])
            series1['umb'] = series1.apply(lambda row: umb1(row), axis=1)
            series1 = series1.rename(columns={'freq': 'Frecuencia (Hz)', 'level': 'total_mediciones'})
            contar1 = series1.groupby('Frecuencia (Hz)').count()
            contar1['occupation'] = ((contar1['umb'] / contar1['total_mediciones']) * 100).round(6)
            contar1 = contar1.drop(columns=['total_mediciones', 'umb'])
            contar4 = contar1.rename(columns={'occupation': 'Ocupación (%)'})

            """Draw plot FM"""
            fig, ax = plt.subplots(figsize=(20, 10), facecolor='white', dpi=80)
            ax.vlines(x=contar1.index, ymin=0, ymax=contar1.occupation, color='steelblue', alpha=0.7, linewidth=2)
            ax.scatter(x=contar1.index, y=contar1.occupation, s=75, color='steelblue', alpha=0.7)

            """Title, Label, Ticks and Ylim"""
            ax.set_title(
                f'Banda: Radiodifusión FM, Ciudad: {Ciudad}, Umbral: {Umbral_FM} dBµV/m, Periodo: {fecha_init} a {fecha_end}',
                fontdict={'size': 20})
            ax.set_xlabel('Frecuencia (Hz)')
            ax.set_ylabel('Ocupación (%)')
            ax.set_xticks(contar1.index)
            ax.set_xticklabels(contar1.index, rotation=90, fontdict={'horizontalalignment': 'center', 'size': 10})
            ax.set_ylim(0, 100)

            """Annotate"""
            for row in contar1.itertuples():
                ax.text(row.Index, row.occupation + .5, s=int(row.occupation), horizontalalignment='center',
                        verticalalignment='bottom', fontsize=9)

            plt.grid(color='#292929', linestyle='--', linewidth=0.5)
            """save the plot in the file image1.png"""
            plt.savefig('image1.png')
            plt.close()

            """contar5: dataframe with the occupation by frequency (TV broadcasting)"""
            series2 = df20.drop(columns=['tiempo'])
            series2['umb'] = series2.apply(lambda row: umb2(row), axis=1)
            series2 = series2.rename(columns={'freq': 'Frecuencia (Hz)', 'level': 'total_mediciones'})
            contar2 = series2.groupby('Frecuencia (Hz)').count()
            contar2['occupation'] = ((contar2['umb'] / contar2['total_mediciones']) * 100).round(6)
            contar2 = contar2.drop(columns=['total_mediciones', 'umb'])
            contar5 = contar2.rename(columns={'occupation': 'Ocupación (%)'})

            """Draw plot TV"""
            fig1, ax1 = plt.subplots(figsize=(20, 10), facecolor='white', dpi=80)
            ax1.vlines(x=contar2.index, ymin=0, ymax=contar2.occupation, color='steelblue', alpha=0.7, linewidth=2)
            ax1.scatter(x=contar2.index, y=contar2.occupation, s=75, color='steelblue', alpha=0.7)

            """Title, Label, Ticks and Ylim"""
            ax1.set_title(
                f'Banda: Televisión Abierta, Ciudad: {Ciudad}, Umbral: {Umbral_TV} dBµV/m, Periodo: {fecha_init} a {fecha_end}',
                fontdict={'size': 20})
            ax1.set_xlabel('Frecuencia (Hz)')
            ax1.set_ylabel('Ocupación (%)')
            ax1.set_xticks(contar2.index)
            ax1.set_xticklabels(contar2.index, rotation=90, fontdict={'horizontalalignment': 'center', 'size': 10})
            ax1.set_ylim(0, 100)

            """Annotate"""
            for row in contar2.itertuples():
                ax1.text(row.Index, row.occupation + .5, s=int(row.occupation), horizontalalignment='center',
                         verticalalignment='bottom', fontsize=9)

            plt.grid(color='#292929', linestyle='--', linewidth=0.5)
            """save the plot in the file image2.png"""
            plt.savefig('image2.png')
            plt.close()

            if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                """contar6: dataframe with the occupation by frequency (AM broadcasting)"""
                series3 = df21.drop(columns=['tiempo'])
                series3['umb'] = series3.apply(lambda row: umb3(row), axis=1)
                series3 = series3.rename(columns={'freq': 'Frecuencia (Hz)', 'level': 'total_mediciones'})
                contar3 = series3.groupby('Frecuencia (Hz)').count()
                contar3['occupation'] = ((contar3['umb'] / contar3['total_mediciones']) * 100).round(6)
                contar3 = contar3.drop(columns=['total_mediciones', 'umb'])
                contar6 = contar3.rename(columns={'occupation': 'Ocupación (%)'})

                """Draw plot AM"""
                fig2, ax2 = plt.subplots(figsize=(20, 10), facecolor='white', dpi=80)
                ax2.vlines(x=contar3.index, ymin=0, ymax=contar3.occupation, color='steelblue', alpha=0.7, linewidth=2)
                ax2.scatter(x=contar3.index, y=contar3.occupation, s=75, color='steelblue', alpha=0.7)

                """Title, Label, Ticks and Ylim"""
                ax2.set_title(
                    f'Banda: Radiodifusión AM, Ciudad: {Ciudad}, Umbral: {Umbral_AM} dBµV/m, Periodo: {fecha_init} a {fecha_end}',
                    fontdict={'size': 20})
                ax2.set_xlabel('Frecuencia (Hz)')
                ax2.set_ylabel('Ocupación (%)')
                ax2.set_xticks(contar3.index)
                ax2.set_xticklabels(contar3.index, rotation=90, fontdict={'horizontalalignment': 'center', 'size': 10})
                ax2.set_ylim(0, 100)

                """Annotate"""
                for row in contar3.itertuples():
                    ax2.text(row.Index, row.occupation + .5, s=int(row.occupation), horizontalalignment='center',
                             verticalalignment='bottom', fontsize=9)

                plt.grid(color='#292929', linestyle='--', linewidth=0.5)
                """save the plot in the file image3.png"""
                plt.savefig('image3.png')
                plt.close()

            """REPORT CREATION"""
            """create, write and save"""
            with pd.ExcelWriter(f'{download_route}/RTV_Ocupación.xlsx') as writer:
                contar4.to_excel(writer, sheet_name='OC_Radiodifusión FM')
                contar5.to_excel(writer, sheet_name='OC_Televisión')
                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    contar6.to_excel(writer, sheet_name='OC_Radiodifusión AM')

                """Get the xlsxwriter workbook and worksheet objects."""
                workbook = writer.book
                worksheet = writer.sheets['OC_Radiodifusión FM']
                worksheet1 = writer.sheets['OC_Televisión']
                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    worksheet2 = writer.sheets['OC_Radiodifusión AM']

                """Add a format."""
                format1 = workbook.add_format({'border': 1, 'border_color': 'black'})
                format2 = workbook.add_format({'bold': True})

                """Get the dimensions of the dataframe (FM broadcasting)."""
                (max_row, max_col) = contar4.shape

                """Apply a conditional format to the required cell range."""
                worksheet.conditional_format(0, 1, int(max_row), int(max_col),
                                             {'type': 'no_errors',
                                              'format': format1})
                worksheet.autofilter(0, 0, 0, int(max_col))

                worksheet.write('C1', 'Umbral (dBµV/m)')
                worksheet.write('C2', float(Umbral_FM))

                worksheet.conditional_format('C1:C2', {'type': 'no_errors',
                                                       'format': format1})
                worksheet.conditional_format('C1', {'type': 'no_errors',
                                                    'format': format2})
                worksheet.insert_image('E1', 'image1.png')

                """Get the dimensions of the dataframe (TV broadcasting)."""
                (max_row1, max_col1) = contar5.shape

                """Apply a conditional format to the required cell range."""
                worksheet1.conditional_format(0, 1, int(max_row1), int(max_col1),
                                              {'type': 'no_errors',
                                               'format': format1})
                worksheet1.autofilter(0, 0, 0, int(max_col1))

                worksheet1.write('C1', 'Umbral (dBµV/m)')
                worksheet1.write('C2', float(Umbral_TV))

                worksheet1.conditional_format('C1:C2', {'type': 'no_errors',
                                                        'format': format1})
                worksheet1.conditional_format('C1', {'type': 'no_errors',
                                                     'format': format2})
                worksheet1.insert_image('E1', 'image2.png')

                if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                    """Get the dimensions of the dataframe (AM broadcasting)."""
                    (max_row2, max_col2) = contar6.shape

                    """Apply a conditional format to the required cell range."""
                    worksheet2.conditional_format(0, 1, int(max_row2), int(max_col2),
                                                  {'type': 'no_errors',
                                                   'format': format1})
                    worksheet2.autofilter(0, 0, 0, int(max_col2))

                    worksheet2.write('C1', 'Umbral (dBµV/m)')
                    worksheet2.write('C2', float(Umbral_AM))

                    worksheet2.conditional_format('C1:C2', {'type': 'no_errors',
                                                            'format': format1})
                    worksheet2.conditional_format('C1', {'type': 'no_errors',
                                                         'format': format2})
                    worksheet2.insert_image('E1', 'image3.png')

            """Change the name of the file"""
            old_name = 'RTV_Ocupación.xlsx'
            if Year1 == Year2 and Mes_inicio == Mes_fin:
                new_name = f'RTV_Ocupación_{Ciudad}_{Mes_inicio}_{Year1}.xlsx'
            else:
                new_name = f'RTV_Ocupación_{Ciudad}_{Mes_inicio} {Year1}-{Mes_fin} {Year2}.xlsx'

            """Remove the previous file if already exist"""
            if os.path.exists(f'{download_route}/{new_name}'):
                os.remove(f'{download_route}/{new_name}')

            """Rename the file"""
            os.rename(f'{download_route}/{old_name}',
                      f'{download_route}/{new_name}')

            """Remove image files"""
            os.remove('image1.png')
            os.remove('image2.png')
            if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                os.remove('image3.png')

        elif Ocupacion == False and AM_Reporte_individual == False and FM_Reporte_individual == True and TV_Reporte_individual == False and Autorizaciones == False:
            """REPORTE INDIVIDUAL FM"""

            """Get the data only for the selected "Frecuencia_FM" and save it in the dataframe df9"""
            df9 = df9.loc[df9['Frecuencia (Hz)'] == int(Frecuencia_FM)]
            """Group the information in df9 according to the requirements for the final report: max Level and average 
            Bandwidth per Frequency and Day (FM broadcasting)"""
            df11 = df9.groupby(by=[
                pd.Grouper(key='Tiempo', freq='D'),
                pd.Grouper(key='Frecuencia (Hz)'),
                pd.Grouper(key='Estación'),
                pd.Grouper(key='Potencia'),
                pd.Grouper(key='BW Asignado')
            ]).agg({
                'Level (dBµV/m)': 'max',
                'Bandwidth (Hz)': 'mean'
            }).reset_index()

            """Make the pivot tables with the data structured in the way we want to show in the report 
            (FM broadcasting)"""
            df_final1 = pd.pivot_table(df11,
                                       index=[pd.Grouper(key='Tiempo')],
                                       values=['Level (dBµV/m)', 'Bandwidth (Hz)'],
                                       columns=['Frecuencia (Hz)', 'Estación', 'Potencia', 'BW Asignado'],
                                       aggfunc={'Level (dBµV/m)': max, 'Bandwidth (Hz)': np.average}).round(2)
            df_final1 = df_final1.T
            df_final3 = df_final1.replace(0, '-')
            """Reset the index (unstack) to have the columns 'Potencia', 'BW Asignado' and 'Tiempo' in the index so we
            can use the flags in the columns 'Potencia' and 'BW Asignado' """
            df_final3 = df_final3.reset_index()

            """sorter first by 'Level (dBµV/m)' and after by 'Bandwidth (Hz)' and rename the column header as
            'Parámetros' """
            sorter = ['Level (dBµV/m)', 'Bandwidth (Hz)']
            df_final3.level_0 = df_final3.level_0.astype("category")
            df_final3.level_0 = df_final3.level_0.cat.set_categories(sorter)
            df_final3 = df_final3.sort_values(["level_0"])
            df_final3 = df_final3.rename(columns={'level_0': 'Parámetros'})
            df_final3 = df_final3.set_index('Parámetros')

            """PLOT"""
            series = df11
            series = series.groupby(pd.Grouper(key='Tiempo', freq='1D')).max()
            series = series.rename(
                columns={'Frecuencia (Hz)': 'freq', 'Estación': 'est', 'Potencia': 'pot', 'BW Asignado': 'bw',
                         'Level (dBµV/m)': 'level', 'Bandwidth (Hz)': 'bandwidth'})
            nombre = series['est'].iloc[0]
            pot = series['pot'].iloc[0]
            bw = series['bw'].iloc[0]
            if pot == '-':
                pot = 0
            else:
                pot = 1
            if bw == '-':
                bw = 220

            def minus(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['level'] > 0 and row['level'] < 30:
                    return row['level']
                return 0

            def bet(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if pot == 0 and bw == 220:
                    if row['level'] >= 30 and row['level'] < 54:
                        return row['level']
                    return 0
                elif pot == 0 and bw == 200:
                    if row['level'] >= 30 and row['level'] < 54:
                        return row['level']
                    return 0
                elif pot == 0 and bw == 180:
                    if row['level'] >= 30 and row['level'] < 48:
                        return row['level']
                elif pot == 1:
                    if row['level'] >= 30 and row['level'] < 43:
                        return row['level']
                    return 0

            def plus(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if pot == 0 and bw == 220:
                    if row['level'] >= 54:
                        return row['level']
                    return 0
                elif pot == 0 and bw == 200:
                    if row['level'] >= 54:
                        return row['level']
                    return 0
                elif pot == 0 and bw == 180:
                    if row['level'] >= 48:
                        return row['level']
                elif pot == 1:
                    if row['level'] >= 43:
                        return row['level']
                    return 0

            def valor(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['level'] == 0:
                    return 120
                return 0

            """create a new column in the series frame for every definition (minus, bet, plus, valor)"""
            series['minus'] = series.apply(lambda row: minus(row), axis=1)
            series['bet'] = series.apply(lambda row: bet(row), axis=1)
            series['plus'] = series.apply(lambda row: plus(row), axis=1)
            series['valor'] = series.apply(lambda row: valor(row), axis=1)

            fig = plt.figure()
            """111 means 1x1 grid, first subplot, so it will put every subplot in the same figure"""
            ax = fig.add_subplot(111, ylabel='Nivel de Intensidad de Campo Eléctrico (dBµV/m)', xlabel='Tiempo',
                                 title=f'Ciudad: {Ciudad}, Estación: {nombre}, Frecuencia: {Frecuencia_FM} Hz')
            """set the y-axis limits"""
            ax.set_ylim(0, 120)
            """area plot of every definition"""
            series['plus'].plot.area(y=ax, colormap='Accent', linewidth=0)
            series['bet'].plot.area(y=ax, colormap='Set3_r', linewidth=0)
            series['minus'].plot.area(y=ax, colormap='Pastel1', linewidth=0)
            if Seleccionar == True:
                series['valor'].plot.area(y=ax, colormap='Paired', linewidth=0)
            elif Seleccionar == False:
                try:
                    series['level'][series.level == 0].plot(y=ax, marker='v', markersize=7, color='b', linewidth=0)
                except ValueError:  # raise if array is empty (no cero values for level column)
                    pass
            plt.grid(color='#292929', linestyle='--', linewidth=0.5)
            """save the plot in the finle image.png"""
            plt.savefig('image.png')
            """close the plot to not be show after the code is executed"""
            plt.close()

            """REPORT CREATION: REPORTE INDIVIDUAL FM"""
            """create, write and save"""
            with pd.ExcelWriter(f'{download_route}/RTV_Verificación de parámetros.xlsx') as writer:
                df_final3.to_excel(writer, sheet_name=f'Radiodifusión FM_{Frecuencia_FM} Hz')

                """Get the xlsxwriter workbook and worksheet objects."""
                workbook = writer.book
                worksheet = writer.sheets[f'Radiodifusión FM_{Frecuencia_FM} Hz']

                """Add a format."""
                format1 = workbook.add_format({'bg_color': '#C6EFCE',
                                               'font_color': '#006100'})

                format2 = workbook.add_format({'bg_color': '#FFC7CE',
                                               'font_color': '#9C0006'})

                format3 = workbook.add_format({'bg_color': '#FFDC47',
                                               'font_color': '#9C6500'})

                format4 = workbook.add_format({'num_format': 'dd/mm/yy'})

                format5 = workbook.add_format({'border': 1, 'border_color': 'black'})

                format6 = workbook.add_format({'bg_color': '#99CCFF',
                                               'font_color': '#0066FF'})

                """Get the dimensions of the dataframe."""
                (max_row, max_col) = df_final3.shape

                """Apply a conditional format to the required cell range."""
                worksheet.conditional_format(0, 5, 0, int(max_col),
                                             {'type': 'no_errors',
                                              'format': format4})
                worksheet.conditional_format(0, 1, int(max_row), int(max_col),
                                             {'type': 'no_errors',
                                              'format': format5})
                worksheet.conditional_format(3, 5, 3, int(max_col),
                                             {'type': 'no_errors',
                                              'format': format4})
                worksheet.autofilter(0, 0, 0, int(max_col))

                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F3="-")',
                                              'format': format6})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F3<30)',
                                              'format': format2})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$3="BAJA",F3>=43)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$3="BAJA",F3>=30,F3<43)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$3="-",$E$3="-",F3>=54)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$3="-",$E$3="-",F3>=30,F3<54)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$3="-",$E$3=200,F3>=54)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$3="-",$E$3=200,F3>=30,F3<54)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$3="-",$E$3=180,F3>=48)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$3="-",$E$3=180,F3>=30,F3<48)',
                                              'format': format3})

                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F2="-")',
                                              'format': format6})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$2=180,F2<=180000)',
                                              'format': format1})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$2=180,F2>180000)',
                                              'format': format2})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$2=200,F2<=200000)',
                                              'format': format1})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$2=200,F2>200000)',
                                              'format': format2})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$2="-",F2<=220000)',
                                              'format': format1})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$2="-",F2>220000)',
                                              'format': format2})

                worksheet.conditional_format('A6', {'type': 'no_errors',
                                                    'format': format1})
                worksheet.conditional_format('A7', {'type': 'no_errors',
                                                    'format': format3})
                worksheet.conditional_format('A8', {'type': 'no_errors',
                                                    'format': format2})
                worksheet.conditional_format('A9', {'type': 'no_errors',
                                                    'format': format6})
                if pot == 0 and bw == 220:
                    tipo = 'Estereofónico'
                    maximo = 54
                elif pot == 0 and bw == 200:
                    tipo = 'Estereofónico'
                    maximo = 54
                elif pot == 0 and bw == 180:
                    tipo = 'Monofónico'
                    maximo = 48
                elif pot == 1:
                    tipo = 'Baja Potencia'
                    maximo = 43
                minimo = 30

                worksheet.write('B5',
                                '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                worksheet.write('B6',
                                f'- Color VERDE: los valores de campo eléctrico diario superan el valor del borde de área de cobertura ({tipo}: >={maximo} dBuV/m).')
                worksheet.write('B7',
                                f'- Color AMARILLO: los valores de campo eléctrico diario se encuentran entre el valor del borde de área de protección y el valor del borde de área de cobertura ({tipo}: entre {minimo} y {maximo} dBuV/m).')
                worksheet.write('B8',
                                f'- Color ROJO: los valores de campo eléctrico diario son inferiores al valor del borde de área de protección ({tipo}: <{minimo} dBuV/m).')
                worksheet.write('B9', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                worksheet.write('B10', f'- El valor de ancho de banda corresponde a {int(bw)} kHz.')
                worksheet.insert_image('A12', 'image.png')

            """Change the name of the file"""
            old_name = 'RTV_Verificación de parámetros.xlsx'
            if Year1 == Year2 and Mes_inicio == Mes_fin:
                new_name = 'FM_Verificación de parámetros_{} ({} Hz)_{}_{}_{}.xlsx'.format(nombre, Frecuencia_FM,
                                                                                           Ciudad,
                                                                                           Mes_inicio, Year1)
            else:
                new_name = 'FM_Verificación de parámetros_{} ({} Hz)_{}_{}{}_{}{}.xlsx'.format(nombre, Frecuencia_FM,
                                                                                               Ciudad,
                                                                                               Mes_inicio, Year1,
                                                                                               Mes_fin,
                                                                                               Year2)

            """Remove the previous file if already exist"""
            if os.path.exists(f'{download_route}/{new_name}'):
                os.remove(f'{download_route}/{new_name}')

            """Rename the file"""
            os.rename(f'{download_route}/{old_name}',
                      f'{download_route}/{new_name}')

            """Remove the image file"""
            os.remove('image.png')

        elif Ocupacion == False and AM_Reporte_individual == False and FM_Reporte_individual == False and TV_Reporte_individual == True and Autorizaciones == False:
            """REPORTE INDIVIDUAL TV"""

            """Get the data only for the selected "Canal_TV" and save it in the dataframe df10"""
            df10 = df10.loc[df10['Canal (Número)'] == float(Canal_TV)]

            """Group the information in df10 acording to the requirements for the final report: max Level per Frequency 
            and Day (TV Broadcasting)"""
            df12 = df10.groupby(by=[
                pd.Grouper(key='Tiempo', freq='D'),
                pd.Grouper(key='Frecuencia (Hz)'),
                pd.Grouper(key='Estación'),
                pd.Grouper(key='Canal (Número)'),
                pd.Grouper(key='Analógico/Digital'),
            ]).agg({
                'Level (dBµV/m)': 'max'
            }).reset_index()

            """Make the pivot tables with the data structured in the way we want to show in the report
            (TV broadcasting)"""
            df_final2 = pd.pivot_table(df12,
                                       index=[pd.Grouper(key='Tiempo')],
                                       values=['Level (dBµV/m)'],
                                       columns=['Frecuencia (Hz)', 'Estación', 'Canal (Número)', 'Analógico/Digital'],
                                       aggfunc={'Level (dBµV/m)': max}).round(2)
            df_final2 = df_final2.T
            df_final4 = df_final2.replace(0, '-')
            df_final4 = df_final4.reset_index()

            """sorter first by 'Level (dBµV/m)' and rename the column header as 'Parámetros' """
            sorter = ['Level (dBµV/m)']
            df_final4.level_0 = df_final4.level_0.astype("category")
            df_final4.level_0 = df_final4.level_0.cat.set_categories(sorter)
            df_final4 = df_final4.sort_values(["level_0"])
            df_final4 = df_final4.rename(columns={'level_0': 'Parámetros'})
            df_final4 = df_final4.set_index('Parámetros')

            """PLOT"""
            series = df12
            series = series.groupby(pd.Grouper(key='Tiempo', freq='1D')).max()
            series = series.rename(
                columns={'Frecuencia (Hz)': 'freq', 'Estación': 'est', 'Canal (Número)': 'canal',
                         'Analógico/Digital': 'ad',
                         'Level (dBµV/m)': 'level'})
            nombre = series['est'].iloc[0]
            andig = series['ad'].iloc[0]
            if andig == '-':
                andig = 0
            else:
                andig = 1

            def minus(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['freq'] >= 54000000 and row['freq'] <= 88000000 and andig == 0:
                    if row['level'] > 0 and row['level'] < 47:
                        return row['level']
                    return 0
                elif row['freq'] >= 174000000 and row['freq'] <= 216000000 and andig == 0:
                    if row['level'] > 0 and row['level'] < 56:
                        return row['level']
                    return 0
                elif row['freq'] >= 470000000 and row['freq'] <= 880000000 and andig == 0:
                    if row['level'] > 0 and row['level'] < 64:
                        return row['level']
                    return 0
                elif andig == 1:
                    if row['level'] > 0 and row['level'] < 30:
                        return row['level']
                    return 0

            def bet(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['freq'] >= 54000000 and row['freq'] <= 88000000 and andig == 0:
                    if row['level'] >= 47 and row['level'] < 68:
                        return row['level']
                    return 0
                elif row['freq'] >= 174000000 and row['freq'] <= 216000000 and andig == 0:
                    if row['level'] >= 56 and row['level'] < 71:
                        return row['level']
                    return 0
                elif row['freq'] >= 470000000 and row['freq'] <= 880000000 and andig == 0:
                    if row['level'] >= 64 and row['level'] < 74:
                        return row['level']
                    return 0
                elif andig == 1:
                    if row['level'] >= 30 and row['level'] < 51:
                        return row['level']
                    return 0

            def plus(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['freq'] >= 54000000 and row['freq'] <= 88000000 and andig == 0:
                    if row['level'] >= 68:
                        return row['level']
                    return 0
                elif row['freq'] >= 174000000 and row['freq'] <= 216000000 and andig == 0:
                    if row['level'] >= 71:
                        return row['level']
                    return 0
                elif row['freq'] >= 470000000 and row['freq'] <= 880000000 and andig == 0:
                    if row['level'] >= 74:
                        return row['level']
                    return 0
                elif andig == 1:
                    if row['level'] >= 51:
                        return row['level']
                    return 0

            def valor(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['level'] == 0:
                    return 120
                return 0

            """create a new column in the series frame for every definition (minus, bet, plus, valor)"""
            series['minus'] = series.apply(lambda row: minus(row), axis=1)
            series['bet'] = series.apply(lambda row: bet(row), axis=1)
            series['plus'] = series.apply(lambda row: plus(row), axis=1)
            series['valor'] = series.apply(lambda row: valor(row), axis=1)

            fig = plt.figure()
            """111 means 1x1 grid, first subplot, so it will put every subplot in the same figure"""
            ax = fig.add_subplot(111, ylabel='Nivel de Intensidad de Campo Eléctrico (dBµV/m)', xlabel='Tiempo',
                                 title=f'Ciudad: {Ciudad}, Estación: {nombre}, Canal: {Canal_TV}')
            """set the y-axis limits"""
            ax.set_ylim(0, 120)
            """area plot of every definition"""
            series['plus'].plot.area(y=ax, colormap='Accent', linewidth=0)
            series['bet'].plot.area(y=ax, colormap='Set3_r', linewidth=0)
            series['minus'].plot.area(y=ax, colormap='Pastel1', linewidth=0)
            if Seleccionar == True:
                series['valor'].plot.area(y=ax, colormap='Paired', linewidth=0)
            elif Seleccionar == False:
                try:
                    series['level'][series.level == 0].plot(y=ax, marker='v', markersize=7, color='b', linewidth=0)
                except ValueError:  # raise if array is empty (no cero values for level column)
                    pass
            plt.grid(color='#292929', linestyle='--', linewidth=0.5)
            """save the plot in the finle image.png"""
            plt.savefig('image.png')
            """close the plot to not be show after the code is executed"""
            plt.close()

            """REPORT CREATION: REPORTE INDIVIDUAL TV"""
            """create, write and save"""
            with pd.ExcelWriter(f'{download_route}/RTV_Verificación de parámetros.xlsx') as writer:
                df_final4.to_excel(writer, sheet_name=f'Televisión_Canal {Canal_TV}')

                """Get the xlsxwriter workbook and worksheet objects."""
                workbook = writer.book
                worksheet = writer.sheets[f'Televisión_Canal {Canal_TV}']

                """Add a format."""
                format1 = workbook.add_format({'bg_color': '#C6EFCE',
                                               'font_color': '#006100'})

                format2 = workbook.add_format({'bg_color': '#FFC7CE',
                                               'font_color': '#9C0006'})

                format3 = workbook.add_format({'bg_color': '#FFDC47',
                                               'font_color': '#9C6500'})

                format4 = workbook.add_format({'num_format': 'dd/mm/yy'})

                format5 = workbook.add_format({'border': 1, 'border_color': 'black'})

                format6 = workbook.add_format({'bg_color': '#99CCFF',
                                               'font_color': '#0066FF'})

                """Get the dimensions of the dataframe."""
                (max_row, max_col) = df_final4.shape

                """Apply a conditional format to the required cell range."""
                worksheet.conditional_format(0, 5, 0, int(max_col),
                                             {'type': 'no_errors',
                                              'format': format4})
                worksheet.conditional_format(0, 1, int(max_row), int(max_col),
                                             {'type': 'no_errors',
                                              'format': format5})
                worksheet.autofilter(0, 0, 0, int(max_col))

                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F2="-")',
                                              'format': format6})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="D",F2>=51)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="D",F2<51)',
                                              'format': format2})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=54000000,$B$2<=88000000,F2>=68)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=54000000,$B$2<=88000000,F2>=47,F2<68)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=54000000,$B$2<=88000000,F2<47)',
                                              'format': format2})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=174000000,$B$2<=216000000,F2>=71)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=174000000,$B$2<=216000000,F2>=56,F2<71)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=174000000,$B$2<=216000000,F2<56)',
                                              'format': format2})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=470000000,$B$2<=880000000,F2>=74)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=470000000,$B$2<=880000000,F2>=64,F2<74)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=470000000,$B$2<=880000000,F2<64)',
                                              'format': format2})

                worksheet.conditional_format('A5', {'type': 'no_errors',
                                                    'format': format1})
                worksheet.conditional_format('A6', {'type': 'no_errors',
                                                    'format': format3})
                worksheet.conditional_format('A7', {'type': 'no_errors',
                                                    'format': format2})
                worksheet.conditional_format('A8', {'type': 'no_errors',
                                                    'format': format6})

                if series['freq'].iloc[0] >= 54000000 and series['freq'].iloc[0] <= 88000000 and andig == 0:
                    worksheet.write('B4',
                                    '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet.write('B5',
                                    f'- Color VERDE significa que los valores de campo eléctrico diario superan el límite del área de cobertura primaria (Canal {Canal_TV} >= 68 dBuV/m).')
                    worksheet.write('B6',
                                    f'- Color AMARILLO significa que los valores de campo eléctrico diario superan el límite del área de cobertura secundario pero son inferiores al límite del área de cobertura principal establecido (Canal {Canal_TV}: entre 47 y 68 dBuV/m).')
                    worksheet.write('B7',
                                    f'- Color ROJO significa que los valores de campo eléctrico diario son inferiores al límite de área de cobertura secundario. (Canal {Canal_TV}: < 47 dBuV/m).')
                    worksheet.write('B8', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                elif series['freq'].iloc[0] >= 174000000 and series['freq'].iloc[0] <= 216000000 and andig == 0:
                    worksheet.write('B4',
                                    '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet.write('B5',
                                    f'- Color VERDE significa que los valores de campo eléctrico diario superan el límite del área de cobertura primaria (Canal {Canal_TV} >= 71 dBuV/m).')
                    worksheet.write('B6',
                                    f'- Color AMARILLO significa que los valores de campo eléctrico diario superan el límite del área de cobertura secundario pero son inferiores al límite del área de cobertura principal establecido (Canal {Canal_TV}: entre 56 y 71 dBuV/m).')
                    worksheet.write('B7',
                                    f'- Color ROJO significa que los valores de campo eléctrico diario son inferiores al límite de área de cobertura secundario. (Canal {Canal_TV}: < 56 dBuV/m).')
                    worksheet.write('B8', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                elif series['freq'].iloc[0] >= 470000000 and series['freq'].iloc[0] <= 880000000 and andig == 0:
                    worksheet.write('B4',
                                    '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet.write('B5',
                                    f'- Color VERDE significa que los valores de campo eléctrico diario superan el límite del área de cobertura primaria (Canal {Canal_TV} >= 74 dBuV/m).')
                    worksheet.write('B6',
                                    f'- Color AMARILLO significa que los valores de campo eléctrico diario superan el límite del área de cobertura secundario pero son inferiores al límite del área de cobertura principal establecido (Canal {Canal_TV}: entre 64 y 74 dBuV/m).')
                    worksheet.write('B7',
                                    f'- Color ROJO significa que los valores de campo eléctrico diario son inferiores al límite de área de cobertura secundario. (Canal {Canal_TV}: < 64 dBuV/m).')
                    worksheet.write('B8', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                elif andig == 1:
                    worksheet.write('B4',
                                    '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet.write('B5',
                                    f'- Color VERDE significa que los valores de campo eléctrico diario superan el límite del área de cobertura primaria (Canal {Canal_TV} >= 51 dBuV/m).')
                    worksheet.write('B6',
                                    f'- Color AMARILLO significa que los valores de campo eléctrico diario superan el límite del área de cobertura secundario pero son inferiores al límite del área de cobertura principal establecido (Canal {Canal_TV}: entre 30 y 51 dBuV/m).')
                    worksheet.write('B7',
                                    f'- Color ROJO significa que los valores de campo eléctrico diario son inferiores al límite de área de cobertura secundario. (Canal {Canal_TV}: < 30 dBuV/m).')
                    worksheet.write('B8', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                worksheet.insert_image('A10', 'image.png')

            """Change the name of the file"""
            old_name = 'RTV_Verificación de parámetros.xlsx'
            if Year1 == Year2 and Mes_inicio == Mes_fin:
                new_name = 'TV_Verificación de parámetros_{} (Canal {})_{}_{}_{}.xlsx'.format(nombre, Canal_TV, Ciudad,
                                                                                              Mes_inicio, Year1)
            else:
                new_name = 'TV_Verificación de parámetros_{} (Canal {})_{}_{}{}_{}{}.xlsx'.format(nombre, Canal_TV,
                                                                                                  Ciudad,
                                                                                                  Mes_inicio, Year1,
                                                                                                  Mes_fin,
                                                                                                  Year2)

            """Remove the previous file if already exist"""
            if os.path.exists(f'{download_route}/{new_name}'):
                os.remove(f'{download_route}/{new_name}')

            """Rename the file"""
            os.rename(f'{download_route}/{old_name}',
                      f'{download_route}/{new_name}')

            """Remove the image file"""
            os.remove('image.png')

        elif Ocupacion == False and AM_Reporte_individual == True and FM_Reporte_individual == False and TV_Reporte_individual == False and Autorizaciones == False:
            """REPORTE INDIVIDUAL AM"""
            if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':

                """Get the data only for the selected "Frecuencia_AM" and save it in the dataframe df17"""
                df17 = df17.loc[df17['Frecuencia (Hz)'] == int(Frecuencia_AM)]

                """Group the information in df17 acording to the requirements for the final report: max Level and average
                 Bandwidth per Day"""
                df18 = df17.groupby(by=[
                    pd.Grouper(key='Tiempo', freq='D'),
                    pd.Grouper(key='Frecuencia (Hz)'),
                    pd.Grouper(key='Estación')
                ]).agg({
                    'Level (dBµV/m)': 'max',
                    'Bandwidth (Hz)': 'mean'
                }).reset_index()

                """make the pivot table with the data structured in the way we want to show in the report"""
                df_final7 = pd.pivot_table(df18,
                                           index=[pd.Grouper(key='Tiempo')],
                                           values=['Level (dBµV/m)', 'Bandwidth (Hz)'],
                                           columns=['Frecuencia (Hz)', 'Estación'],
                                           aggfunc={'Level (dBµV/m)': max, 'Bandwidth (Hz)': np.average}).round(2)
                df_final7 = df_final7.T
                df_final8 = df_final7.replace(0, '-')

                """Reset the index (unstack) to have the columns 'Potencia', 'BW Asignado' and 'Tiempo' in the index 
                so we can use the flags in the columns 'Potencia' and 'BW Asignado'"""
                df_final8 = df_final8.reset_index()

                """sorter first by 'Level (dBµV/m)' and after by 'Bandwidth (Hz)' and rename the column header as 
                'Parámetros'"""
                sorter = ['Level (dBµV/m)', 'Bandwidth (Hz)']
                df_final8.level_0 = df_final8.level_0.astype("category")
                df_final8.level_0 = df_final8.level_0.cat.set_categories(sorter)
                df_final8 = df_final8.sort_values(["level_0"])
                df_final8 = df_final8.rename(columns={'level_0': 'Parámetros'})
                df_final8 = df_final8.set_index('Parámetros')

                """PLOT"""
                series = df18
                series = series.groupby(pd.Grouper(key='Tiempo', freq='1D')).max()
                series = series.rename(columns={'Frecuencia (Hz)': 'freq', 'Estación': 'est', 'Level (dBµV/m)': 'level',
                                                'Bandwidth (Hz)': 'bandwidth'})
                nombre = series['est'].iloc[0]

                def minus(row):
                    """function to return a specific value if the value in every row of the column 'level' meet the
                    condition"""
                    if row['level'] > 0 and row['level'] < 40:
                        return row['level']
                    return 0

                def bet(row):
                    """function to return a specific value if the value in every row of the column 'level' meet the
                    condition"""
                    if row['level'] >= 40 and row['level'] < 62:
                        return row['level']
                    return 0

                def plus(row):
                    """function to return a specific value if the value in every row of the column 'level' meet the
                    condition"""
                    if row['level'] >= 62:
                        return row['level']
                    return 0

                def valor(row):
                    """function to return a specific value if the value in every row of the column 'level' meet the
                    condition"""
                    if row['level'] == 0:
                        return 120
                    return 0

                """create a new column in the series frame for every definition (minus, bet, plus, valor)"""
                series['minus'] = series.apply(lambda row: minus(row), axis=1)
                series['bet'] = series.apply(lambda row: bet(row), axis=1)
                series['plus'] = series.apply(lambda row: plus(row), axis=1)
                series['valor'] = series.apply(lambda row: valor(row), axis=1)

                fig = plt.figure()
                """111 means 1x1 grid, first subplot, so it will put every subplot in the same figure"""
                ax = fig.add_subplot(111, ylabel='Nivel de Intensidad de Campo Eléctrico (dBµV/m)', xlabel='Tiempo',
                                     title=f'Ciudad: {Ciudad}, Estación: {nombre}, Frecuencia: {Frecuencia_AM} Hz')
                """set the y-axis limits"""
                ax.set_ylim(0, 120)
                """area plot of every definition"""
                series['plus'].plot.area(y=ax, colormap='Accent', linewidth=0)
                series['bet'].plot.area(y=ax, colormap='Set3_r', linewidth=0)
                series['minus'].plot.area(y=ax, colormap='Pastel1', linewidth=0)
                if Seleccionar == True:
                    series['valor'].plot.area(y=ax, colormap='Paired', linewidth=0)
                elif Seleccionar == False:
                    try:
                        series['level'][series.level == 0].plot(y=ax, marker='v', markersize=7, color='b', linewidth=0)
                    except ValueError:  # raise if array is empty (no cero values for level column)
                        pass
                plt.grid(color='#292929', linestyle='--', linewidth=0.5)
                """save the plot in the file image.png"""
                plt.savefig('image.png')
                """close the plot to not be show after the code is executed"""
                plt.close()

                """REPORT CREATION: REPORTE INDIVIDUAL AM"""
                """create, write and save"""
                with pd.ExcelWriter(f'{download_route}/RTV_Verificación de parámetros.xlsx') as writer:
                    df_final8.to_excel(writer, sheet_name=f'Radiodifusión AM_{Frecuencia_AM} Hz')

                    """Get the xlsxwriter workbook and worksheet objects."""
                    workbook = writer.book
                    worksheet = writer.sheets[f'Radiodifusión AM_{Frecuencia_AM} Hz']

                    """Add a format."""
                    format1 = workbook.add_format({'bg_color': '#C6EFCE',
                                                   'font_color': '#006100'})

                    format2 = workbook.add_format({'bg_color': '#FFC7CE',
                                                   'font_color': '#9C0006'})

                    format3 = workbook.add_format({'bg_color': '#FFDC47',
                                                   'font_color': '#9C6500'})

                    format4 = workbook.add_format({'num_format': 'dd/mm/yy'})

                    format5 = workbook.add_format({'border': 1, 'border_color': 'black'})

                    format6 = workbook.add_format({'bg_color': '#99CCFF',
                                                   'font_color': '#0066FF'})

                    """Get the dimensions of the dataframe."""
                    (max_row, max_col) = df_final8.shape

                    """Apply a conditional format to the required cell range."""
                    worksheet.conditional_format(0, 3, 0, int(max_col),
                                                 {'type': 'no_errors',
                                                  'format': format4})
                    worksheet.conditional_format(0, 1, int(max_row), int(max_col),
                                                 {'type': 'no_errors',
                                                  'format': format5})
                    worksheet.autofilter(0, 0, 0, int(max_col))

                    worksheet.conditional_format(1, 3, 1, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D3="-")',
                                                  'format': format6})
                    worksheet.conditional_format(1, 3, 1, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D3<40)',
                                                  'format': format2})
                    worksheet.conditional_format(1, 3, 1, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D3>=62)',
                                                  'format': format1})
                    worksheet.conditional_format(1, 3, 1, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D3>=40,F3<62)',
                                                  'format': format3})

                    worksheet.conditional_format(2, 3, 2, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D2="-")',
                                                  'format': format6})
                    worksheet.conditional_format(2, 3, 2, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D2<=15000)',
                                                  'format': format1})
                    worksheet.conditional_format(2, 3, 2, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D2>15000)',
                                                  'format': format2})

                    worksheet.conditional_format('A6', {'type': 'no_errors',
                                                        'format': format1})
                    worksheet.conditional_format('A7', {'type': 'no_errors',
                                                        'format': format3})
                    worksheet.conditional_format('A8', {'type': 'no_errors',
                                                        'format': format2})
                    worksheet.conditional_format('A9', {'type': 'no_errors',
                                                        'format': format6})

                    worksheet.write('B5',
                                    '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet.write('B6',
                                    '- Color VERDE: los valores de campo eléctrico diario superan el valor del borde de área de cobertura (>=62 dBuV/m).')
                    worksheet.write('B7',
                                    '- Color AMARILLO: los valores de campo eléctrico diario se encuentran entre el valor del borde de área de protección y el valor del borde de área de cobertura (entre 40 y 62 dBuV/m).')
                    worksheet.write('B8',
                                    '- Color ROJO: los valores de campo eléctrico diario son inferiores al valor del borde de área de protección (<40 dBuV/m).')
                    worksheet.write('B9', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                    worksheet.write('B10', '- El valor de ancho de banda corresponde a 15 kHz.')
                    worksheet.insert_image('A12', 'image.png')

                """Change the name of the file"""
                old_name = 'RTV_Verificación de parámetros.xlsx'
                if Year1 == Year2 and Mes_inicio == Mes_fin:
                    new_name = 'AM_Verificación de parámetros_{} ({} Hz)_{}_{}_{}.xlsx'.format(nombre, Frecuencia_AM,
                                                                                               Ciudad,
                                                                                               Mes_inicio, Year1)
                else:
                    new_name = 'AM_Verificación de parámetros_{} ({} Hz)_{}_{}{}_{}{}.xlsx'.format(nombre,
                                                                                                   Frecuencia_AM,
                                                                                                   Ciudad, Mes_inicio,
                                                                                                   Year1,
                                                                                                   Mes_fin, Year2)

                """Remove the previous file if already exist"""
                if os.path.exists(f'{download_route}/{new_name}'):
                    os.remove(f'{download_route}/{new_name}')

                """Rename the file"""
                os.rename(f'{download_route}/{old_name}',
                          f'{download_route}/{new_name}')

                """Remove the image file"""
                os.remove('image.png')

        elif Ocupacion == False and AM_Reporte_individual == False and FM_Reporte_individual == True and TV_Reporte_individual == False and Autorizaciones == True:
            """REPORTE INDIVIDUAL FM - AUTORIZACIONES"""
            """Get the data only for the selected "Frecuencia_FM" and save it in the dataframe df9"""
            df9 = df9.loc[df9['Frecuencia (Hz)'] == int(Frecuencia_FM)]

            """Filter dfau1 to get FM frequencies"""
            dfau2 = dfau1[(dfau1.freq > 87700000) & (dfau1.freq < 108100000)]
            """copy the column 'Fecha_inicio' to a list and save it in col1"""
            col1 = dfau2['Fecha_inicio'].copy().tolist()
            """insert a column in the index 13 with the name 'Fecha_ini' with the col1 information"""
            dfau2.insert(13, 'Fecha_ini', col1)
            dfau2 = dfau2.rename(
                columns={'freq': 'Frecuencia (Hz)', 'Fecha_inicio': 'Tiempo', 'Fecha_ini': 'Fecha_inicio'})
            dfau2 = dfau2.drop(columns=['est'])
            """Filter dfau2 by 'Frecuencia_FM'"""
            dfau2 = dfau2.loc[dfau2['Frecuencia (Hz)'] == int(Frecuencia_FM)]
            """Filter dfau2 by 'autori' (name of the city)"""
            dfau2 = dfau2.loc[dfau2['ciu'] == autori]

            """Group the information in df9 according to the requirements for the final report: max Level and average
            Bandwidth per Frequency and Day (FM broadcasting)"""
            df9 = df9.groupby(by=[
                pd.Grouper(key='Tiempo', freq='D'),
                pd.Grouper(key='Frecuencia (Hz)'),
                pd.Grouper(key='Estación'),
                pd.Grouper(key='Potencia'),
                pd.Grouper(key='BW Asignado')
            ]).agg({
                'Level (dBµV/m)': 'max',
                'Bandwidth (Hz)': 'mean'
            }).reset_index()

            dfau2 = dfau2.drop(columns=['ciu'])
            dfau3 = []
            for index, row in dfau2.iterrows():
                for t in pd.date_range(start=row['Tiempo'], end=row['Fecha_fin']):
                    dfau3.append(
                        (row['Frecuencia (Hz)'], row['Tipo'], row['Plazo'], t, row['Oficio'], row['Fecha_oficio'],
                         row['Fecha_inicio'], row['Fecha_fin']))
            dfau3 = pd.DataFrame(dfau3, columns=(
                'Frecuencia (Hz)', 'Tipo', 'Plazo', 'Tiempo', 'Oficio', 'Fecha_oficio', 'Fecha_inicio', 'Fecha_fin'))

            """Merge dfau3 with df9 dataframe to add the autorization, df9 is the dataframe that contains all the 
            information for FM"""
            df9 = dfau3.merge(df9, how='right', on=['Tiempo', 'Frecuencia (Hz)'])
            df9 = df9.fillna('-')

            """Make the pivot tables with the data structured in the way we want to show in the report
            (FM broadcasting)"""
            df_final1 = pd.pivot_table(df9,
                                       index=[pd.Grouper(key='Tiempo')],
                                       values=['Level (dBµV/m)', 'Bandwidth (Hz)', 'Fecha_fin'],
                                       columns=['Frecuencia (Hz)', 'Estación', 'Potencia', 'BW Asignado'],
                                       aggfunc={'Level (dBµV/m)': max, 'Bandwidth (Hz)': np.average,
                                                'Fecha_fin': max}).round(2)
            df_final1 = df_final1.rename(columns={'Fecha_fin': 'Fin de Autorización'})
            df_final1 = df_final1.T
            df_final3 = df_final1.replace(0, '-')
            """Reset the index (unstack) to have the columns 'Potencia', 'BW Asignado' and 'Tiempo' in the index so we 
            can use the flags in the columns 'Potencia' and 'BW Asignado'"""
            df_final3 = df_final3.reset_index()

            """sorter first by 'Level (dBµV/m)', after by 'Bandwidth (Hz)' and finally by 'Fin de Autorización', and 
            rename the column header as 'Parámetros' """
            sorter = ['Level (dBµV/m)', 'Bandwidth (Hz)', 'Fin de Autorización']
            df_final3.level_0 = df_final3.level_0.astype("category")
            df_final3.level_0 = df_final3.level_0.cat.set_categories(sorter)
            df_final3 = df_final3.sort_values(["level_0"])
            df_final3 = df_final3.rename(columns={'level_0': 'Parámetros'})
            df_final3 = df_final3.set_index('Parámetros')

            """PLOT"""
            series = df9
            series = series.groupby(pd.Grouper(key='Tiempo', freq='1D')).max()
            series = series.rename(
                columns={'Frecuencia (Hz)': 'freq', 'Estación': 'est', 'Potencia': 'pot', 'BW Asignado': 'bw',
                         'Level (dBµV/m)': 'level', 'Bandwidth (Hz)': 'bandwidth'})
            series['Fecha_inicio'] = series['Fecha_inicio'].replace('-', 0)
            series['Fecha_fin'] = series['Fecha_fin'].replace('-', 0)
            nombre = series['est'].iloc[0]
            pot = series['pot'].iloc[0]
            bw = series['bw'].iloc[0]
            if pot == '-':
                pot = 0
            else:
                pot = 1
            if bw == '-':
                bw = 220

            def minus(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['level'] > 0 and row['level'] < 30:
                    return row['level']
                return 0

            def bet(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if pot == 0 and bw == 220:
                    if row['level'] >= 30 and row['level'] < 54:
                        return row['level']
                    return 0
                elif pot == 0 and bw == 200:
                    if row['level'] >= 30 and row['level'] < 54:
                        return row['level']
                    return 0
                elif pot == 0 and bw == 180:
                    if row['level'] >= 30 and row['level'] < 48:
                        return row['level']
                elif pot == 1:
                    if row['level'] >= 30 and row['level'] < 43:
                        return row['level']
                    return 0

            def plus(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if pot == 0 and bw == 220:
                    if row['level'] >= 54:
                        return row['level']
                    return 0
                elif pot == 0 and bw == 200:
                    if row['level'] >= 54:
                        return row['level']
                    return 0
                elif pot == 0 and bw == 180:
                    if row['level'] >= 48:
                        return row['level']
                elif pot == 1:
                    if row['level'] >= 43:
                        return row['level']
                    return 0

            def valor(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['level'] == 0:
                    return 120
                return 0

            def aut(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['Fecha_fin'] != 0 and row['level'] == 0:
                    return 0
                elif row['Fecha_fin'] != 0 and row['level'] != 0:
                    return row['level']
                return 0

            """create a new column in the series frame for every definition (minus, bet, plus, valor, aut)"""
            series['minus'] = series.apply(lambda row: minus(row), axis=1)
            series['bet'] = series.apply(lambda row: bet(row), axis=1)
            series['plus'] = series.apply(lambda row: plus(row), axis=1)
            series['valor'] = series.apply(lambda row: valor(row), axis=1)
            series['aut'] = series.apply(lambda row: aut(row), axis=1)

            fig = plt.figure()
            """111 means 1x1 grid, first subplot, so it will put every subplot in the same figure"""
            ax = fig.add_subplot(111, ylabel='Nivel de Intensidad de Campo Eléctrico (dBµV/m)', xlabel='Tiempo',
                                 title=f'Ciudad: {Ciudad}, Estación: {nombre}, Frecuencia: {Frecuencia_FM} Hz')
            """set the y-axis limits"""
            ax.set_ylim(0, 120)
            """area plot of every definition"""
            series['plus'].plot.area(y=ax, colormap='Accent', linewidth=0)
            series['bet'].plot.area(y=ax, colormap='Set3_r', linewidth=0)
            series['minus'].plot.area(y=ax, colormap='Pastel1', linewidth=0)
            if Seleccionar == True:
                series['valor'].plot.area(y=ax, colormap='Paired', linewidth=0)
            elif Seleccionar == False:
                try:
                    series['level'][series.level == 0].plot(y=ax, marker='v', markersize=7, color='b', linewidth=0)
                except ValueError:  # raise if array is empty (no cero values for level column)
                    pass
            if Autorizaciones == True:
                series['aut'].plot.area(y=ax, colormap='Set2_r', linewidth=0)

            """Annotations for initial and final dates of the authorization"""
            ini_date = series.Fecha_inicio.tolist()
            fin_date = series.Fecha_fin.tolist()

            for mark_time1 in ini_date:
                time_x1 = pd.Timestamp(mark_time1)  # convert to compatible format
                time_x1 = time_x1.strftime('%Y-%m-%d')
                ax.annotate(f'Inicio: {time_x1}', xy=(time_x1, 0), xycoords='data',
                            bbox=dict(boxstyle="round", fc="white", ec="black"),
                            xytext=(time_x1, 12), ha='center',
                            arrowprops=dict(arrowstyle="->",
                                            connectionstyle="angle, angleA = 0, angleB = 90, rad = 10")
                            )

            for mark_time2 in fin_date:
                time_x2 = pd.Timestamp(mark_time2)  # convert to compatible format
                time_x2 = time_x2.strftime('%Y-%m-%d')
                ax.annotate(f'Fin: {time_x2}', xy=(time_x2, 0), xycoords='data',
                            bbox=dict(boxstyle="round", fc="white", ec="black"),
                            xytext=(time_x2, 6), ha='center',
                            arrowprops=dict(arrowstyle="->",
                                            connectionstyle="angle, angleA = 0, angleB = 90, rad = 10")
                            )

            plt.grid(color='#292929', linestyle='--', linewidth=0.5)
            """save the plot in the file image.png"""
            plt.savefig('image.png')
            """close the plot to not be show after the code is executed"""
            plt.close()

            """REPORT CREATION: REPORTE INDIVIDUAL FM - AUTORIZACIONES"""
            """create, write and save"""
            with pd.ExcelWriter(f'{download_route}/RTV_Verificación de parámetros.xlsx') as writer:
                df_final3.to_excel(writer, sheet_name=f'Radiodifusión FM_{Frecuencia_FM} Hz')
                dfau2.set_index('Frecuencia (Hz)').rename(
                    columns={'Fecha_inicio': 'Inicio Autorización', 'Fecha_fin': 'Fin Autorización'}).drop(
                    columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).to_excel(writer,
                                                                                         sheet_name='Autorizaciones')

                """Get the xlsxwriter workbook and worksheet objects."""
                workbook = writer.book
                worksheet = writer.sheets[f'Radiodifusión FM_{Frecuencia_FM} Hz']
                worksheet1 = writer.sheets['Autorizaciones']

                """Add a format."""
                format1 = workbook.add_format({'bg_color': '#C6EFCE',
                                               'font_color': '#006100'})

                format2 = workbook.add_format({'bg_color': '#FFC7CE',
                                               'font_color': '#9C0006'})

                format3 = workbook.add_format({'bg_color': '#FFDC47',
                                               'font_color': '#9C6500'})

                format4 = workbook.add_format({'num_format': 'dd/mm/yy'})

                format5 = workbook.add_format({'border': 1, 'border_color': 'black'})

                format6 = workbook.add_format({'bg_color': '#99CCFF',
                                               'font_color': '#0066FF'})

                format7 = workbook.add_format({'bg_color': '#C0C0C0',
                                               'font_color': '#000000'})

                """Get the dimensions of the dataframe."""
                (max_row1, max_col1) = dfau2.set_index('Frecuencia (Hz)').rename(
                    columns={'Fecha_inicio': 'Inicio Autorización', 'Fecha_fin': 'Fin Autorización'}).drop(
                    columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).shape

                """Apply a conditional format to the required cell range."""
                worksheet1.conditional_format(1, 1, int(max_row1), int(max_col1),
                                              {'type': 'no_errors',
                                               'format': format5})
                worksheet1.autofilter(0, 0, 0, int(max_col1))

                """Get the dimensions of the dataframe."""
                (max_row, max_col) = df_final3.shape

                """Apply a conditional format to the required cell range."""
                worksheet.conditional_format(0, 5, 0, int(max_col),
                                             {'type': 'no_errors',
                                              'format': format4})
                worksheet.conditional_format(0, 1, int(max_row), int(max_col),
                                             {'type': 'no_errors',
                                              'format': format5})
                worksheet.conditional_format(3, 5, 3, int(max_col),
                                             {'type': 'no_errors',
                                              'format': format4})
                worksheet.autofilter(0, 0, 0, int(max_col))

                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F2="-")',
                                              'format': format6})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F2<30)',
                                              'format': format2})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$2="BAJA",F2>=43)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$2="BAJA",F2>=30,F2<43)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$2="-",$E$2="-",F2>=54)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$2="-",$E$2="-",F2>=30,F2<54)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$2="-",$E$2=200,F2>=54)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$2="-",$E$2=200,F2>=30,F2<54)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$2="-",$E$2=180,F2>=48)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($D$2="-",$E$2=180,F2>=30,F2<48)',
                                              'format': format3})

                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F3="-")',
                                              'format': format6})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$3=180,F3<=180000)',
                                              'format': format1})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$3=180,F3>180000)',
                                              'format': format2})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$3=200,F3<=200000)',
                                              'format': format1})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$3=200,F3>200000)',
                                              'format': format2})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$3="-",F3<=220000)',
                                              'format': format1})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E$3="-",F3>220000)',
                                              'format': format2})

                worksheet.conditional_format(3, 5, 3, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F4="-")',
                                              'format': format6})
                worksheet.conditional_format(3, 5, 3, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F4<>"-")',
                                              'format': format7})

                worksheet.conditional_format('A7', {'type': 'no_errors',
                                                    'format': format1})
                worksheet.conditional_format('A8', {'type': 'no_errors',
                                                    'format': format3})
                worksheet.conditional_format('A9', {'type': 'no_errors',
                                                    'format': format2})
                worksheet.conditional_format('A10', {'type': 'no_errors',
                                                     'format': format6})
                worksheet.conditional_format('A11', {'type': 'no_errors',
                                                     'format': format7})
                if pot == 0 and bw == 220:
                    tipo = 'Estereofónico'
                    maximo = 54
                elif pot == 0 and bw == 200:
                    tipo = 'Estereofónico'
                    maximo = 54
                elif pot == 0 and bw == 180:
                    tipo = 'Monofónico'
                    maximo = 48
                elif pot == 1:
                    tipo = 'Baja Potencia'
                    maximo = 43
                minimo = 30

                worksheet.write('B6',
                                '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                worksheet.write('B7',
                                f'- Color VERDE: los valores de campo eléctrico diario superan el valor del borde de área de cobertura ({tipo}: >={maximo} dBuV/m).')
                worksheet.write('B8',
                                f'- Color AMARILLO: los valores de campo eléctrico diario se encuentran entre el valor del borde de área de protección y el valor del borde de área de cobertura ({tipo}: entre {minimo} y {maximo} dBuV/m).')
                worksheet.write('B9',
                                f'- Color ROJO: los valores de campo eléctrico diario son inferiores al valor del borde de área de protección ({tipo}: <{minimo} dBuV/m).')
                worksheet.write('B10', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                worksheet.write('B11', '- Color GRIS: Dispone de autorización para suspensión de emisiones.')
                worksheet.write('B12', f'- El valor de ancho de banda corresponde a {int(bw)} kHz.')
                worksheet.insert_image('A14', 'image.png')

            """Change the name of the file"""
            old_name = 'RTV_Verificación de parámetros.xlsx'
            if Year1 == Year2 and Mes_inicio == Mes_fin:
                new_name = 'FM_Verificación de parámetros_{} ({} Hz)_{}_{}_{}.xlsx'.format(nombre, Frecuencia_FM,
                                                                                           Ciudad,
                                                                                           Mes_inicio, Year1)
            else:
                new_name = 'FM_Verificación de parámetros_{} ({} Hz)_{}_{}{}_{}{}.xlsx'.format(nombre, Frecuencia_FM,
                                                                                               Ciudad,
                                                                                               Mes_inicio, Year1,
                                                                                               Mes_fin,
                                                                                               Year2)

            """Remove the previous file if already exist"""
            if os.path.exists(f'{download_route}/{new_name}'):
                os.remove(f'{download_route}/{new_name}')

            """Rename the file"""
            os.rename(f'{download_route}/{old_name}',
                      f'{download_route}/{new_name}')

            """Remove the image file"""
            os.remove('image.png')

        elif Ocupacion == False and AM_Reporte_individual == False and FM_Reporte_individual == False and TV_Reporte_individual == True and Autorizaciones == True:
            """REPORTE INDIVIDUAL TV - AUTORIZACIONES"""

            """Get the data only for the selected "Canal_TV" and save it in the dataframe df10"""
            df10 = df10.loc[df10['Canal (Número)'] == int(Canal_TV)]

            """Filter dfau1 to get TV channels"""
            dfau2 = dfau1[(dfau1.freq >= 2) & (dfau1.freq <= 51)]
            """copy the column 'Fecha_inicio' to a list and save it in col1"""
            col1 = dfau2['Fecha_inicio'].copy().tolist()
            """insert a column in the index 13 with the name 'Fecha_ini' with the col1 information"""
            dfau2.insert(13, 'Fecha_ini', col1)
            dfau2 = dfau2.rename(
                columns={'freq': 'Canal (Número)', 'Fecha_inicio': 'Tiempo', 'Fecha_ini': 'Fecha_inicio'})
            dfau2 = dfau2.drop(columns=['est'])
            """Filter dfau2 by 'Canal_TV'"""
            dfau2 = dfau2.loc[dfau2['Canal (Número)'] == int(Canal_TV)]
            """Filter dfau2 by 'autori' (name of the city)"""
            dfau2 = dfau2.loc[dfau2['ciu'] == autori]

            """Group the information in df10 according to the requirements for the final report: max Level per Frequency
             and Day (TV broadcasting)"""
            df10 = df10.groupby(by=[
                pd.Grouper(key='Tiempo', freq='D'),
                pd.Grouper(key='Frecuencia (Hz)'),
                pd.Grouper(key='Estación'),
                pd.Grouper(key='Canal (Número)'),
                pd.Grouper(key='Analógico/Digital'),
            ]).agg({
                'Level (dBµV/m)': 'max'
            }).reset_index()

            dfau2 = dfau2.drop(columns=['ciu'])
            dfau3 = []
            for index, row in dfau2.iterrows():
                for t in pd.date_range(start=row['Tiempo'], end=row['Fecha_fin']):
                    dfau3.append(
                        (row['Canal (Número)'], row['Tipo'], row['Plazo'], t, row['Oficio'], row['Fecha_oficio'],
                         row['Fecha_inicio'], row['Fecha_fin']))
            dfau3 = pd.DataFrame(dfau3, columns=(
                'Canal (Número)', 'Tipo', 'Plazo', 'Tiempo', 'Oficio', 'Fecha_oficio', 'Fecha_inicio', 'Fecha_fin'))

            """Merge dfau3 with df10 dataframe to add the autorization, df10 is the dataframe that contains all the
            information for TV"""
            df10 = dfau3.merge(df10, how='right', on=['Tiempo', 'Canal (Número)'])
            df10 = df10.fillna('-')

            """Make the pivot tables with the data structured in the way we want to show in the report
            (TV broadcasting)"""
            df_final2 = pd.pivot_table(df10,
                                       index=[pd.Grouper(key='Tiempo')],
                                       values=['Level (dBµV/m)', 'Fecha_fin'],
                                       columns=['Frecuencia (Hz)', 'Estación', 'Canal (Número)', 'Analógico/Digital'],
                                       aggfunc={'Level (dBµV/m)': max, 'Fecha_fin': max}).round(2)
            df_final2 = df_final2.rename(columns={'Fecha_fin': 'Fin de Autorización'})
            df_final2 = df_final2.T
            df_final4 = df_final2.replace(0, '-')
            df_final4 = df_final4.reset_index()

            """sorter first by 'Level (dBµV/m)' and after by 'Fin de Autorización' and rename the column header as
            'Parámetros' """
            sorter = ['Level (dBµV/m)', 'Fin de Autorización']
            df_final4.level_0 = df_final4.level_0.astype("category")
            df_final4.level_0 = df_final4.level_0.cat.set_categories(sorter)
            df_final4 = df_final4.sort_values(["level_0"])
            df_final4 = df_final4.rename(columns={'level_0': 'Parámetros'})
            df_final4 = df_final4.set_index('Parámetros')

            """PLOT"""
            series = df10
            series = series.groupby(pd.Grouper(key='Tiempo', freq='1D')).max()
            series = series.rename(
                columns={'Frecuencia (Hz)': 'freq', 'Estación': 'est', 'Canal (Número)': 'canal',
                         'Analógico/Digital': 'ad',
                         'Level (dBµV/m)': 'level'})
            series['Fecha_inicio'] = series['Fecha_inicio'].replace('-', 0)
            series['Fecha_fin'] = series['Fecha_fin'].replace('-', 0)
            nombre = series['est'].iloc[0]
            andig = series['ad'].iloc[0]
            if andig == '-':
                andig = 0
            else:
                andig = 1

            def minus(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['freq'] >= 54000000 and row['freq'] <= 88000000 and andig == 0:
                    if row['level'] > 0 and row['level'] < 47:
                        return row['level']
                    return 0
                elif row['freq'] >= 174000000 and row['freq'] <= 216000000 and andig == 0:
                    if row['level'] > 0 and row['level'] < 56:
                        return row['level']
                    return 0
                elif row['freq'] >= 470000000 and row['freq'] <= 880000000 and andig == 0:
                    if row['level'] > 0 and row['level'] < 64:
                        return row['level']
                    return 0
                elif andig == 1:
                    if row['level'] > 0 and row['level'] < 30:
                        return row['level']
                    return 0

            def bet(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['freq'] >= 54000000 and row['freq'] <= 88000000 and andig == 0:
                    if row['level'] >= 47 and row['level'] < 68:
                        return row['level']
                    return 0
                elif row['freq'] >= 174000000 and row['freq'] <= 216000000 and andig == 0:
                    if row['level'] >= 56 and row['level'] < 71:
                        return row['level']
                    return 0
                elif row['freq'] >= 470000000 and row['freq'] <= 880000000 and andig == 0:
                    if row['level'] >= 64 and row['level'] < 74:
                        return row['level']
                    return 0
                elif andig == 1:
                    if row['level'] >= 30 and row['level'] < 51:
                        return row['level']
                    return 0

            def plus(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['freq'] >= 54000000 and row['freq'] <= 88000000 and andig == 0:
                    if row['level'] >= 68:
                        return row['level']
                    return 0
                elif row['freq'] >= 174000000 and row['freq'] <= 216000000 and andig == 0:
                    if row['level'] >= 71:
                        return row['level']
                    return 0
                elif row['freq'] >= 470000000 and row['freq'] <= 880000000 and andig == 0:
                    if row['level'] >= 74:
                        return row['level']
                    return 0
                elif andig == 1:
                    if row['level'] >= 51:
                        return row['level']
                    return 0

            def valor(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['level'] == 0:
                    return 120
                return 0

            def aut(row):
                """function to return a specific value if the value in every row of the column 'level' meet the
                condition"""
                if row['Fecha_fin'] != 0 and row['level'] == 0:
                    return 0
                elif row['Fecha_fin'] != 0 and row['level'] != 0:
                    return row['level']
                return 0

            """create a new column in the series frame for every definition (minus, bet, plus, valor, aut)"""
            series['minus'] = series.apply(lambda row: minus(row), axis=1)
            series['bet'] = series.apply(lambda row: bet(row), axis=1)
            series['plus'] = series.apply(lambda row: plus(row), axis=1)
            series['valor'] = series.apply(lambda row: valor(row), axis=1)
            series['aut'] = series.apply(lambda row: aut(row), axis=1)

            fig = plt.figure()
            """111 means 1x1 grid, first subplot, so it will put every subplot in the same figure"""
            ax = fig.add_subplot(111, ylabel='Nivel de Intensidad de Campo Eléctrico (dBµV/m)', xlabel='Tiempo',
                                 title=f'Ciudad: {Ciudad}, Estación: {nombre}, Canal: {Canal_TV}')
            """set the y-axis limits"""
            ax.set_ylim(0, 120)
            """area plot of every definition"""
            series['plus'].plot.area(y=ax, colormap='Accent', linewidth=0)
            series['bet'].plot.area(y=ax, colormap='Set3_r', linewidth=0)
            series['minus'].plot.area(y=ax, colormap='Pastel1', linewidth=0)
            if Seleccionar == True:
                series['valor'].plot.area(y=ax, colormap='Paired', linewidth=0)
            elif Seleccionar == False:
                try:
                    series['level'][series.level == 0].plot(y=ax, marker='v', markersize=7, color='b', linewidth=0)
                except ValueError:  # raise if array is empty (no cero values for level column)
                    pass
            if Autorizaciones == True:
                series['aut'].plot.area(y=ax, colormap='Set2_r', linewidth=0)

            """Annotations for initial and final dates of the authorization"""
            ini_date = series.Fecha_inicio.tolist()
            fin_date = series.Fecha_fin.tolist()

            for mark_time1 in ini_date:
                time_x1 = pd.Timestamp(mark_time1)  # convert to compatible format
                time_x1 = time_x1.strftime('%Y-%m-%d')
                ax.annotate(f'Inicio: {time_x1}', xy=(time_x1, 0), xycoords='data',
                            bbox=dict(boxstyle="round", fc="white", ec="black"),
                            xytext=(time_x1, 12), ha='center',
                            arrowprops=dict(arrowstyle="->",
                                            connectionstyle="angle, angleA = 0, angleB = 90, rad = 10")
                            )

            for mark_time2 in fin_date:
                time_x2 = pd.Timestamp(mark_time2)  # convert to compatible format
                time_x2 = time_x2.strftime('%Y-%m-%d')
                ax.annotate(f'Fin: {time_x2}', xy=(time_x2, 0), xycoords='data',
                            bbox=dict(boxstyle="round", fc="white", ec="black"),
                            xytext=(time_x2, 6), ha='center',
                            arrowprops=dict(arrowstyle="->",
                                            connectionstyle="angle, angleA = 0, angleB = 90, rad = 10")
                            )

            plt.grid(color='#292929', linestyle='--', linewidth=0.5)
            """save the plot in the finle image.png"""
            plt.savefig('image.png')
            """close the plot to not be show after the code is executed"""
            plt.close()

            """REPORT CREATION: REPORTE INDIVIDUAL TV - AUTORIZACIONES"""
            """create, write and save"""
            with pd.ExcelWriter(f'{download_route}/RTV_Verificación de parámetros.xlsx') as writer:
                df_final4.to_excel(writer, sheet_name=f'Televisión_Canal {Canal_TV}')
                dfau2.set_index('Canal (Número)').rename(
                    columns={'Fecha_inicio': 'Inicio Autorización', 'Fecha_fin': 'Fin Autorización'}).drop(
                    columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).to_excel(writer,
                                                                                         sheet_name='Autorizaciones')

                """Get the xlsxwriter workbook and worksheet objects."""
                workbook = writer.book
                worksheet = writer.sheets[f'Televisión_Canal {Canal_TV}']
                worksheet1 = writer.sheets['Autorizaciones']

                """Add a format."""
                format1 = workbook.add_format({'bg_color': '#C6EFCE',
                                               'font_color': '#006100'})

                format2 = workbook.add_format({'bg_color': '#FFC7CE',
                                               'font_color': '#9C0006'})

                format3 = workbook.add_format({'bg_color': '#FFDC47',
                                               'font_color': '#9C6500'})

                format4 = workbook.add_format({'num_format': 'dd/mm/yy'})

                format5 = workbook.add_format({'border': 1, 'border_color': 'black'})

                format6 = workbook.add_format({'bg_color': '#99CCFF',
                                               'font_color': '#0066FF'})

                format7 = workbook.add_format({'bg_color': '#C0C0C0',
                                               'font_color': '#000000'})

                """Get the dimensions of the dataframe."""
                (max_row1, max_col1) = dfau2.set_index('Canal (Número)').rename(
                    columns={'Fecha_inicio': 'Inicio Autorización', 'Fecha_fin': 'Fin Autorización'}).drop(
                    columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).shape

                """Apply a conditional format to the required cell range."""
                worksheet1.conditional_format(1, 1, int(max_row1), int(max_col1),
                                              {'type': 'no_errors',
                                               'format': format5})
                worksheet1.autofilter(0, 0, 0, int(max_col1))

                """Get the dimensions of the dataframe."""
                (max_row, max_col) = df_final4.shape

                """Apply a conditional format to the required cell range."""
                worksheet.conditional_format(0, 5, 0, int(max_col),
                                             {'type': 'no_errors',
                                              'format': format4})
                worksheet.conditional_format(0, 1, int(max_row), int(max_col),
                                             {'type': 'no_errors',
                                              'format': format5})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'no_errors',
                                              'format': format4})
                worksheet.autofilter(0, 0, 0, int(max_col))

                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F2="-")',
                                              'format': format6})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="D",F2>=51)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="D",F2<51)',
                                              'format': format2})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=54000000,$B$2<=88000000,F2>=68)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=54000000,$B$2<=88000000,F2>=47,F2<68)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=54000000,$B$2<=88000000,F2<47)',
                                              'format': format2})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=174000000,$B$2<=216000000,F2>=71)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=174000000,$B$2<=216000000,F2>=56,F2<71)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=174000000,$B$2<=216000000,F2<56)',
                                              'format': format2})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=470000000,$B$2<=880000000,F2>=74)',
                                              'format': format1})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=470000000,$B$2<=880000000,F2>=64,F2<74)',
                                              'format': format3})
                worksheet.conditional_format(1, 5, 1, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND($E2="-",$B$2>=470000000,$B$2<=880000000,F2<64)',
                                              'format': format2})

                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F3="-")',
                                              'format': format6})
                worksheet.conditional_format(2, 5, 2, int(max_col),
                                             {'type': 'formula',
                                              'criteria': '=AND(F3<>"-")',
                                              'format': format7})

                worksheet.conditional_format('A6', {'type': 'no_errors',
                                                    'format': format1})
                worksheet.conditional_format('A7', {'type': 'no_errors',
                                                    'format': format3})
                worksheet.conditional_format('A8', {'type': 'no_errors',
                                                    'format': format2})
                worksheet.conditional_format('A9', {'type': 'no_errors',
                                                    'format': format6})
                worksheet.conditional_format('A10', {'type': 'no_errors',
                                                     'format': format7})
                if series['freq'].iloc[0] >= 54000000 and series['freq'].iloc[0] <= 88000000 and andig == 0:
                    worksheet.write('B5',
                                    '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet.write('B6',
                                    f'- Color VERDE significa que los valores de campo eléctrico diario superan el límite del área de cobertura primaria (Canal {Canal_TV} >= 68 dBuV/m).')
                    worksheet.write('B7',
                                    f'- Color AMARILLO significa que los valores de campo eléctrico diario superan el límite del área de cobertura secundario pero son inferiores al límite del área de cobertura principal establecido (Canal {Canal_TV}: entre 47 y 68 dBuV/m).')
                    worksheet.write('B8',
                                    f'- Color ROJO significa que los valores de campo eléctrico diario son inferiores al límite de área de cobertura secundario. (Canal {Canal_TV}: < 47 dBuV/m).')
                    worksheet.write('B9', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                    worksheet.write('B10', '- Color GRIS: Dispone de autorización para suspensión de emisiones.')
                elif series['freq'].iloc[0] >= 174000000 and series['freq'].iloc[0] <= 216000000 and andig == 0:
                    worksheet.write('B5',
                                    '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet.write('B6',
                                    f'- Color VERDE significa que los valores de campo eléctrico diario superan el límite del área de cobertura primaria (Canal {Canal_TV} >= 71 dBuV/m).')
                    worksheet.write('B7',
                                    f'- Color AMARILLO significa que los valores de campo eléctrico diario superan el límite del área de cobertura secundario pero son inferiores al límite del área de cobertura principal establecido (Canal {Canal_TV}: entre 56 y 71 dBuV/m).')
                    worksheet.write('B8',
                                    f'- Color ROJO significa que los valores de campo eléctrico diario son inferiores al límite de área de cobertura secundario. (Canal {Canal_TV}: < 56 dBuV/m).')
                    worksheet.write('B9', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                    worksheet.write('B10', '- Color GRIS: Dispone de autorización para suspensión de emisiones.')
                elif series['freq'].iloc[0] >= 470000000 and series['freq'].iloc[0] <= 880000000 and andig == 0:
                    worksheet.write('B5',
                                    '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet.write('B6',
                                    f'- Color VERDE significa que los valores de campo eléctrico diario superan el límite del área de cobertura primaria (Canal {Canal_TV} >= 74 dBuV/m).')
                    worksheet.write('B7',
                                    f'- Color AMARILLO significa que los valores de campo eléctrico diario superan el límite del área de cobertura secundario pero son inferiores al límite del área de cobertura principal establecido (Canal {Canal_TV}: entre 64 y 74 dBuV/m).')
                    worksheet.write('B8',
                                    f'- Color ROJO significa que los valores de campo eléctrico diario son inferiores al límite de área de cobertura secundario. (Canal {Canal_TV}: < 64 dBuV/m).')
                    worksheet.write('B9', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                    worksheet.write('B10', '- Color GRIS: Dispone de autorización para suspensión de emisiones.')
                elif andig == 1:
                    worksheet.write('B5',
                                    '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet.write('B6',
                                    f'- Color VERDE significa que los valores de campo eléctrico diario superan el límite del área de cobertura primaria (Canal {Canal_TV} >= 51 dBuV/m).')
                    worksheet.write('B7',
                                    f'- Color AMARILLO significa que los valores de campo eléctrico diario superan el límite del área de cobertura secundario pero son inferiores al límite del área de cobertura principal establecido (Canal {Canal_TV}: entre 30 y 51 dBuV/m).')
                    worksheet.write('B8',
                                    f'- Color ROJO significa que los valores de campo eléctrico diario son inferiores al límite de área de cobertura secundario. (Canal {Canal_TV}: < 30 dBuV/m).')
                    worksheet.write('B9', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                    worksheet.write('B10', '- Color GRIS: Dispone de autorización para suspensión de emisiones.')
                worksheet.insert_image('A12', 'image.png')

            """Change the name of the file"""
            old_name = 'RTV_Verificación de parámetros.xlsx'
            if Year1 == Year2 and Mes_inicio == Mes_fin:
                new_name = 'TV_Verificación de parámetros_{} (Canal {})_{}_{}_{}.xlsx'.format(nombre, Canal_TV, Ciudad,
                                                                                              Mes_inicio, Year1)
            else:
                new_name = 'TV_Verificación de parámetros_{} (Canal {})_{}_{}{}_{}{}.xlsx'.format(nombre, Canal_TV,
                                                                                                  Ciudad,
                                                                                                  Mes_inicio, Year1,
                                                                                                  Mes_fin,
                                                                                                  Year2)

            """Remove the previous file if already exist"""
            if os.path.exists(f'{download_route}/{new_name}'):
                os.remove(f'{download_route}/{new_name}')

            """Rename the file"""
            os.rename(f'{download_route}/{old_name}',
                      f'{download_route}/{new_name}')

            """Remove the image file"""
            os.remove('image.png')

        elif Ocupacion == False and AM_Reporte_individual == True and FM_Reporte_individual == False and TV_Reporte_individual == False and Autorizaciones == True:
            """REPORTE INDIVIDUAL AM - AUTORIZACIONES"""

            if Ciudad == 'Quito' or Ciudad == 'Guayaquil' or Ciudad == 'Cuenca':
                """Get the data only for the selected "Frecuencia_AM" and save it in the dataframe df17"""
                df17 = df17.loc[df17['Frecuencia (Hz)'] == int(Frecuencia_AM)]

                """Filter dfau1 to get AM frequencies"""
                dfau2 = dfau1[(dfau1.freq >= 570000) & (dfau1.freq <= 1590000)]
                """copy the column 'Fecha_inicio' to a list and save it in col1"""
                col1 = dfau2['Fecha_inicio'].copy().tolist()
                """insert a column in the index 13 with the name 'Fecha_ini' with the col1 information"""
                dfau2.insert(13, 'Fecha_ini', col1)
                dfau2 = dfau2.rename(
                    columns={'freq': 'Frecuencia (Hz)', 'Fecha_inicio': 'Tiempo', 'Fecha_ini': 'Fecha_inicio'})
                dfau2 = dfau2.drop(columns=['est'])
                """Filter dfau2 by 'Frecuencia_AM'"""
                dfau2 = dfau2.loc[dfau2['Frecuencia (Hz)'] == int(Frecuencia_AM)]
                """Filter dfau2 by 'autori' (name of the city)"""
                dfau2 = dfau2.loc[dfau2['ciu'] == autori]

                """Group the information in df17 acording to the requirements for the final report: max Level and
                average Bandwidth per Day"""
                df17 = df17.groupby(by=[
                    pd.Grouper(key='Tiempo', freq='D'),
                    pd.Grouper(key='Frecuencia (Hz)'),
                    pd.Grouper(key='Estación')
                ]).agg({
                    'Level (dBµV/m)': 'max',
                    'Bandwidth (Hz)': 'mean'
                }).reset_index()

                dfau2 = dfau2.drop(columns=['ciu'])
                dfau3 = []
                for index, row in dfau2.iterrows():
                    for t in pd.date_range(start=row['Tiempo'], end=row['Fecha_fin']):
                        dfau3.append(
                            (row['Frecuencia (Hz)'], row['Tipo'], row['Plazo'], t, row['Oficio'], row['Fecha_oficio'],
                             row['Fecha_inicio'], row['Fecha_fin']))
                dfau3 = pd.DataFrame(dfau3, columns=(
                    'Frecuencia (Hz)', 'Tipo', 'Plazo', 'Tiempo', 'Oficio', 'Fecha_oficio', 'Fecha_inicio',
                    'Fecha_fin'))

                """Merge dfau3 with df9 dataframe to add the autorization, df9 is the dataframe that contains all the 
                information for FM"""
                df17 = dfau3.merge(df17, how='right', on=['Tiempo', 'Frecuencia (Hz)'])
                df17 = df17.fillna('-')

                """make the pivot table with the data structured in the way we want to show in the report"""
                df_final7 = pd.pivot_table(df17,
                                           index=[pd.Grouper(key='Tiempo')],
                                           values=['Level (dBµV/m)', 'Bandwidth (Hz)', 'Fecha_fin'],
                                           columns=['Frecuencia (Hz)', 'Estación'],
                                           aggfunc={'Level (dBµV/m)': max, 'Bandwidth (Hz)': np.average,
                                                    'Fecha_fin': max}).round(2)
                df_final7 = df_final7.rename(columns={'Fecha_fin': 'Fin de Autorización'})
                df_final7 = df_final7.T
                df_final8 = df_final7.replace(0, '-')

                """Reset the index (unstack) to have the columns 'Potencia', 'BW Asignado' and 'Tiempo' in the index so 
                we can use the flags in the columns 'Potencia' and 'BW Asignado'"""
                df_final8 = df_final8.reset_index()

                """sorter first by 'Level (dBµV/m)', after by 'Bandwidth (Hz)' and finally 'Fin de Autorización' and
                rename the column header as 'Parámetros'"""
                sorter = ['Level (dBµV/m)', 'Bandwidth (Hz)', 'Fin de Autorización']
                df_final8.level_0 = df_final8.level_0.astype("category")
                df_final8.level_0 = df_final8.level_0.cat.set_categories(sorter)
                df_final8 = df_final8.sort_values(["level_0"])
                df_final8 = df_final8.rename(columns={'level_0': 'Parámetros'})
                df_final8 = df_final8.set_index('Parámetros')

                """PLOT"""
                series = df17
                series = series.groupby(pd.Grouper(key='Tiempo', freq='1D')).max()
                series = series.rename(columns={'Frecuencia (Hz)': 'freq', 'Estación': 'est', 'Level (dBµV/m)': 'level',
                                                'Bandwidth (Hz)': 'bandwidth'})
                series['Fecha_inicio'] = series['Fecha_inicio'].replace('-', 0)
                series['Fecha_fin'] = series['Fecha_fin'].replace('-', 0)
                nombre = series['est'].iloc[0]

                def minus(row):
                    """function to return a specific value if the value in every row of the column 'level' meet the
                    condition"""
                    if row['level'] > 0 and row['level'] < 40:
                        return row['level']
                    return 0

                def bet(row):
                    """function to return a specific value if the value in every row of the column 'level' meet the
                    condition"""
                    if row['level'] >= 40 and row['level'] < 62:
                        return row['level']
                    return 0

                def plus(row):
                    """function to return a specific value if the value in every row of the column 'level' meet the
                    condition"""
                    if row['level'] >= 62:
                        return row['level']
                    return 0

                def valor(row):
                    """function to return a specific value if the value in every row of the column 'level' meet the
                    condition"""
                    if row['level'] == 0:
                        return 120
                    return 0

                def aut(row):
                    """function to return a specific value if the value in every row of the columns 'level' and
                    'Fecha_fin' meet the condition"""
                    if row['Fecha_fin'] != 0 and row['level'] == 0:
                        return 0
                    elif row['Fecha_fin'] != 0 and row['level'] != 0:
                        return row['level']
                    return 0

                """create a new column in the series frame for every definition (minus, bet, plus, valor)"""
                series['minus'] = series.apply(lambda row: minus(row), axis=1)
                series['bet'] = series.apply(lambda row: bet(row), axis=1)
                series['plus'] = series.apply(lambda row: plus(row), axis=1)
                series['valor'] = series.apply(lambda row: valor(row), axis=1)
                series['aut'] = series.apply(lambda row: aut(row), axis=1)

                fig = plt.figure()
                """111 means 1x1 grid, first subplot, so it will put every subplot in the same figure"""
                ax = fig.add_subplot(111, ylabel='Nivel de Intensidad de Campo Eléctrico (dBµV/m)', xlabel='Tiempo',
                                     title=f'Ciudad: {Ciudad}, Estación: {nombre}, Frecuencia: {Frecuencia_AM} Hz')
                """set the y-axis limits"""
                ax.set_ylim(0, 120)
                """area plot of every definition"""
                series['plus'].plot.area(y=ax, colormap='Accent', linewidth=0)
                series['bet'].plot.area(y=ax, colormap='Set3_r', linewidth=0)
                series['minus'].plot.area(y=ax, colormap='Pastel1', linewidth=0)
                if Seleccionar == True:
                    series['valor'].plot.area(y=ax, colormap='Paired', linewidth=0)
                elif Seleccionar == False:
                    try:
                        series['level'][series.level == 0].plot(y=ax, marker='v', markersize=7, color='b', linewidth=0)
                    except ValueError:  # raise if array is empty (no cero values for level column)
                        pass
                if Autorizaciones == True:
                    series['aut'].plot.area(y=ax, colormap='Set2_r', linewidth=0)

                """Annotations for initial and final dates of the authorization"""
                ini_date = series.Fecha_inicio.tolist()
                fin_date = series.Fecha_fin.tolist()

                for mark_time1 in ini_date:
                    time_x1 = pd.Timestamp(mark_time1)  # convert to compatible format
                    time_x1 = time_x1.strftime('%Y-%m-%d')
                    ax.annotate(f'Inicio: {time_x1}', xy=(time_x1, 0), xycoords='data',
                                bbox=dict(boxstyle="round", fc="white", ec="black"),
                                xytext=(time_x1, 12), ha='center',
                                arrowprops=dict(arrowstyle="->",
                                                connectionstyle="angle, angleA = 0, angleB = 90, rad = 10")
                                )

                for mark_time2 in fin_date:
                    time_x2 = pd.Timestamp(mark_time2)  # convert to compatible format
                    time_x2 = time_x2.strftime('%Y-%m-%d')
                    ax.annotate(f'Fin: {time_x2}', xy=(time_x2, 0), xycoords='data',
                                bbox=dict(boxstyle="round", fc="white", ec="black"),
                                xytext=(time_x2, 6), ha='center',
                                arrowprops=dict(arrowstyle="->",
                                                connectionstyle="angle, angleA = 0, angleB = 90, rad = 10")
                                )

                plt.grid(color='#292929', linestyle='--', linewidth=0.5)
                """save the plot in the file image.png"""
                plt.savefig('image.png')
                """close the plot to not be show after the code is executed"""
                plt.close()

                """REPORT CREATION: REPORTE INDIVIDUAL AM - AUTORIZACIONES"""
                """create, write and save"""
                with pd.ExcelWriter(f'{download_route}/RTV_Verificación de parámetros.xlsx') as writer:
                    df_final8.to_excel(writer, sheet_name=f'Radiodifusión AM_{Frecuencia_AM} Hz')
                    dfau2.set_index('Frecuencia (Hz)').rename(
                        columns={'Fecha_inicio': 'Inicio Autorización', 'Fecha_fin': 'Fin Autorización'}).drop(
                        columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).to_excel(writer,
                                                                                             sheet_name='Autorizaciones')

                    """Get the xlsxwriter workbook and worksheet objects."""
                    workbook = writer.book
                    worksheet = writer.sheets[f'Radiodifusión AM_{Frecuencia_AM} Hz']
                    worksheet1 = writer.sheets['Autorizaciones']

                    """Add a format."""
                    format1 = workbook.add_format({'bg_color': '#C6EFCE',
                                                   'font_color': '#006100'})

                    format2 = workbook.add_format({'bg_color': '#FFC7CE',
                                                   'font_color': '#9C0006'})

                    format3 = workbook.add_format({'bg_color': '#FFDC47',
                                                   'font_color': '#9C6500'})

                    format4 = workbook.add_format({'num_format': 'dd/mm/yy'})

                    format5 = workbook.add_format({'border': 1, 'border_color': 'black'})

                    format6 = workbook.add_format({'bg_color': '#99CCFF',
                                                   'font_color': '#0066FF'})

                    format7 = workbook.add_format({'bg_color': '#C0C0C0',
                                                   'font_color': '#000000'})

                    """Get the dimensions of the dataframe."""
                    (max_row1, max_col1) = dfau2.set_index('Frecuencia (Hz)').rename(
                        columns={'Fecha_inicio': 'Inicio Autorización', 'Fecha_fin': 'Fin Autorización'}).drop(
                        columns=['DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'Tiempo']).shape

                    """Apply a conditional format to the required cell range."""
                    worksheet1.conditional_format(1, 1, int(max_row1), int(max_col1),
                                                  {'type': 'no_errors',
                                                   'format': format5})
                    worksheet1.autofilter(0, 0, 0, int(max_col1))

                    """Get the dimensions of the dataframe."""
                    (max_row, max_col) = df_final8.shape

                    """Apply a conditional format to the required cell range."""
                    worksheet.conditional_format(0, 3, 0, int(max_col),
                                                 {'type': 'no_errors',
                                                  'format': format4})
                    worksheet.conditional_format(0, 1, int(max_row), int(max_col),
                                                 {'type': 'no_errors',
                                                  'format': format5})
                    worksheet.conditional_format(3, 3, 3, int(max_col),
                                                 {'type': 'no_errors',
                                                  'format': format4})
                    worksheet.autofilter(0, 0, 0, int(max_col))

                    worksheet.conditional_format(1, 3, 1, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D3="-")',
                                                  'format': format6})
                    worksheet.conditional_format(1, 3, 1, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D3<40)',
                                                  'format': format2})
                    worksheet.conditional_format(1, 3, 1, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D3>=62)',
                                                  'format': format1})
                    worksheet.conditional_format(1, 3, 1, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D3>=40,F3<62)',
                                                  'format': format3})

                    worksheet.conditional_format(2, 3, 2, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D2="-")',
                                                  'format': format6})
                    worksheet.conditional_format(2, 3, 2, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D2<=15000)',
                                                  'format': format1})
                    worksheet.conditional_format(2, 3, 2, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D2>15000)',
                                                  'format': format2})

                    worksheet.conditional_format(3, 3, 3, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D4="-")',
                                                  'format': format6})
                    worksheet.conditional_format(3, 3, 3, int(max_col),
                                                 {'type': 'formula',
                                                  'criteria': '=AND(D4<>"-")',
                                                  'format': format7})

                    worksheet.conditional_format('A7', {'type': 'no_errors',
                                                        'format': format1})
                    worksheet.conditional_format('A8', {'type': 'no_errors',
                                                        'format': format3})
                    worksheet.conditional_format('A9', {'type': 'no_errors',
                                                        'format': format2})
                    worksheet.conditional_format('A10', {'type': 'no_errors',
                                                         'format': format6})
                    worksheet.conditional_format('A11', {'type': 'no_errors',
                                                         'format': format7})

                    worksheet.write('B6',
                                    '- Los valores de intensidad de campo eléctrico en dBuV/m corresponden a los máximos diarios.')
                    worksheet.write('B7',
                                    '- Color VERDE: los valores de campo eléctrico diario superan el valor del borde de área de cobertura (>=62 dBuV/m).')
                    worksheet.write('B8',
                                    '- Color AMARILLO: los valores de campo eléctrico diario se encuentran entre el valor del borde de área de protección y el valor del borde de área de cobertura (entre 40 y 62 dBuV/m).')
                    worksheet.write('B9',
                                    '- Color ROJO: los valores de campo eléctrico diario son inferiores al valor del borde de área de protección (<40 dBuV/m).')
                    worksheet.write('B10', '- Color AZUL: No se dispone de mediciones del sistema SACER.')
                    worksheet.write('B11', '- Color GRIS: Dispone de autorización para suspensión de emisiones.')
                    worksheet.write('B12', '- El valor de ancho de banda corresponde a 15 kHz.')
                    worksheet.insert_image('A14', 'image.png')

                """Change the name of the file"""
                old_name = 'RTV_Verificación de parámetros.xlsx'
                if Year1 == Year2 and Mes_inicio == Mes_fin:
                    new_name = 'AM_Verificación de parámetros_{} ({} Hz)_{}_{}_{}.xlsx'.format(nombre, Frecuencia_AM,
                                                                                               Ciudad,
                                                                                               Mes_inicio, Year1)
                else:
                    new_name = 'AM_Verificación de parámetros_{} ({} Hz)_{}_{}{}_{}{}.xlsx'.format(nombre,
                                                                                                   Frecuencia_AM,
                                                                                                   Ciudad, Mes_inicio,
                                                                                                   Year1,
                                                                                                   Mes_fin, Year2)
                """Remove the previous file if already exist"""
                if os.path.exists(f'{download_route}/{new_name}'):
                    os.remove(f'{download_route}/{new_name}')

                """Rename the file"""
                os.rename(f'{download_route}/{old_name}',
                          f'{download_route}/{new_name}')

                """Remove the image file"""
                os.remove('image.png')

        elif Ocupacion == True and (
                AM_Reporte_individual == True or FM_Reporte_individual == True or TV_Reporte_individual == True):
            print(
                "La Ocupacion está habilitada únicamente para el reporte general. Desmarque las opciones de Reporte Individual.")

        elif AM_Reporte_individual == True and (Ciudad != 'Quito' or Ciudad != 'Guayaquil' or Ciudad != 'Cuenca'):
            print(f"No se dispone de mediciones de Radiodifusión AM para la ciudad de {Ciudad}.")

        elif Ocupacion == False and (
                (AM_Reporte_individual == True and FM_Reporte_individual == True and TV_Reporte_individual == True) or (
                AM_Reporte_individual == False and FM_Reporte_individual == True and TV_Reporte_individual == True) or (
                        AM_Reporte_individual == True and FM_Reporte_individual == False and TV_Reporte_individual == True) or (
                        AM_Reporte_individual == True and FM_Reporte_individual == True and TV_Reporte_individual == False)):
            print("Solo se puede generar un Reporte Individual. Marque solo una opción de reporte individual")


if __name__ == "__main__":
    """Creates a new instance of the Tkinter class, which represents a main window or root window of
    a GUI application. The variable root is commonly used to reference this root window throughout the program."""
    root = tk.Tk()

    """Creates an object of the SacerApp class with the root window. The SacerApp class is a custom class defined in the
    code, and the master is the root window that serves as the parent window for any widgets or components that the
    SacerApp class will create or manage."""
    app = SacerApp(master=root)

    """app.grid() is a method used to configure the grid layout of widgets in a tkinter GUI (Graphical User Interface)
    application. The method arranges widgets in rows and columns, allowing them to be aligned and spaced out evenly."""
    app.grid()

    """Define start_button command execution"""
    start_button = tk.Button(root, text="Iniciar", command=app.start)
    start_button.grid(row=12, column=0, sticky=tk.W, padx=30, pady=5)

    """Define quit_button command execution"""
    quit_button = tk.Button(root, text="Cerrar", command=app.quit)
    quit_button.grid(row=13, column=0, sticky=tk.W, padx=30, pady=5)

    """When root.mainloop() is called, it starts the event loop of the Tkinter application and waits for user input or
    any event-driven inputs to occur. This method is essential in making sure that the GUI remains responsive and
    interactive."""
    root.mainloop()
