import os
import tkinter as tk
import tkcalendar as tc
from tkinter import messagebox
import pandas as pd
import sqlalchemy as sa
import datetime
from sqlalchemy.engine import URL
from sqlalchemy import create_engine
from dotenv import load_dotenv

# load the environment variables
load_dotenv()


class SacerServerApp(tk.Frame):
    """Create the tkinter application class"""

    def __init__(self, master=None):
        """This is a constructor method for a class in Python. It initializes the object's internal state and is
        automatically called when an object is created."""

        # Use the built-in constructor method of the tkinter module in Python. The super() function is used to call the
        # constructor of the parent class of the current class, and __init__() is the method that is called when creating
        # a new instance of the class. In this case, master is the parameter being passed to the constructor method, which
        # is the main window object that's being created. This line of code is used to initialize an instance of a tkinter
        # widget or frame.
        super().__init__(master)

        # Initialize all the variables
        self.master = master
        self.master.title("Reportes Servidor SACER ")
        self.master.geometry("800x450")
        self.list_of_operators = ["scn-l01", "scn-l02", "scn-l03", "scn-l05", "scn-l06"]
        self.repgen = tk.BooleanVar()
        self.mapa = tk.BooleanVar()
        self.operadora = tk.StringVar()
        self.operadora.set("scn-l01")
        self.create_widgets()
        self.program_is_running = False

    def create_widgets(self):
        """Create the widges to be used with tkinter and tkcalendar"""

        self.button0 = tk.Checkbutton(self.master, text="Reporte FM.", variable=self.repgen,
                                      command=lambda: self.toggle_button0_state())
        self.button0.grid(row=0, column=0, sticky=tk.W, padx=30, pady=5)

        self.button1 = tk.Checkbutton(self.master, text="Reporte TV.", variable=self.mapa,
                                      command=lambda: self.toggle_button1_state())
        self.button1.grid(row=0, column=0, sticky=tk.W, padx=150, pady=5)

        self.lbl_1 = tk.Label(self.master, text="ERT:", width=20, font=("bold", 11))
        self.lbl_1.grid(row=1, column=0, sticky=tk.W, padx=13)

        self.option_menu1 = tk.OptionMenu(self.master, self.operadora, *self.list_of_operators)
        self.option_menu1.grid(row=1, column=0, sticky=tk.W, padx=145)

        self.lbl_2 = tk.Label(self.master, text="Fecha inicio:", width=10, font=("bold", 11))
        self.lbl_2.grid(row=1, column=0, sticky=tk.W, padx=285)

        self.fecha_inicio = tc.DateEntry(self.master, selectmode='day', date_pattern='yyyy-mm-dd')
        self.fecha_inicio.grid(row=1, column=0, sticky=tk.W, padx=380)

        self.lbl_3 = tk.Label(self.master, text="Fecha fin:", width=10, font=("bold", 11))
        self.lbl_3.grid(row=1, column=0, sticky=tk.W, padx=475)

        self.fecha_fin = tc.DateEntry(self.master, selectmode='day', date_pattern='yyyy-mm-dd')
        self.fecha_fin.grid(row=1, column=0, sticky=tk.W, padx=570)

    def toggle_button0_state(self):
        """Define button0 state conditions"""
        if self.repgen.get():
            self.button1.config(state=tk.DISABLED)
        else:
            self.button1.config(state=tk.NORMAL)

    def toggle_button1_state(self):
        """Define button1 state conditions"""
        if self.mapa.get():
            self.button0.config(state=tk.DISABLED)
        else:
            self.button0.config(state=tk.NORMAL)

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
        ert = self.operadora.get()
        fecha_inicio = self.fecha_inicio.get_date().strftime("%Y-%m-%d")
        fecha_fin = self.fecha_fin.get_date().strftime("%Y-%m-%d")

        # Add HH:MM:SS to fecha_inicio and fecha_fin so the range of dates we want to show in the report is correct
        add_string1 = ' 00:00:01'
        add_string2 = ' 23:59:59'
        fecha_inicio += add_string1
        fecha_fin += add_string2

        def convert(date_time):
            """function to convert a string to datetime object"""
            date_format = '%Y-%m-%d %H:%M:%S'  # The format
            datetime_str = datetime.datetime.strptime(date_time, date_format)
            return datetime_str

        # Convert fecha_inicio and fecha_fin to datetime object
        fecha_inicio = convert(fecha_inicio)
        fecha_fin = convert(fecha_fin)

        # Define the SQL querys to be executed
        if self.repgen.get():
            # execute a test query and pass the results to a pandas dataframe
            query1 = f"SELECT vargusresultextended.meaunitname, " \
                     f"vargusresultextended.firstmeastime, " \
                     f"rvecu_technical.frequency, " \
                     f"rvecu_technical.level, " \
                     f"rvecu_technical.bandwidth, " \
                     f"rvecu_technical.freqoffset, " \
                     f"vargusresultextended.demodulation, " \
                     f"rvecu_technical.resultname, rvecu_technical.argusresult_id, vargusresultextended.argusmeaunit_id " \
                     f"FROM cenargus.rvecu_technical " \
                     f"JOIN cenargus.vargusresultextended ON rvecu_technical.argusresult_id = vargusresultextended.id " \
                     f"JOIN cenargus.vargusmeaunit ON vargusmeaunit.id = vargusresultextended.argusmeaunit_id " \
                     f"WHERE rvecu_technical.level > -10000 AND " \
                     f"vargusresultextended.firstmeastime >= '{fecha_inicio}' AND " \
                     f"vargusresultextended.firstmeastime <= '{fecha_fin}' AND " \
                     f"vargusresultextended.meaunitname like '%{ert}%' AND " \
                     f"vargusresultextended.detector IS NOT NULL AND " \
                     f"vargusresultextended.demodulation IS NOT NULL AND " \
                     f"rvecu_technical.frequency between '88000000' and '108000000'"
        elif self.mapa.get():
            # execute a test query and pass the results to a pandas dataframe
            query1 = f"SELECT vargusresultextended.meaunitname, " \
                     f"vargusresultextended.firstmeastime, " \
                     f"rvecu_technical.frequency, " \
                     f"rvecu_technical.level, " \
                     f"rvecu_technical.bandwidth, " \
                     f"rvecu_technical.freqoffset, " \
                     f"vargusresultextended.demodulation," \
                     f"rvecu_technical.resultname, rvecu_technical.argusresult_id, vargusresultextended.argusmeaunit_id " \
                     f"FROM cenargus.rvecu_technical " \
                     f"JOIN cenargus.vargusresultextended ON rvecu_technical.argusresult_id = vargusresultextended.id " \
                     f"JOIN cenargus.vargusmeaunit ON vargusmeaunit.id = vargusresultextended.argusmeaunit_id " \
                     f"WHERE rvecu_technical.level > -10000 AND " \
                     f"vargusresultextended.firstmeastime >= '{fecha_inicio}' AND " \
                     f"vargusresultextended.firstmeastime <= '{fecha_fin}' AND " \
                     f"vargusresultextended.meaunitname like '%{ert}%' AND " \
                     f"vargusresultextended.detector IS NOT NULL AND " \
                     f"vargusresultextended.demodulation IS NOT NULL AND " \
                     f"rvecu_technical.frequency between '54000000'and '88000000' " \
                     f"UNION " \
                     f"SELECT vargusresultextended.meaunitname, " \
                     f"vargusresultextended.firstmeastime, " \
                     f"rvecu_technical.frequency, " \
                     f"rvecu_technical.level, " \
                     f"rvecu_technical.bandwidth, " \
                     f"rvecu_technical.freqoffset, " \
                     f"vargusresultextended.demodulation," \
                     f"rvecu_technical.resultname, rvecu_technical.argusresult_id, vargusresultextended.argusmeaunit_id " \
                     f"FROM cenargus.rvecu_technical " \
                     f"JOIN cenargus.vargusresultextended ON rvecu_technical.argusresult_id = vargusresultextended.id " \
                     f"JOIN cenargus.vargusmeaunit ON vargusmeaunit.id = vargusresultextended.argusmeaunit_id " \
                     f"WHERE rvecu_technical.level > -10000 AND " \
                     f"vargusresultextended.firstmeastime >= '{fecha_inicio}' AND " \
                     f"vargusresultextended.firstmeastime <= '{fecha_fin}' AND " \
                     f"vargusresultextended.meaunitname like '%{ert}%' AND " \
                     f"vargusresultextended.detector IS NOT NULL AND " \
                     f"vargusresultextended.demodulation IS NOT NULL AND " \
                     f"rvecu_technical.frequency between '174000000'and '216000000' " \
                     f"UNION " \
                     f"SELECT vargusresultextended.meaunitname, " \
                     f"vargusresultextended.firstmeastime, " \
                     f"rvecu_technical.frequency, " \
                     f"rvecu_technical.level, " \
                     f"rvecu_technical.bandwidth, " \
                     f"rvecu_technical.freqoffset, " \
                     f"vargusresultextended.demodulation," \
                     f"rvecu_technical.resultname, rvecu_technical.argusresult_id, vargusresultextended.argusmeaunit_id " \
                     f"FROM cenargus.rvecu_technical " \
                     f"JOIN cenargus.vargusresultextended ON rvecu_technical.argusresult_id = vargusresultextended.id " \
                     f"JOIN cenargus.vargusmeaunit ON vargusmeaunit.id = vargusresultextended.argusmeaunit_id " \
                     f"WHERE rvecu_technical.level > -10000 AND " \
                     f"vargusresultextended.firstmeastime >= '{fecha_inicio}' AND " \
                     f"vargusresultextended.firstmeastime <= '{fecha_fin}' AND " \
                     f"vargusresultextended.meaunitname like '%{ert}%' AND " \
                     f"vargusresultextended.detector IS NOT NULL AND " \
                     f"vargusresultextended.demodulation IS NOT NULL AND " \
                     f"rvecu_technical.frequency between '470000000'and '698000000'"

        # create the database URI with SSPI authentication method
        url_object = URL.create(
            "postgresql+psycopg2",
            username=os.getenv('USER_NAME'),
            password=os.getenv('USER_PASSWORD'),
            host=os.getenv('SERVER_NAME'),
            database=os.getenv('DATABASE_NAME'),
        )

        # create the engine object using create_engine
        engine = create_engine(url_object)

        # Define decimal separator as comma
        decimal_sep = ","

        # Establish a connection to the database, execute the query and create the df with the results
        with engine.begin() as conn:
            df = pd.read_sql_query(sa.text(query1), conn, params={"decimal": decimal_sep})

        # Close the connection object
        conn.close()

        fecha_inicio1 = fecha_inicio.strftime('%Y-%m-%d')
        fecha_fin1 = fecha_fin.strftime('%Y-%m-%d')
        if self.repgen.get():
            # Pass the df to a .csv file
            df.to_csv(f'FM_{ert}_{fecha_inicio1}_{fecha_fin1}.csv', index=False, sep=';',
                      encoding='utf-8',
                      header=True, decimal=',')

            # Remove the previous files if already exist
            if os.path.exists(
                    f"{os.getenv('download_route')}/FM_{ert}_{fecha_inicio1}_{fecha_fin1}.csv"):
                os.remove(
                    f"{os.getenv('download_route')}/FM_{ert}_{fecha_inicio1}_{fecha_fin1}.csv")

            # Download the .csv file
            os.rename(f"FM_{ert}_{fecha_inicio1}_{fecha_fin1}.csv",
                      f"{os.getenv('download_route')}/FM_{ert}_{fecha_inicio1}_{fecha_fin1}.csv")

        elif self.mapa.get():
            # Pass the df to a .csv file
            df.to_csv(f'TV_{ert}_{fecha_inicio1}_{fecha_fin1}.csv', index=False, sep=';',
                      encoding='utf-8',
                      header=True, decimal=',')

            # Remove the previous files if already exist
            if os.path.exists(
                    f"{os.getenv('download_route')}/TV_{ert}_{fecha_inicio1}_{fecha_fin1}.csv"):
                os.remove(
                    f"{os.getenv('download_route')}/TV_{ert}_{fecha_inicio1}_{fecha_fin1}.csv")

            # Download the .csv file
            os.rename(f"TV_{ert}_{fecha_inicio1}_{fecha_fin1}.csv",
                      f"{os.getenv('download_route')}/TV_{ert}_{fecha_inicio1}_{fecha_fin1}.csv")


if __name__ == "__main__":
    # Creates a new instance of the Tkinter class, which represents a main window or root window of
    # a GUI application. The variable root is commonly used to reference this root window throughout the program.
    root = tk.Tk()

    # Creates an object of the SammApp class with the root window. The SammApp class is a custom class defined in the
    # code, and the master is the root window that serves as the parent window for any widgets or components that the
    # SammApp class will create or manage.
    app = SacerServerApp(master=root)

    # app.grid() is a method used to configure the grid layout of widgets in a tkinter GUI (Graphical User Interface)
    # application. The method arranges widgets in rows and columns, allowing them to be aligned and spaced out evenly.
    app.grid()

    # Define start_button command execution
    start_button = tk.Button(root, text="Iniciar", command=app.start)
    start_button.grid(row=12, column=0, sticky=tk.W, padx=30, pady=5)

    # Define quit_button command execution
    quit_button = tk.Button(root, text="Cerrar", command=app.quit)
    quit_button.grid(row=13, column=0, sticky=tk.W, padx=30, pady=5)

    # When root.mainloop() is called, it starts the event loop of the Tkinter application and waits for user input or
    # any event-driven inputs to occur. This method is essential in making sure that the GUI remains responsive and
    # interactive.
    root.mainloop()
