"""
EGM Wizard for use with
EGM-4 from PP Systems

Copyright: Ryan Anderson 2017
"""

from Tkinter import Tk, BOTH, Listbox, Label, Menu, Toplevel, Entry, Radiobutton, IntVar, Button, Frame, Canvas, Scrollbar, DoubleVar
from ttk import Style
import sqlite3
import time
import xlwt
import io
import tkFileDialog
import threading
import tkMessageBox
import tkSimpleDialog
import numpy
import winsound
import tempfile
import base64
import zlib
import shelve
from serial import Serial, SerialException
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

searchState = "stop"


root = Tk()
DeftColor = root.cget("bg")


class EGM(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent  # parent is root, reference to parent widget
        self.initUI()  # creation of user interface
        self.parent.bind('f', self.toggleStart)
        self.parent.bind('s', self.toggleStop)
        self.parent.bind('r', self.recordNow)
        self.parent.bind('c', self.deleteButton)
        self.peakFlag = 0

        print("init")

    def initUI(self):  # The user interface
        # ------------------------------------------FRAMES
        fTop = Frame(root, bg='SlateGray2')
        fTop.grid(row=0, column=0, columnspan=7, pady=12)

        self.canvas = Canvas(root)
        self.vsb = Scrollbar(root, orient="vertical",
                             command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set,
                              height=500, width=420)

        self.canvas.grid(row=2, column=0, rowspan=8, columnspan=7)
        self.vsb.grid(row=2, column=8, sticky="ns", rowspan=8)

        # ------------------------------------------BUTTONS
        self.startButton = Button(
            fTop, text="Find", width=12, command=self.toggleStart)  # Find Peaks button
        self.startButton.pack(side='left', padx=6, pady=3)
        self.stopButton = Button(fTop, text="Stop", width=12, relief='sunken', background='red', command=self.toggleStop,
                                 state='disabled', disabledforeground='black')  # command=self.stopPeak #Stop Finding button
        self.stopButton.pack(side='left', padx=6, pady=3)
        findPeaksButton = Button(
            fTop, text="Record Now", command=self.recordNow)  # Record Now button
        findPeaksButton.pack(side='left', padx=6, pady=3)
        deleteButton = Button(fTop, text="Delete Last Row",
                              command=self.deleteButton)  # Clear last row button
        deleteButton.pack(side='left', padx=6, pady=3)

        # ------------------------------------------LABELS
        # X labels
        label1 = Label(root, text="ID", bg='SlateGray2')
        label1.grid(row=1, column=1)
        label2 = Label(root, text="Peak 1", bg='SlateGray2')
        label2.grid(row=1, column=2)
        label3 = Label(root, text="Peak 2", bg='SlateGray2')
        label3.grid(row=1, column=3)
        label4 = Label(root, text="Peak 3", bg='SlateGray2')
        label4.grid(row=1, column=4)
        label5 = Label(root, text="Peak 4", bg='SlateGray2')
        label5.grid(row=1, column=5)
        label6 = Label(root, text="Peak 5", bg='SlateGray2')
        label6.grid(row=1, column=6)

        labelSpace = Label(root, text=" ", bg='SlateGray2')
        labelSpace.grid(row=1, column=0)
        root.grid_columnconfigure(0, minsize=40)

        # ----------------------------------------- MENU BAR
        menubar = Menu(root)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="New", command=self.newButton)
        filemenu.add_command(label="Open", command=self.openButton)
        filemenu.add_command(label="Export to Excel",
                             command=lambda: export(self.crntDBmanager))
        filemenu.add_command(label="Exit", command=root.quit)
        menubar.add_cascade(label="File", menu=filemenu)

        editmenu = Menu(menubar, tearoff=0)
        editmenu.add_command(
            label="Settings", command=lambda: settingsDialog(root))
        menubar.add_cascade(label="Edit", menu=editmenu)

        viewmenu = Menu(menubar, tearoff=0)
        viewmenu.add_command(label="First Curve Plot",
                             command=lambda: recoveryDialog(self))
        menubar.add_cascade(label="View", menu=viewmenu)

        root.config(menu=menubar)

        # ----------------------------------------- MAIN FRAME
        root.title("EMG-4 Wizard")  # title
        root.style = Style()
        root.style.theme_use('classic')  # Theme

    def makefMain(self):
        self.fMain = Frame(self.canvas)
        self.fMain.pack()

        label7 = Label(self.fMain, text="0 mM")
        label7.grid(row=0, column=0)
        label8 = Label(self.fMain, text=".25 mM")
        label8.grid(row=1, column=0)
        label9 = Label(self.fMain, text=".5 mM")
        label9.grid(row=2, column=0)
        label10 = Label(self.fMain, text="1 mM")
        label10.grid(row=3, column=0)
        label11 = Label(self.fMain, text="2 mM")
        label11.grid(row=4, column=0)
        label12 = Label(self.fMain, text="4 mM")
        label12.grid(row=5, column=0)

        self.canvas.create_window(
            (4, 4), window=self.fMain, anchor="nw", tags="self.frame",)
        self.fMain.bind("<Configure>", self.onFrameConfigure)

    def onFrameConfigure(self, event):
        '''Reset the scroll region to encompass the inner frame'''
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def toggleStart(self, d=0):
        global searchState
        searchState = "searching"
        self.startPeak()
        self.startButton.config(
            relief="sunken", background='green', state='disabled', disabledforeground='black')
        self.stopButton.config(
            relief='raised', background=DeftColor, state='active')

    def toggleStop(self, d=0):
        global searchState
        searchState = "stop"
        self.stopPeak()
        self.stopButton.config(
            relief='sunken', background='red', state='disabled', disabledforeground='black')
        self.startButton.config(
            relief='raised', background=DeftColor, state='active')

    def plotButton(self):
        crntPlot = plot()
        crntPlot.graph()

    def settingsButton(self):
        crntSettingDiolog = settingsDialog(root)

    def newButton(self):  # called by File --> New
        self.crntDBmanager = DBmanager()  # create DB object
        self.crntDBmanager.new()  # create new DB connection/file
        print(self.crntDBmanager)
        self.update()

    def openButton(self):  # called by File --> open
        self.crntDBmanager = DBmanager()
        self.crntDBmanager.openFile()
        self.update()

    def update(self):  # updates GUI with latest data from the DB

        print("updating...")
        dataUpdate = self.crntDBmanager.get()  # get's latest data

        rows = self.crntDBmanager.lastRow()  # get's number of rows

        try:  # clear spreadsheet
            self.fMain.destroy()
            self.makefMain()
        except AttributeError as exception:
            self.makefMain()

        height = 40
        width = 6
        for i in range(height):  # Rows
            for j in range(width):  # Columns
                listbox = Listbox(self.fMain, height=1, width=10)
                listbox.grid(row=i, column=j+1)
                if i < rows[0][0]:  # Check db if there is an entry
                    listbox.insert(1, dataUpdate[i][j])  # fill cells

    def startPeak(self, z=0):  # connection to start new thread to find peaks4
        d = WaitDialog(self.parent, self)  # calls WaitDialog
        root.wait_window(d.top)
        self.parent.crntSerial.openPort()
        self.loopPeak()

    def loopPeak(self):
        if self.peakFlag == 0:
            self.crntSearch.run()
            self.parent.after(300, self.loopPeak)

    def stopPeak(self, z=0):  # connection to stop
        print("STOPPPPPPPPPPP")
        self.parent.crntSerial.closePort()
        self.peakFlag = 1

    def recordNow(self, z=0):  # manualy record data point NOW
        self.toggleStop()
        root.crntSerial.openPort()
        currentRead = root.crntSerial.readLine()  # 'M000000000000000003400000'
        root.crntSerial.closePort()
        if currentRead[0] == 'M':
            read = currentRead[16:20]
            rows = self.crntDBmanager.lastRow()[0][0]  # find last row
            lastRow = self.crntDBmanager.get()[rows-1]

            for i in range(6):
                if i > 0:
                    if lastRow[i] == 0:
                        self.crntDBmanager.update('sample'+str(i), read, rows)
                        break
            self.update()  # updates graphic interface

    def deleteButton(self, z=0):  # clear row connection
        if searchState == "stop":
            self.crntDBmanager.deleteLastRow()
        if searchState == "searching":
            self.crntDBmanager.clearLastRow()
            self.crntSearch.count = 1

        self.update()


# -------------------------------------------------Dialogs
class settingsDialog:
    def __init__(self, parent):
        top = self.top = Toplevel(parent)
        self.parent = parent

        self.selected = IntVar()  # tkinter variable BS
        self.selected.set(self.parent.crntSetting.com)  # get current value
        Radiobutton(top, text="COM1", variable=self.selected,
                    value=0).grid(row=0, column=0, padx=10)
        Radiobutton(top, text="COM2", variable=self.selected,
                    value=1).grid(row=1, column=0)
        Radiobutton(top, text="COM3", variable=self.selected,
                    value=2).grid(row=2, column=0)
        Radiobutton(top, text="COM4", variable=self.selected,
                    value=3).grid(row=3, column=0)
        Radiobutton(top, text="COM5", variable=self.selected,
                    value=4).grid(row=4, column=0)
        Radiobutton(top, text="COM6", variable=self.selected,
                    value=5).grid(row=5, column=0)
        Radiobutton(top, text="COM7", variable=self.selected,
                    value=6).grid(row=6, column=0)
        Radiobutton(top, text="COM8", variable=self.selected,
                    value=7).grid(row=7, column=0)

        Label(top, text="Dickson Batch #:").grid(row=0, column=1)

        crntBatch = root.crntSetting.readBatchNum()
        self.batch = IntVar(root, value=crntBatch)
        Entry(top, width=10, textvariable=self.batch).grid(row=1, column=1)

        Label(top, text="DIC Value:").grid(row=2, column=1)
        crntDIC = root.crntSetting.readDIC()
        self.DIC = DoubleVar(root, value=crntDIC)
        Entry(top, width=10, textvariable=self.DIC).grid(row=3, column=1)

        Button(top, text="OK", command=self.close).grid(
            row=8, column=0, columnspan=2)

        top.geometry("200x230+200+300")

    def close(self):
        self.parent.crntSetting.com = self.selected.get()
        crtBatch = self.batch.get()
        crtDIC = self.DIC.get()
        root.crntSetting.setDickson(crtDIC, crtBatch)

        self.top.destroy()
        print(self.parent.crntSetting.com)


class WaitDialog:  # pop up window that waits for EGM stabilization
    def __init__(self, grandparent, parent):
        self.parent = parent  # parent is app
        self.grandparent = grandparent  # grandparent is root
        top = self.top = Toplevel(grandparent)
        Label(top, text="45 Second delay for flow and stabilization").pack(pady=20)
        grandparent.after(100, self.beUseful)
        grandparent.after(45000, self.close)  # Wait then close dialog
        top.geometry("300x100+300+300")
        top.grab_set()  # deactivates main window

    def close(self):  # closes window
        self.top.grab_release()  # reactivates main window
        self.top.destroy()

    def beUseful(self):
        # Do this while waiting
        self.parent.crntDBmanager.newRow()  # new row in DB
        self.parent.crntSearch = findPeaks(
            self.parent.crntDBmanager, self.parent)  # search instance
        self.parent.peakFlag = 0  # flag for stop button
        self.parent.crntSearch.resetReads()  # Set read to 9999


class startupDialog:  # Confirming EGM is connected and in REC mode
    def __init__(self, parent):
        top = self.top = Toplevel(parent)
        top.wm_transient(parent)

        Label(top, text="Is EMG connected and in REC mode?").pack(pady=5)
        b = Button(top, text="YES", command=self.close)
        b.pack(pady=5)

        top.geometry("300x100+200+300")
        top.grab_set()  # deactivates main window

    def close(self):  # closes window
        self.top.grab_release()  # reactivates main window
        self.top.destroy()


class found5peaks:
    def __init__(self, d=0):
        top = self.top = Toplevel(root)
        Label(top, text="5 Peaks Found").pack()
        Label(top, text="Procede to next sample").pack()
        Button(top, text="OK", command=self.closeFound).pack()
        top.geometry("100x100+300+220")
        top.grab_set()
        root.bind('<Return>', self.closeFound)

        self.soundFlag = 1
        self.playSound()

    def playSound(self):  # Play F to get attention
        if self.soundFlag == 1:
            winsound.PlaySound(
                'SystemQuestion', winsound.SND_ALIAS | winsound.SND_ASYNC)
            root.after(10000, self.playSound)

    def closeFound(self, d=0):
        self.soundFlag = 0
        self.top.grab_release()
        self.top.destroy()


class found4peaks:
    def __init__(self, grandparent, d=0):
        self.grandparent = grandparent  # parent is app
        top = self.top = Toplevel(root)
        Label(top, text="4 Peaks Found").pack()
        Button(top, text="Stop Search", command=self.closeFound).pack()
        Button(top, text="Continue Search", command=self.continueSearch).pack()
        top.geometry("100x100+300+220")
        top.grab_set()

        self.soundFlag = 1
        self.playSound()

    def playSound(self):  # play sound to get attention
        if self.soundFlag == 1:
            winsound.PlaySound(
                'SystemQuestion', winsound.SND_ALIAS | winsound.SND_ASYNC)
            root.after(10000, self.playSound)

    def closeFound(self):
        self.soundFlag = 0  # stop playing sound
        self.grandparent.toggleStop()
        self.top.grab_release()
        self.top.destroy()

    def continueSearch(self):
        self.soundFlag = 0
        self.top.grab_release()
        self.top.destroy()


class recoveryDialog:
    def __init__(self, parent):
        top = self.top = Toplevel(root)
        batchNum = root.crntSetting.readBatchNum()
        Label(top, text="Batch %s Recovery:" %
              batchNum, font=("Ariel", 12)).grid(row=0, column=0)
        Plot = plot(parent)
        L = str(Plot.recovery*1000)[:5]
        Label(top, text=L, font=("Ariel", 16)).grid(row=1, column=0)
        Label(top, text=u'R\u00B2', font=("Ariel", 12)).grid(row=0, column=1)
        R = str(Plot.R2)[:8]
        Label(top, text=R, font=("Ariel", 16)).grid(row=1, column=1)

        canvas = Plot.graph(self, top)  # Graph
        canvas.get_tk_widget().grid(row=2, column=0, columnspan=2)
        top.grab_set()

    def closeRecovery(self):
        self.top.grab_release()
        self.top.destroy()


class serialError:
    def __init__(self, parent, grandparent):
        self.grandparent = grandparent
        self.parent = parent
        top = self.top = Toplevel(grandparent)

        Label(top, text="Cannot Connect to EGM").pack(pady=5)

        self.selected = IntVar()  # tkinter variable BS
        # get current value
        self.selected.set(self.grandparent.crntSetting.com)
        Radiobutton(top, text="COM1", variable=self.selected,
                    value=0).pack(anchor='c')
        Radiobutton(top, text="COM2", variable=self.selected,
                    value=1).pack(anchor='c')
        Radiobutton(top, text="COM3", variable=self.selected,
                    value=2).pack(anchor='c')
        Radiobutton(top, text="COM4", variable=self.selected,
                    value=3).pack(anchor='c')
        Radiobutton(top, text="COM5", variable=self.selected,
                    value=4).pack(anchor='c')
        Radiobutton(top, text="COM6", variable=self.selected,
                    value=5).pack(anchor='c')
        Radiobutton(top, text="COM7", variable=self.selected,
                    value=6).pack(anchor='c')
        Radiobutton(top, text="COM8", variable=self.selected,
                    value=7).pack(anchor='c')
        Button(top, text="TRY AGAIN", command=self.closeError).pack(anchor='c')
        top.geometry("100x270+300+220")

        top.grab_set()  # deactivates main window

    def closeError(self):  # closes error dialog
        self.grandparent.crntSetting.com = self.selected.get()
        self.top.grab_release()
        self.top.destroy()
        self.parent.__init__(self.grandparent)

#---------------------------------------------------------SETUPS and DEFAULTS


class serial:
    def __init__(self, parent):
        self.parent = parent
        try:
            # setup serial port and open
            self.ser = Serial(parent.crntSetting.com, stopbits=2, timeout=None)
            self.ser.close()
        except SerialException:  # Creates error dialog instance
            crntError = serialError(self, self.parent)

    def readLine(self):

        # stream buffer, 84 because that makes sure one reading (part of line) is captured
        buffered = self.ser.read(84)

        for i in range(len(buffered)):  # Looks for M in stream buffer
            if buffered[i] == 'M':
                self.crntRead = buffered[i:i+21]
                print("CrntRead: "+self.crntRead)
                break

        return self.crntRead

    def closePort(self):
        self.ser.close()

    def openPort(self):
        self.ser.open()


class settings:  # object class to hold com settings
    def __init__(self):  # default to com 1
        self.com = 0
        self.setting = shelve.open('EGMsettings')

    def setDickson(self, DIC, BatchNum):
        self.setting['DIC'] = DIC
        self.setting['BatchNum'] = BatchNum

    def readDIC(self):
        try:
            return self.setting['DIC']
        except KeyError as exception:
            return 'Not Set'

    def readBatchNum(self):
        try:
            return self.setting['BatchNum']
        except KeyError as exception:
            return 0


class DBmanager():  # Handles all actions with the database
    def __init__(self):
        self.value = 1

    def new(self):  # creates new database
        filename = tkFileDialog.asksaveasfilename(title="New File")
        self.data = sqlite3.connect(filename, check_same_thread=False)
        self.c = self.data.cursor()
        self.c.execute(
            '''CREATE TABLE egm(ID,sample1,sample2,sample3,sample4,sample5);''')
        #self.c.execute('''INSERT INTO egm VALUES (1,0,0,0,0,0)''')
        self.data.commit()  # Save (commit) the changes

    def openFile(self):
        filename = tkFileDialog.askopenfilename(title="Open Database")
        self.data = sqlite3.connect(filename, check_same_thread=False)
        self.c = self.data.cursor()

    def get(self):
        self.c.execute('select * from EGM')
        self.dataUpdate = self.c.fetchall()  # Gets latest data from db file
        return(self.dataUpdate)

    def update(self, sampleNum, value, row):  # updates a row with new data
        self.c.execute("UPDATE egm SET %s = %s WHERE ID = %s" %
                       (sampleNum, value, row))
        self.data.commit()

    def newRow(self):  # creates a new blank crow
        self.c.execute('select * from EGM')
        self.rowID = self.c.fetchall()
        if self.rowID == []:  # if first row
            self.c.execute("INSERT INTO egm VALUES (1,0,0,0,0,0)")
        else:
            self.c.execute("INSERT INTO egm VALUES (%s,0,0,0,0,0)" %
                           (self.rowID[-1][0]+1))
        self.data.commit()  # Save (commit) the changes

    def insertP1(self, p1):
        self.c.execute('select * from EGM')
        self.rowID = self.c.fetchall()[-1][0]  # last row's ID
        print("rowID", self.rowID)

        self.c.execute("INSERT INTO egm VALUES (%s,%s,%s,%s,%s,%s)" %
                       (self.rowID+1, p1, 0, 0, 0, 0))
        self.data.commit()

    def lastRow(self):
        self.c.execute('SELECT Count(*) FROM egm')
        self.rows = self.c.fetchall()  # Gets latest number of rows from db file
        return(self.rows)

    def clearLastRow(self):
        self.c.execute('SELECT Count(*) FROM egm')
        self.rows = self.c.fetchall()  # Gets latest number of rows from db file
        self.c.execute(
            "UPDATE egm SET sample1 = 0, sample2 = 0, sample3 = 0, sample4 = 0, sample5 = 0 WHERE ID = %s" % (self.rows[0][0]))

    def deleteLastRow(self):
        self.c.execute('SELECT Count(*) FROM egm')
        # Gets latest number of rows from db file
        self.rows = self.c.fetchall()[0][0]
        print(self.rows)
        self.c.execute("DELETE FROM egm WHERE ID = %s" % (self.rows))


# ------------------------------------------------------ACTION CLASSES & DEINITIONS

class plot:
    def __init__(self, parent):
        self.parent = parent
        data = parent.crntDBmanager.get()
        pt1count = 0
        pt2count = 0
        pt3count = 0
        pt4count = 0
        pt5count = 0
        pt6count = 0
        ptDcount = 0

        pt1 = 0
        pt2 = 0
        pt3 = 0
        pt4 = 0
        pt5 = 0
        pt6 = 0
        ptD = 0

        for i in range(6)[1:]:  # 0 mM get y value if not 0
            if data[0][i] != 0:
                pt1 += data[0][i]
                pt1count += 1
        pt1 = float(pt1)/pt1count

        for i in range(6)[1:]:  # .25 mM
            if data[1][i] != 0:
                pt2 += + data[1][i]
                pt2count += 1
        pt2 = float(pt2)/pt2count

        for i in range(6)[1:]:  # .5 mM
            if data[2][i] != 0:
                pt3 += data[2][i]
                pt3count += 1
        pt3 = float(pt3)/pt3count

        for i in range(6)[1:]:  # 1 mM
            if data[3][i] != 0:
                pt4 += data[3][i]
                pt4count += 1
        pt4 = float(pt4)/pt4count

        for i in range(6)[1:]:  # 2 mM
            if data[4][i] != 0:
                pt5 += data[4][i]
                pt5count += 1
        pt5 = float(pt5)/pt5count

        for i in range(6)[1:]:  # 2 mM
            if data[5][i] != 0:
                pt6 += data[5][i]
                pt6count += 1
        pt6 = float(pt6)/pt6count

        for i in range(6)[1:]:  # 2 mM
            if data[6][i] != 0:
                ptD += data[6][i]
                ptDcount += 1
        ptD = float(ptD)/ptDcount

        self.x = numpy.array([0, .25, .5, 1, 2, 4])  # standard concentration
        self.y = numpy.array([pt1, pt2, pt3, pt4, pt5, pt6])
        # calculate slope and intercept
        m, b = numpy.polyfit(self.x, self.y, 1)

        # Calculating R squared
        p = numpy.poly1d([m, b])
        yhat = p(self.x)
        ybar = numpy.sum(self.y)/len(self.y)
        yhat = p(self.x)                         # or [p(z) for z in x]
        ybar = numpy.sum(self.y)/len(self.y)          # or sum(y)/len(y)
        # or sum([ (yihat - ybar)**2 for yihat in yhat])
        ssreg = numpy.sum((yhat-ybar)**2)
        # or sum([ (yi - ybar)**2 for yi in y])
        sstot = numpy.sum((self.y - ybar)**2)
        self.R2 = ssreg / sstot

        # Calc Recovery
        crntDIC = root.crntSetting.readDIC()
        dicksonPredictX = (ptD-b)/m
        # stuff in quotes is to remove extra decimals
        self.recovery = (dicksonPredictX/crntDIC)*100
        print(self.recovery)

    def graph(self, tParent, top):
        fig = plt.figure()
        ax1 = fig.add_subplot(1, 1, 1)
        ax1.plot(self.x, self.y, 'ob-')
        ax1.set_xlabel('Standard Concentration (nM Sodium Carbonate)')
        ax1.set_ylabel('Mean EGM Reading (ppm)')
        canvas = FigureCanvasTkAgg(fig, master=top)
        return canvas


class findPeaks:  # to find peaks
    def __init__(self, obj, parent):  # initilizes variables, loads in variables
        self.parent = parent  # parent is app
        print("initilization")
        self.passedDB = obj
        self.count = 1
        self.read1 = 9999
        self.read2 = 9999
        self.read3 = 9999
        self.flag = 0

    def resetReads(self):
        self.read1 = 9999
        self.read2 = 9999
        self.read3 = 9999

    def run(self):  # Peak finding loop
        currentRead = root.crntSerial.readLine()  # 'M000000000000000003400000'

        # If in zero mode, clear the reads, the reads after zero are different, can lead to false peak
        if currentRead[0] == 'Z':
            self.resetReads()
        # check if is a Record mode (not warmup, not zero)
        if currentRead[0] == 'M':

            # check if current read is different from past
            if currentRead[16:20] != self.read3:
                self.read1 = self.read2  # Discard stored read 1, move rest down
                self.read2 = self.read3
                self.read3 = currentRead[16:20]

                rows = self.passedDB.lastRow()[0][0]
                # Gets latest number of rows from db file
                if self.read2 > self.read1 and self.read2 > self.read3:  # Check to see if peak
                    if self.count == 5:
                        print("pass logic5")
                        p5 = self.read2
                        self.passedDB.update('sample5', p5, rows)
                        self.count += 1
                        self.parent.update()
                        self.parent.toggleStop()  # stops search
                        self.found4.closeFound()
                        found5peaks()
                    elif self.count == 4:
                        print("pass logic4")
                        p4 = self.read2
                        self.passedDB.update('sample4', p4, rows)
                        self.count += 1
                        self.parent.update()
                        self.found4 = found4peaks(self.parent)
                    elif self.count == 3:
                        print("pass logic3")
                        print(self.read1)
                        print(self.read2)
                        print(self.read3)
                        p3 = self.read2
                        self.passedDB.update('sample3', p3, rows)
                        self.count += 1
                        self.parent.update()
                    elif self.count == 2:
                        print("pass logic2")
                        print(self.read1)
                        print(self.read2)
                        print(self.read3)
                        p2 = self.read2
                        self.passedDB.update('sample2', p2, rows)
                        self.count += 1
                        self.parent.update()
                    elif self.count == 1:  # if first peak, record into db
                        if self.read1 != 9999:  # for some reason it thinks 9999 is less than say 400
                            print("pass logic1")
                            print(self.read1)
                            print(self.read2)
                            print(self.read3)
                            p1 = self.read2
                            self.passedDB.update('sample1', p1, rows)
                            self.count += 1
                            self.parent.update()


def export(crntDBmanager):  # File -> Export

    filename = tkFileDialog.asksaveasfilename(
        title="Export", defaultextension='.xls')  # get location to save file
    book = xlwt.Workbook(encoding="utf-8")  # Setup excel doc
    sheet1 = book.add_sheet("Raw_data")  # add sheet

    sheet1.write(0, 0, "Sample")
    sheet1.write(0, 1, "Peak1")
    sheet1.write(0, 2, "Peak2")
    sheet1.write(0, 3, "Peak3")
    sheet1.write(0, 4, "Peak4")
    sheet1.write(0, 5, "Peak5")
    sheet1.write(1, 0, "0mM")
    sheet1.write(2, 0, ".25mM")
    sheet1.write(3, 0, ".5mM")
    sheet1.write(4, 0, "1mM")
    sheet1.write(5, 0, "2mM")
    sheet1.write(6, 0, "4mM")

    dataUpdate = crntDBmanager.get()  # gets latest data

    rows = crntDBmanager.lastRow()  # gets number of rows
    height = 40
    width = 6
    for i in range(height):  # Rows
        for j in range(width)[:-1]:  # Columns
            if i < rows[0][0]:
                sheet1.write(i+1, j+1, dataUpdate[i][j+1])
    book.save(filename)  # save file


# ------------------------------------------Removing Tk Logo
ICON = zlib.decompress(base64.b64decode('eJxjYGAEQgEBBiDJwZDBy'
                                        'sAgxsDAoAHEQCEGBQaIOAg4sDIgACMUj4JRMApGwQgF/ykEAFXxQRc='))
_, ICON_PATH = tempfile.mkstemp()
with open(ICON_PATH, 'wb') as icon_file:
    icon_file.write(ICON)
root.iconbitmap(default=ICON_PATH)

# --------------------------------------------------------MAINLOOP


def main():
    root.geometry("441x581+100+100")  # Root geometry
    root.configure(background='SlateGray2')
    root.resizable(width='false', height='false')

    app = EGM(root)
    root.style = Style()
    root.style.theme_use('winnative')

    root.crntSetting = settings()  # initilizes default settings
    d = startupDialog(root)  # creates startup dialog instance
    root.wait_window(d.top)  # startup dialog
    root.crntSerial = serial(root)  # connects to serial port

    root.mainloop()


if __name__ == '__main__':
    main()

root.crntSetting.setting.close()
