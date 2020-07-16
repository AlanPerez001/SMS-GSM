from Firstqt_ui import *
import sys
import glob
import serial
from gsmmodem.modem import GsmModem
import serial.tools.list_ports
import time
import logging
from PyQt5.QtWidgets import QMessageBox, QFileDialog
import concurrent.futures
import threading
import sqlite3
import serial.tools.list_ports



def actualiza_mensajes(number,time,text):
    print(u'== SMS message received ==\nFrom: {0}\nTime: {1}\nMessage:\n{2}\n'.format(number,time,text))

files = ''
class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    
    def __init__(self, *args, **kwargs):
        QtWidgets.QMainWindow.__init__(self, *args, **kwargs)
        self.setupUi(self)
        self.stackedWidget.setCurrentIndex(1)
        self.progressBar.hide()
        self.progressBar_2.hide()
        self.treeWidget.itemClicked.connect(self.MenuSend)
        self.actionConectarModem.triggered.connect(self.actionConnect)
        self.actionEnviar_mensaje.triggered.connect(self.actionSend)
        self.actionProgramar_mensaje.triggered.connect(self.actionSchedul)
        self.actionExit.triggered.connect(self.actionExitapp)
        self.pushButton_5.clicked.connect(self.ConectModem)
        self.Actualizar.clicked.connect(self.ActualizarPort)
        self.ExitBTN.clicked.connect(self.actionExitapp)
        self.IniciarBTN.clicked.connect(self.startcon)
        self.Enviar.clicked.connect(self.start_send)
        self.comboBox.currentIndexChanged.connect(self.combo)
        self.PuertosBTN.clicked.connect(self.actionConnect)
        self.Connect_DB.clicked.connect(self.getfile)
        self.Actualiza_db.clicked.connect(self.updateDB) 
        self.crear_campaa.clicked.connect(self.create_show)
    
    def create_show(self):
        self.stackedWidget.setCurrentIndex(6)
        self.Create_DB.clicked.connect(self.create_db)

    def create_db(self):
        name_db=self.Name_create_db.toPlainText()
        print(name_db)
        files, _filter = QFileDialog.getExistingDirectoryUrl(None,'C:\\Users\\Developer\\Desktop\\')
        print(files )

    def getfile(self):
        global files
        files, _filter  = QFileDialog.getOpenFileName(self, 'Selecciona tu base de datos', 
        'C:\\Users\\Developer\\Desktop\\',"data base file (*.db)")
        print(files)
        self.updateDB()

    def actionSchedul(self):
        self.stackedWidget.setCurrentIndex(2)
        self.pushButton_5.clicked.connect(self.ConectModem)

    def actionExitapp(self):
        try:
            app.exit()
            self.modem.close()
        except Exception as e:
            print(e)

    def actionConnect(self):
        self.stackedWidget.setCurrentIndex(1)
    
    def actionSend(self):
        self.stackedWidget.setCurrentIndex(0)
        self.pushButton.clicked.connect(self.SendSms)

    def combo(self):
        seleccion = self.comboBox.currentText()
        if str(seleccion)=='Componer':
            self.stackedWidget.setCurrentIndex(0)
            self.pushButton.clicked.connect(self.SendSms)
        elif str(seleccion) == 'Programado':
            self.stackedWidget.setCurrentIndex(2)

    def ConectModem(self):
        try:
            portSelected = self.listWidget.currentItem().text() 
            print(portSelected)
            try:
                self.modem = GsmModem(str(portSelected), 115200,smsReceivedCallbackFunc= MainWindow.ReceivedSms)
                self.modem.smsTextMode = False
                self.modem.connect()
                print('Conectado')
            except Exception as F:
                print('El puerto '+str(portSelected)+' No esta disponible: '+ str(F))
        except Exception as e:
            print(e)

    def startcon (self):
        try:
            name = self.modem.networkName
            print(name)
            portSelected = self.listWidget.currentItem().text() 
            print(portSelected)
            try:
                self.modem = GsmModem(str(portSelected), 115200,smsReceivedCallbackFunc= MainWindow.ReceivedSms)
                self.modem.smsTextMode = False
                self.modem.connect()
                print('Conectado')
            except Exception as F:
                print('El puerto '+str(portSelected)+' No esta disponible: '+ str(F))
        except Exception as e:
            print(e)
            if str(e) == "'MainWindow' object has no attribute 'modem'":
                print('Modem No conectado, selecciona un puerto por favor')
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Error")
                msg.setInformativeText('Modem No conectado, selecciona un puerto por favor')
                msg.setWindowTitle("Error")
                msg.exec_()
                    
    def ActualizarPort(self):
        ports = serial.tools.list_ports.comports()
        advance = 0
        porcantaje = float(100/len(ports))
        self.listWidget.clear()
        hilos = len(ports)
        count_hilos = 1
        num_hilo = 1
        for p in sorted(ports,key=None, reverse=False):
            advance += float(porcantaje)
            port = p.device
            print(port)
            try:
                self.progressBar.show()
                self.progressBar.setGeometry(QtCore.QRect(300, 330, 201, 21))
                self.progressBar.setProperty("value", advance)
                self.progressBar.setObjectName("progressBar")
                hilo = threading.Thread(name='hilo %s' % str(num_hilo), target=self.connectport, args=(port,))
                hilo.start()
                num_hilo+=1
            finally:
                self.progressBar.hide()
                
    def connectport(self,port):
        try: 
            phone = serial.Serial(str(port),  115200, timeout=0.2)
            phone.write(b'ATZ\r')
            phone.write(b'AT+CMGF=1\r')
            if str(phone.read()) != str(b''):
                self.listWidget.addItem(port)
        except Exception as E:
            if str(E) == 'None':
                print('No disponibe :c')
        finally:
            try:
                phone.close()
            except Exception as f:
                print(f)

    def SendSms(self):
        num = self.Num_edit.toPlainText()
        text = self.SMSText_edit.toPlainText()
        try:
            self.modem.sendSms(num, text,waitForDeliveryReport=False, deliveryTimeout=1)
            print('Mensaje enviado')
            self.Num_edit.clear()
            self.SMSText_edit.clear()
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setInformativeText('Mensaje enviado con exito :D')
            msg.setWindowTitle("info")
            msg.exec_()
            
        except Exception as F:
            if str(F) == 'CMS 500':
                print('El puerto seleccionado devolvio '+str(F)+', [Error Unknown]')
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Error")
                msg.setInformativeText(str(F))
                msg.setWindowTitle("Error")
                msg.exec_()

    def ReceivedSms(self):
        received = threading.Thread(target = actualiza_mensajes(self.number,self.time,self.text))
        received.start()

    def updateDB(self):
        if files != '':
            con = sqlite3.connect(str(files))
            cursor = con.cursor()
            cursor.execute('SELECT* FROM SMS')
            rows = cursor.fetchall()
            count = 0
            for i in rows:
                numero = QtWidgets.QTableWidgetItem()
                self.tableWidget_4.setItem(count, 0, numero)
                numero = self.tableWidget_4.item(count, 0)
                numero.setText(str(i[1]))
                mensaje = QtWidgets.QTableWidgetItem()
                self.tableWidget_4.setItem(count, 1, mensaje)
                mensaje = self.tableWidget_4.item(count, 1)
                mensaje.setText(str(i[2]))
                enviado = QtWidgets.QTableWidgetItem()
                self.tableWidget_4.setItem(count, 2, enviado)
                enviado = self.tableWidget_4.item(count, 2)
                enviado.setText(str(i[3]))
                count +=1
            cursor.execute('SELECT* FROM Enviados')
            rows = cursor.fetchall()
            count2 = 0

            for i in rows:
                numero = QtWidgets.QTableWidgetItem()
                self.tableWidget_3.setItem(count2, 0, numero)
                numero = self.tableWidget_3.item(count2, 0)
                numero.setText(str(i[1]))
                mensaje = QtWidgets.QTableWidgetItem()
                self.tableWidget_3.setItem(count2, 1, mensaje)
                mensaje = self.tableWidget_3.item(count2, 1)
                mensaje.setText(str(i[2]))
                com = QtWidgets.QTableWidgetItem()
                self.tableWidget_3.setItem(count2, 2, com)
                com = self.tableWidget_3.item(count2, 2)
                com.setText(str(i[3]))
                date = QtWidgets.QTableWidgetItem()
                self.tableWidget_3.setItem(count2, 3, date)
                date = self.tableWidget_3.item(count2, 3)
                date.setText(str(i[4]))
                count2 +=1
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Base de datos de campaña no seleccionados.')
            msg.setWindowTitle("Error")
            msg.exec_()


    def Com5(self):
        port = 'COM5'
        try:
            con = sqlite3.connect('SMSTest.db')
            cursor = con.cursor()
            cursor.execute('Select * FROM SMS WHERE Enviado ="F"')
            rows = cursor.fetchall()
            parametro = rows[0]
            print('El hilo es: {0},   El parametro es: {1}'.format(threading.current_thread().getName(),parametro[0]))
            if str(parametro[3])=='F':
                cursor.execute('UPDATE SMS SET Enviado ="I" WHERE ID ='+str(parametro[0]))
                con.commit()
                modem = GsmModem(port, 115200)
                try:
                    modem.smsTextMode = False
                    modem.connect()
                    modem.sendSms(str(parametro[1]), str(parametro[2]),waitForDeliveryReport=False, deliveryTimeout=1)
                    modem.close()
                    #time.sleep(2)
                    cursor.execute('UPDATE SMS SET Enviado ="T" WHERE ID ='+str(parametro[0]))
                    con.commit()
                    cursor.execute('INSERT INTO Enviados (Numero,Mensaje,Puerto,Hora_envio) VALUES("'+str(parametro[1])+'","'+str(parametro[2])+'","'+str(port)+'","'+str(time.asctime( time.localtime(time.time()) ))+'")')
                    con.commit()
                    self.updateDB()
                except Exception as e:
                    print(e)
                    cursor.execute('UPDATE SMS SET Enviado ="F" WHERE ID ='+str(parametro[0]))
                    con.commit()
        except Exception as e:
            e
        finally:
            try:
                modem.close()
            except Exception as e:
                e

    def Com6(self):
        port = 'COM6'
        try:
            con = sqlite3.connect('SMSTest.db')
            cursor = con.cursor()
            cursor.execute('Select * FROM SMS WHERE Enviado ="F"')
            rows = cursor.fetchall()
            parametro = rows[0]
            print('El hilo es: {0},   El parametro es: {1}'.format(threading.current_thread().getName(),parametro[0]))
            if str(parametro[3])=='F':
                cursor.execute('UPDATE SMS SET Enviado ="I" WHERE ID ='+str(parametro[0]))
                con.commit()
                modem = GsmModem(port, 115200)
                try:
                    modem.smsTextMode = False
                    modem.connect()
                    modem.sendSms(str(parametro[1]), str(parametro[2]),waitForDeliveryReport=False, deliveryTimeout=1)
                    modem.close()
                    #time.sleep(2)
                    cursor.execute('UPDATE SMS SET Enviado ="T" WHERE ID ='+str(parametro[0]))
                    con.commit()
                    cursor.execute('INSERT INTO Enviados (Numero,Mensaje,Puerto,Hora_envio) VALUES("'+str(parametro[1])+'","'+str(parametro[2])+'","'+str(port)+'","'+str(time.asctime( time.localtime(time.time()) ))+'")')
                    con.commit()
                    self.updateDB()
                except Exception as e:
                    print(e)
                    cursor.execute('UPDATE SMS SET Enviado ="F" WHERE ID ='+str(parametro[0]))
                    con.commit()
        except Exception as e:
            e
        finally:
            try:
                modem.close()
            except Exception as e:
                e 

    def Com7(self):
        port = 'COM7'
        try:
            con = sqlite3.connect('SMSTest.db')
            cursor = con.cursor()
            cursor.execute('Select * FROM SMS WHERE Enviado ="F"')
            rows = cursor.fetchall()
            parametro = rows[0]
            print('El hilo es: {0},   El parametro es: {1}'.format(threading.current_thread().getName(),parametro[0]))
            if str(parametro[3])=='F':
                cursor.execute('UPDATE SMS SET Enviado ="I" WHERE ID ='+str(parametro[0]))
                con.commit()
                modem = GsmModem(port, 115200)
                try:
                    modem.smsTextMode = False
                    modem.connect()
                    modem.sendSms(str(parametro[1]), str(parametro[2]),waitForDeliveryReport=False, deliveryTimeout=1)
                    modem.close()
                    #time.sleep(2)
                    cursor.execute('UPDATE SMS SET Enviado ="T" WHERE ID ='+str(parametro[0]))
                    con.commit()
                    cursor.execute('INSERT INTO Enviados (Numero,Mensaje,Puerto,Hora_envio) VALUES("'+str(parametro[1])+'","'+str(parametro[2])+'","'+str(port)+'","'+str(time.asctime( time.localtime(time.time()) ))+'")')
                    con.commit()
                    self.updateDB()
                except Exception as e:
                    print(e)
                    cursor.execute('UPDATE SMS SET Enviado ="F" WHERE ID ='+str(parametro[0]))
                    con.commit()
        except Exception as e:
            e
        finally:
            try:
                modem.close()
            except Exception as e:
                e 
            
    def Com8(self):
        port = 'COM8'
        try:
            con = sqlite3.connect('SMSTest.db')
            cursor = con.cursor()
            cursor.execute('Select * FROM SMS WHERE Enviado ="F"')
            rows = cursor.fetchall()
            parametro = rows[0]
            print('El hilo es: {0},   El parametro es: {1}'.format(threading.current_thread().getName(),parametro[0]))
            if str(parametro[3])=='F':
                cursor.execute('UPDATE SMS SET Enviado ="I" WHERE ID ='+str(parametro[0]))
                con.commit()
                modem = GsmModem(port, 115200)
                try:
                    modem.smsTextMode = False
                    modem.connect()
                    modem.sendSms(str(parametro[1]), str(parametro[2]),waitForDeliveryReport=False, deliveryTimeout=1)
                    modem.close()
                    #time.sleep(2)
                    cursor.execute('UPDATE SMS SET Enviado ="T" WHERE ID ='+str(parametro[0]))
                    con.commit()
                    cursor.execute('INSERT INTO Enviados (Numero,Mensaje,Puerto,Hora_envio) VALUES("'+str(parametro[1])+'","'+str(parametro[2])+'","'+str(port)+'","'+str(time.asctime( time.localtime(time.time()) ))+'")')
                    con.commit()
                    self.updateDB()
                except Exception as e:
                    print(e)
                    cursor.execute('UPDATE SMS SET Enviado ="F" WHERE ID ='+str(parametro[0]))
                    con.commit()
        except Exception as e:
            print(e)
            if str(e) == 'list index out of range':
                self.updateDB
        finally:
            try:
                modem.close()
            except Exception as e:
                e 
            
    def Com9(self):
        port = 'COM9'
        try:
            con = sqlite3.connect('SMSTest.db')
            cursor = con.cursor()
            cursor.execute('Select * FROM SMS WHERE Enviado ="F"')
            rows = cursor.fetchall()
            parametro = rows[0]
            print('El hilo es: {0},   El parametro es: {1}'.format(threading.current_thread().getName(),parametro[0]))
            if str(parametro[3])=='F':
                cursor.execute('UPDATE SMS SET Enviado ="I" WHERE ID ='+str(parametro[0]))
                con.commit()
                modem = GsmModem(port, 115200)
                try:
                    modem.smsTextMode = False
                    modem.connect()
                    modem.sendSms(str(parametro[1]), str(parametro[2]),waitForDeliveryReport=False, deliveryTimeout=1)
                    modem.close()
                    #time.sleep(2)
                    cursor.execute('UPDATE SMS SET Enviado ="T" WHERE ID ='+str(parametro[0]))
                    con.commit()
                    cursor.execute('INSERT INTO Enviados (Numero,Mensaje,Puerto,Hora_envio) VALUES("'+str(parametro[1])+'","'+str(parametro[2])+'","'+str(port)+'","'+str(time.asctime( time.localtime(time.time()) ))+'")')
                    con.commit()
                    self.updateDB()
                except Exception as e:
                    print(e)
                    cursor.execute('UPDATE SMS SET Enviado ="F" WHERE ID ='+str(parametro[0]))
                    con.commit()
        except Exception as e:
            e
        finally:
            try:
                modem.close()
            except Exception as e:
                e 
            
    def Com10(self):
        port = 'COM10'
        try:
            con = sqlite3.connect('SMSTest.db')
            cursor = con.cursor()
            cursor.execute('Select * FROM SMS WHERE Enviado ="F"')
            rows = cursor.fetchall()
            parametro = rows[0]
            print('El hilo es: {0},   El parametro es: {1}'.format(threading.current_thread().getName(),parametro[0]))
            if str(parametro[3])=='F':
                cursor.execute('UPDATE SMS SET Enviado ="I" WHERE ID ='+str(parametro[0]))
                con.commit()
                modem = GsmModem(port, 115200)
                try:
                    modem.smsTextMode = False
                    modem.connect()
                    modem.sendSms(str(parametro[1]), str(parametro[2]),waitForDeliveryReport=False, deliveryTimeout=1)
                    modem.close()
                    #time.sleep(2)
                    cursor.execute('UPDATE SMS SET Enviado ="T" WHERE ID ='+str(parametro[0]))
                    con.commit()
                    cursor.execute('INSERT INTO Enviados (Numero,Mensaje,Puerto,Hora_envio) VALUES("'+str(parametro[1])+'","'+str(parametro[2])+'","'+str(port)+'","'+str(time.asctime( time.localtime(time.time()) ))+'")')
                    con.commit()
                    self.updateDB()
                except Exception as e:
                    print(e)
                    cursor.execute('UPDATE SMS SET Enviado ="F" WHERE ID ='+str(parametro[0]))
                    con.commit()
        except Exception as e:
            e
        finally:
            try:
                modem.close()
            except Exception as e:
                e 
            
    def Com11(self):
        port = 'COM11'
        try:
            con = sqlite3.connect('SMSTest.db')
            cursor = con.cursor()
            cursor.execute('Select * FROM SMS WHERE Enviado ="F"')
            rows = cursor.fetchall()
            parametro = rows[0]
            print('El hilo es: {0},   El parametro es: {1}'.format(threading.current_thread().getName(),parametro[0]))
            if str(parametro[3])=='F':
                cursor.execute('UPDATE SMS SET Enviado ="I" WHERE ID ='+str(parametro[0]))
                con.commit()
                modem = GsmModem(port, 115200)
                try:
                    modem.smsTextMode = False
                    modem.connect()
                    modem.sendSms(str(parametro[1]), str(parametro[2]),waitForDeliveryReport=False, deliveryTimeout=1)
                    modem.close()
                    #time.sleep(2)
                    cursor.execute('UPDATE SMS SET Enviado ="T" WHERE ID ='+str(parametro[0]))
                    con.commit()
                    cursor.execute('INSERT INTO Enviados (Numero,Mensaje,Puerto,Hora_envio) VALUES("'+str(parametro[1])+'","'+str(parametro[2])+'","'+str(port)+'","'+str(time.asctime( time.localtime(time.time()) ))+'")')
                    con.commit()
                    self.updateDB()
                except Exception as e:
                    print(e)
                    cursor.execute('UPDATE SMS SET Enviado ="F" WHERE ID ='+str(parametro[0]))
                    con.commit()
        except Exception as e:
            e
        finally:
            try:
                modem.close()
            except Exception as e:
                e 
            
    def Com12(self):
        port = 'COM12'
        try:
            con = sqlite3.connect('SMSTest.db')
            cursor = con.cursor()
            cursor.execute('Select * FROM SMS WHERE Enviado ="F"')
            rows = cursor.fetchall()
            parametro = rows[0]
            print('El hilo es: {0},   El parametro es: {1}'.format(threading.current_thread().getName(),parametro[0]))
            if str(parametro[3])=='F':
                cursor.execute('UPDATE SMS SET Enviado ="I" WHERE ID ='+str(parametro[0]))
                con.commit()
                modem = GsmModem(port, 115200)
                try:
                    modem.smsTextMode = False
                    modem.connect()
                    modem.sendSms(str(parametro[1]), str(parametro[2]),waitForDeliveryReport=False, deliveryTimeout=1)
                    modem.close()
                    #time.sleep(2)
                    cursor.execute('UPDATE SMS SET Enviado ="T" WHERE ID ='+str(parametro[0]))
                    con.commit()
                    cursor.execute('INSERT INTO Enviados (Numero,Mensaje,Puerto,Hora_envio) VALUES("'+str(parametro[1])+'","'+str(parametro[2])+'","'+str(port)+'","'+str(time.asctime( time.localtime(time.time()) ))+'")')
                    con.commit()
                    self.updateDB()
                except Exception as e:
                    print(e)
                    cursor.execute('UPDATE SMS SET Enviado ="F" WHERE ID ='+str(parametro[0]))
                    con.commit()
        except Exception as e:
            e
        finally:
            try:
                modem.close()
            except Exception as e:
                e 
   
    def start_send(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setInformativeText('Iniciando envio de mensajes\nEste proceso puede demorar dependiendo de la cantidad de mensajes a enviar')
        msg.setWindowTitle("info")
        msg.exec_()
        hilostart = threading.Thread(target=self.hilos()) 
        hilostart.start()
        hilostart.join()
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Info")
        msg.setInformativeText('Mensajes enviados con exito')
        msg.setWindowTitle("info")
        msg.exec_()

    def SecondRound(self):
        con = sqlite3.connect('SMSTest.db')
        cursor = con.cursor()
        try:
            cursor.execute('SELECT * FROM SMS WHERE Enviado = "F"')
            rows = cursor.fetchall()
            if rows:
                print('Si entro')
                self.hilos()
        except Exception as e:
            e
        time.sleep(5)
        try:
            cursor.execute('SELECT * FROM SMS WHERE Enviado = "I"')
            rows = cursor.fetchall()
            if rows:
                print('Si entro con enviado I')
                try:
                    cursor.execute('SELECT * FROM SMS WHERE Enviado = "F"')
                    rows = cursor.fetchall()
                    if rows:
                        print('Si hay mensajes')
                        self.hilos()
                except Exception as e:
                    e

        except Exception as e:
            print('Hubo pedo carnal'+str(e))



    def hilos(self):
        com5= threading.Thread(name ='HCOM5', target=self.Com5)
        com6= threading.Thread(name ='HCOM6',target=self.Com6)
        com7= threading.Thread(name ='HCOM7',target=self.Com7)
        com8= threading.Thread(name ='HCOM8',target=self.Com8)
        com9= threading.Thread(name ='HCOM9',target=self.Com9 )
        com10= threading.Thread(name ='HCOM10',target=self.Com10 )
        com11= threading.Thread(name ='HCOM11',target=self.Com11 )
        com12= threading.Thread(name ='HCOM12',target=self.Com12 )
        com5.start()
        time.sleep(2)
        com6.start()
        time.sleep(2)
        com7.start()
        time.sleep(2)
        com8.start()
        time.sleep(2)
        com9.start()
        time.sleep(2)
        com10.start()
        time.sleep(2)
        com11.start()
        time.sleep(2)
        com12.start()
        time.sleep(3)
        if com5.isAlive() == False:
            back = threading.Thread(target=self.SecondRound)
            back.start()


    def MenuSend(self, it, col):
        if it.text(col) == 'Enviados':
            print('Entrando a Enviados')
            self.stackedWidget.setCurrentIndex(3)
        elif it.text(col) =='Recibidos':
            print('Entrando a Recibidos')
            self.stackedWidget.setCurrentIndex(4)
        elif it.text(col) =='Base de datos':
            print('Entrando a Base')
            self.stackedWidget.setCurrentIndex(5)
        elif it.text(col) == 'Programados':
            print('Entrando a Programados')
            self.stackedWidget.setCurrentIndex(2)

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()