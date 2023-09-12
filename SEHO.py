import sys
import pandas as pd
import logging
from PyQt5.QtWidgets import QApplication, QLabel, QLineEdit, QPushButton, QVBoxLayout, QWidget, QMessageBox, QListWidget
from PyQt5.QtGui import QFont, QIntValidator, QRegExpValidator
from PyQt5 import QtCore

import serial
import time
from _thread import *
import threading
import json


formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')

def setup_logger(name, log_file, level=logging.INFO):
    """To setup as many loggers as you want"""

    handler = logging.FileHandler(log_file)
    handler.setFormatter(formatter)

    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.addHandler(handler)

    return logger

# first file logger
info_logger = setup_logger('info_logger', 'logs\info_logfile.log')

# second file logger
error_logger = setup_logger('error_logger', 'logs\error_logfile.log')

# logging.basicConfig(filename='usb_log.txt', level=logging.INFO, format='%(asctime)s - %(message)s')

class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.write_lock = threading.Lock()
        self.read_lock = threading.Lock()
        f = open('data.json', 'r')
        self.defaults = json.load(f)
        self.initUI()
        self.load_excel_file()

    def focus_cod(self):
        self.input_field.setFocus()

    def initUI(self):
        self.setWindowTitle('Excel Search App')

        # SEHO 1 label
        seho_label = QLabel(self.defaults['name'], self)
        seho_label.setAlignment(QtCore.Qt.AlignCenter)
        seho_label.setFont(QFont("Arial", 24, QFont.Bold))

        onlyInt = QRegExpValidator()
        reg = QtCore.QRegExp('^[0-9]*$')
        onlyInt.setRegExp(reg)

        label3 = QLabel('Post de lucru:', self)
        self.post_field = QLineEdit(self)
        self.post_field.setValidator(onlyInt)
        self.post_field.setMaxLength(4)
        self.post_field.returnPressed.connect(self.focus_cod)

        label1 = QLabel('Cod produs:', self)
        self.input_field = QLineEdit(self)
        self.input_field.setValidator(onlyInt)
        self.input_field.setMaxLength(10)
        self.input_field.returnPressed.connect(self.search)

        label2 = QLabel('Rezultate:', self)
        self.results_field = QListWidget(self)

        read_button = QPushButton('Verifica', self)
        read_button.clicked.connect(self.readThread)


        search_button = QPushButton('Cauta', self)
        search_button.clicked.connect(self.search)

        send_button = QPushButton('Programeaza', self)
        send_button.clicked.connect(self.thread)

        self.status_label = QLabel('Status', self)
        self.status_label.setAlignment(QtCore.Qt.AlignCenter)
        self.status_label.setFont(QFont("Arial", 24))

        vbox = QVBoxLayout()
        vbox.addWidget(seho_label)  # Add the SEHO 1 label
        vbox.addStretch(1)
        vbox.addWidget(label3)
        vbox.addWidget(self.post_field)
        vbox.addWidget(label1)
        vbox.addWidget(self.input_field)
        vbox.addWidget(label2)
        vbox.addWidget(self.results_field)
        vbox.addWidget(search_button)
        vbox.addWidget(read_button)
        vbox.addWidget(send_button)
        vbox.addStretch(1)
        vbox.addWidget(self.status_label)
        vbox.addStretch(2)

        # Steinel Electronic label
        steinel_label = QLabel('Steinel Electronic', self)
        steinel_label.setAlignment(QtCore.Qt.AlignCenter)
        steinel_label.setFont(QFont("Arial", 10))

        vbox.addWidget(steinel_label)  # Add the Steinel Electronic label

        self.setLayout(vbox)
        self.setGeometry(100, 100, 1400, 800)
        self.setStyleSheet(
            "font-size: 26px;"
            "QPushButton { font-size: 26px; height: 80px; }"
        )
        self.showMaximized()
        self.show()

    def load_excel_file(self):
        try:
            # \\10.60.10.20\Company\IT\Tranzit\SEHO
            # self.data = pd.read_excel('C:/Users/avisca/OneDrive - steinel.de/Desktop/TEST python/Lista programe lipire Seho1 test.xlsx')
            # self.data = pd.read_excel(r'\\Company\Company\IT\Tranzit\SEHO\Lista programe lipire Seho1 test.xlsx')
            # self.data = pd.read_excel('Lista programe lipire Seho1 test.xlsx')
            self.data = pd.read_excel(self.defaults['path'], 0)
            self.code = pd.read_excel(self.defaults['path'], 3)
            self.data.columns = ['Program lipire', 'Denumire', 'NR. Cuib'] + list(self.data.columns[3:])
            self.code.columns = ['a', 'RFID', 'PID', 'b']
            # print(self.code)
        except Exception as e:
            error_message = f"Error loading Excel file: {str(e)}"
            QMessageBox.critical(self, "Error", error_message)
            # logging.error(error_message, exc_info=True)
            error_logger.error(error_message)

    def search(self):
        input_value = self.input_field.text()
        input_value = (int(input_value))

        # if len(input_value) > 10:
        #     QMessageBox.warning(self, "Atentie!", "Lungimea maxima a codului este de 10 caractere.")
        #     return

        # if input_value == "":
        #     QMessageBox.warning(self, "Atentie!", "Va rugam introduceti o valoare.")
        #     return

        try:
            pid = ''
            try:
                pid = str(int(self.code[self.code['RFID']==input_value]['PID'].values[0]))
            except:
                pid = ''
            result = self.data[self.data.astype(str).apply(lambda row: pid in row.values, axis=1)]

            if result.empty:
                self.results_field.clear()
                self.results_field.addItem('No results found.')
            else:
                self.results_field.clear()
                for _, row in result.iterrows():
                    result_string = f"Program lipire: {row['Program lipire']}, Denumire: {row['Denumire']}, NR. Cuib: {row[3]}"
                    self.results_field.addItem(result_string)
        except Exception as e:
            error_message = f"Error occurred during search: {str(e)}"
            QMessageBox.critical(self, "Error", error_message)
            # logging.error(error_message, exc_info=True)
            error_logger.error(error_message)

        # Verify if device is found

    def is_device_found(self, data):
        output = 'Program\t\t\t:i'
        data = data.split('\n')
        return output == data[len(data)-1]

    # Verify if device is written correctly

    def verify_write(self, data, program, station):
        data = data.split('\n')

        # Program
        aux = 0
        result = 1

        info_logger.info('Verifing values:\t Program: '+str(program)+' Station: '+str(station))
        try:
            aux = int(data[len(data)-8].split(':')[1])
            if(aux != program):
                # print('program not written')
                # error_logger.error('Did not write correctly program')
                raise Exception()
                result = 0
        except:
            # error_logger.error('Could not verify program')
            raise Exception('Device was not written correctly.')

            result = 0

        # Station

        try:
            aux = int(data[len(data)-7].split(':')[1])
            if(aux != station):
                # print('Did not write correctly station')
                # error_logger.error('Did not write correctly station')

                raise Exception()
                result = 0
        except:
            # error_logger.error('Could not verify station')
            raise Exception('Device was not written correctly.')
            result = 0

        return result

    def read_buffer(self, serial):
        bufferSize = serial.in_waiting
        return serial.read(bufferSize).decode()

    def write_device(self, program, statie):

        # Setup

        ser = serial.Serial()
        ser.port=self.defaults['port']
        ser.baudrate=9600
        ser.parity=serial.PARITY_NONE
        ser.stopbits=serial.STOPBITS_ONE
        ser.bytesize=serial.EIGHTBITS
        ser.timeout=100
        ser.open()
        self.status_label.setText('Se cauta...')
        time.sleep(2)
        ser.write("\033i".encode())
        time.sleep(1)

        # Searching for device

        print('Searching..')
        time.sleep(1)
        data = self.read_buffer(ser)
        index = 0
        while not self.is_device_found(data) and index < 5:
            time.sleep(1)
            print('Searching..')
            index = index+1

            data = self.read_buffer(ser)
        print('Found')
        self.status_label.setText('Se scrie...')

        # Programare RFID

        aux = str(program)
        aux = '\b'*len(aux)

        time.sleep(2)
        ser.write(aux.encode())
        time.sleep(0.2)
        ser.write(str(program).encode())
        time.sleep(0.2)
        ser.write('\r'.encode())

        time.sleep(0.2)
        ser.write(str(statie).encode())
        time.sleep(0.2)
        ser.write('\r'.encode())

        time.sleep(0.2)
        ser.write('l\r'.encode())
        time.sleep(0.2)
        ser.write('l\r'.encode())

        self.status_label.setText('Se verifica...')
        time.sleep(3.6)
        data = self.read_buffer(ser)
        time.sleep(0.2)
        self.verify_write(data, program, statie)

        print('Program:'+str(program)+'\tStation: '+str(statie))
        self.status_label.setText('Gata\n'+'Program:'+str(program)+'\tStatie: '+str(statie))
        ser.close()
        return 1

    def thread(self):
        self.write_lock.acquire()
        start_new_thread(self.send_to_usb, ())

    def readThread(self):
        self.read_lock.acquire()
        start_new_thread(self.readDevice, ())

    def readDevice(self):
        info_logger.info('Citire USB')
        try:
            ser = serial.Serial()
            ser.port=self.defaults['port']
            ser.baudrate=9600
            ser.parity=serial.PARITY_NONE
            ser.stopbits=serial.STOPBITS_ONE
            ser.bytesize=serial.EIGHTBITS
            ser.timeout=100
            ser.open()
            self.status_label.setText('Se cauta...')
            time.sleep(2)
            ser.write("\033i".encode())
            time.sleep(1)

            self.status_label.setText('Se verifica...')
            time.sleep(3.6)
            data = self.read_buffer(ser)
            time.sleep(0.2)
            program = int(data[len(data)-8].split(':')[1])
            # self.verify_write(data, program, statie)
            statie = int(data[len(data)-7].split(':')[1])

            print('Program:'+str(program)+'\tStation: '+str(statie))
            self.status_label.setText('Program:'+str(program)+'\tStatie: '+str(statie))
            ser.close()
            self.read_lock.release()

        except Exception as e:
            error_message = f"Error occurred reading USB: {str(e)}"
            error_logger.error(error_message)
            self.status_label.setText("Nu s-a putut citi")
            self.read_lock.release()


    def send_to_usb(self):
        self.input_field.setFocus()
        selected_item = self.results_field.currentItem()
        if selected_item is None:
            self.status_label.setText('Selectati un program')
            self.write_lock.release()
            return
        try:
            program_lipire = ''
            try:
                program_lipire = selected_item.text().split(':')[1].strip().split(',')[0]
            except:
                raise Exception('Invalid program selected.')
            post_de_lucru = self.post_field.text()
            # self.post_field.setText('')
            self.input_field.setText('')
            info_logger.info('Writing for product: '+self.input_field.text()+' Station: '+post_de_lucru+'\t'+self.results_field.currentItem().text())
            # Code to send data to serial port1
            self.write_device(int(program_lipire), int(post_de_lucru))

            # print(f"Sending data to USB port: Program Lipire: {program_lipire}, Post de lucru: {post_de_lucru}")
        except Exception as e:
            error_message = f"Error occurred while sending data to USB: {str(e)}"
            self.status_label.setText('Eroare')

            error_logger.error(error_message)
        self.write_lock.release()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyApp()
    sys.exit(app.exec_())
