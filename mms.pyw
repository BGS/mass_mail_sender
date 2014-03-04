#-*- coding: utf-8 -*-

import sip

sip.setapi("QString", 2)
sip.setapi('QVariant', 2)

from PyQt4 import QtCore, QtGui

from email.MIMEBase import MIMEBase
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText

from tempfile import mkstemp
from shutil import move
from os import remove, close

from smtplib import SMTP
import smtplib

import email , mimetypes

import itertools

import time

import copy

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    _fromUtf8 = lambda s: s

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName(_fromUtf8("MainWindow"))
        MainWindow.resize(438, 580)
        MainWindow.setTabShape(QtGui.QTabWidget.Triangular)
        self.centralWidget = QtGui.QWidget(MainWindow)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.centralWidget.sizePolicy().hasHeightForWidth())
        self.centralWidget.setSizePolicy(sizePolicy)
        self.centralWidget.setObjectName(_fromUtf8("centralWidget"))
        self.gridLayout_3 = QtGui.QGridLayout(self.centralWidget)
        self.gridLayout_3.setObjectName(_fromUtf8("gridLayout_3"))
        self.label_4 = QtGui.QLabel(self.centralWidget)
        self.label_4.setObjectName(_fromUtf8("label_4"))
        self.gridLayout_3.addWidget(self.label_4, 0, 0, 1, 1)
        self.textarea = QtGui.QTextEdit(self.centralWidget)
        self.textarea.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.textarea.setLineWrapMode(QtGui.QTextEdit.NoWrap)
        self.textarea.setObjectName(_fromUtf8("textarea"))
        self.gridLayout_3.addWidget(self.textarea, 1, 0, 1, 2)
        self.textEdit = QtGui.QTextEdit(self.centralWidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit.sizePolicy().hasHeightForWidth())
        self.textEdit.setSizePolicy(sizePolicy)
        self.textEdit.setHtml(_fromUtf8("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:8pt;\"></p></body></html>"))
        self.textEdit.setObjectName(_fromUtf8("textEdit"))
        self.gridLayout_3.addWidget(self.textEdit, 2, 0, 1, 1)
        self.settGroup = QtGui.QGroupBox(self.centralWidget)
        self.settGroup.setObjectName(_fromUtf8("settGroup"))
        self.gridLayout_2 = QtGui.QGridLayout(self.settGroup)
        self.gridLayout_2.setObjectName(_fromUtf8("gridLayout_2"))
        self.label_5 = QtGui.QLabel(self.settGroup)
        self.label_5.setObjectName(_fromUtf8("label_5"))
        self.gridLayout_2.addWidget(self.label_5, 0, 0, 1, 1)
        self.rotSpin = QtGui.QSpinBox(self.settGroup)
        self.rotSpin.setPrefix(_fromUtf8(""))
        self.rotSpin.setMaximum(90000)
        self.rotSpin.setProperty("value", 100)
        self.rotSpin.setObjectName(_fromUtf8("rotSpin"))
        self.gridLayout_2.addWidget(self.rotSpin, 0, 1, 1, 1)
        self.smtpListButton = QtGui.QPushButton(self.settGroup)
        self.smtpListButton.setObjectName(_fromUtf8("smtpListButton"))
        self.gridLayout_2.addWidget(self.smtpListButton, 1, 0, 1, 2)
        self.checkButton = QtGui.QPushButton(self.settGroup)
        self.checkButton.setObjectName(_fromUtf8("checkButton"))
        self.gridLayout_2.addWidget(self.checkButton, 2, 0, 1, 2)
        self.sendButton = QtGui.QPushButton(self.settGroup)
        self.sendButton.setObjectName(_fromUtf8("sendButton"))
        self.gridLayout_2.addWidget(self.sendButton, 3, 0, 1, 2)
        self.gridLayout_3.addWidget(self.settGroup, 3, 0, 1, 2)
        self.msgBox = QtGui.QGroupBox(self.centralWidget)
        self.msgBox.setObjectName(_fromUtf8("msgBox"))
        self.gridLayout = QtGui.QGridLayout(self.msgBox)
        self.gridLayout.setObjectName(_fromUtf8("gridLayout"))
        self.label = QtGui.QLabel(self.msgBox)
        self.label.setObjectName(_fromUtf8("label"))
        self.gridLayout.addWidget(self.label, 2, 0, 1, 1)
        self.numeLineEdit = QtGui.QLineEdit(self.msgBox)
        self.numeLineEdit.setObjectName(_fromUtf8("numeLineEdit"))
        self.gridLayout.addWidget(self.numeLineEdit, 3, 0, 1, 1)
        self.label_3 = QtGui.QLabel(self.msgBox)
        self.label_3.setObjectName(_fromUtf8("label_3"))
        self.gridLayout.addWidget(self.label_3, 4, 0, 1, 1)
        self.subiectLineEdit = QtGui.QLineEdit(self.msgBox)
        self.subiectLineEdit.setObjectName(_fromUtf8("subiectLineEdit"))
        self.gridLayout.addWidget(self.subiectLineEdit, 5, 0, 1, 1)
        self.histButton = QtGui.QPushButton(self.msgBox)
        self.histButton.setObjectName(_fromUtf8("histButton"))
        self.gridLayout.addWidget(self.histButton, 0, 0, 1, 1)
        self.clearButton = QtGui.QPushButton(self.msgBox)
        self.clearButton.setObjectName(_fromUtf8("clearButton"))
        self.gridLayout.addWidget(self.clearButton, 1, 0, 1, 1)
        self.gridLayout_3.addWidget(self.msgBox, 2, 1, 1, 1)
        MainWindow.setCentralWidget(self.centralWidget)
        self.statusBar = QtGui.QStatusBar(MainWindow)
        self.statusBar.setObjectName(_fromUtf8("statusBar"))
        MainWindow.setStatusBar(self.statusBar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QtGui.QApplication.translate("MainWindow", "MMS", None, QtGui.QApplication.UnicodeUTF8))
        self.label_4.setText(QtGui.QApplication.translate("MainWindow", "Mesaj :", None, QtGui.QApplication.UnicodeUTF8))
        self.settGroup.setTitle(QtGui.QApplication.translate("MainWindow", "Setari", None, QtGui.QApplication.UnicodeUTF8))
        self.label_5.setText(QtGui.QApplication.translate("MainWindow", "Rotatie smtp la:", None, QtGui.QApplication.UnicodeUTF8))
        self.rotSpin.setSuffix(QtGui.QApplication.translate("MainWindow", " trimiteri", None, QtGui.QApplication.UnicodeUTF8))
        self.smtpListButton.setText(QtGui.QApplication.translate("MainWindow", "Adauga Lista SMTP", None, QtGui.QApplication.UnicodeUTF8))
        self.checkButton.setText(QtGui.QApplication.translate("MainWindow", "Verifica SMTP", None, QtGui.QApplication.UnicodeUTF8))
        self.sendButton.setText(QtGui.QApplication.translate("MainWindow", "Trimite", None, QtGui.QApplication.UnicodeUTF8))
        self.msgBox.setTitle(QtGui.QApplication.translate("MainWindow", "Setari Mesaj", None, QtGui.QApplication.UnicodeUTF8))
        self.label.setText(QtGui.QApplication.translate("MainWindow", "De la ( nume ):", None, QtGui.QApplication.UnicodeUTF8))
        self.label_3.setText(QtGui.QApplication.translate("MainWindow", "Subiect:", None, QtGui.QApplication.UnicodeUTF8))
        self.histButton.setText(QtGui.QApplication.translate("MainWindow", "Istoric Mesaje", None, QtGui.QApplication.UnicodeUTF8))
        self.clearButton.setText(QtGui.QApplication.translate("MainWindow", "Curata Istoric", None, QtGui.QApplication.UnicodeUTF8))


class Main(QtGui.QMainWindow):
    def __init__(self):
        QtGui.QMainWindow.__init__(self)
        self.ui=Ui_MainWindow()
        self.ui.setupUi(self)
        self.connectSlots()


    def connectSlots(self):
        self.ui.smtpListButton.clicked.connect(self.smtpDialog)
        self.ui.checkButton.clicked.connect(self.checkSMTP)
        self.ui.sendButton.clicked.connect(self.beginSend)
        self.ui.histButton.clicked.connect(self.showHistory)
        self.ui.clearButton.clicked.connect(self.clearHistory)

    def parseEmails(self):
        self.email_list = []
        emails = self.ui.textEdit.toPlainText()
        emails = emails.strip()
        emails = emails.split('\n')
        for i in emails:
            self.email_list.append(i)

        
    def clearHistory(self):
        try:
            hist_file = open('history.mms', 'w')
        except IOError:
            pass
        hist_file.close()

    def showHistory(self):
        try:
            hist_file = open('history.mms', 'r').read()
        except IOError:
            pass
        self.ui.textarea.setText(str(hist_file))        

    def smtpDialog(self):
        self.fname = QtGui.QFileDialog.getOpenFileName(self , "Open smtp list")
        self.smtp_list = []
        try:
            with open(self.fname, 'r') as foo:
                for l in foo:
                    l = l.strip()
                    self.smtp_list.append(l)
        except IOError:
            pass
        
    def checkSMTP(self):
        self.parseEmails()
        try:
            self.cs = CheckSmtp(self.smtp_list, self.email_list)
            self.cs.finished.connect(self.finishedCheck)

            if not self.cs.isRunning():
                self.cs.exiting = False
                self.cs.start()
        except AttributeError:
            QtGui.QMessageBox.information(self, "MMS Info", "Completeaza toate campurile si adauga listele!")
        
    def finishedCheck(self):
        QtGui.QMessageBox.information(self, "MMS Info", "Mail de verificare trimis la adresa %s!" % self.email_list[0])                  
                         
    def beginSend(self):
        self.parseEmails()
        
        hist_file = open('history.mms', 'a')
        body = self.ui.textarea.toPlainText()
        rotatie = self.ui.rotSpin.value()
        hist_file.write(body)
        hist_file.close()
        try:
            self.se = SendEmail(rotatie, self.smtp_list, self.email_list, self.ui.numeLineEdit.text(),body, self.ui.subiectLineEdit.text())
            self.se.sigsent[int].connect(self.sentMails)
            self.se.finished.connect(self.finished)
            if not self.se.isRunning():
                self.se.exiting = False
                self.se.start()

        except AttributeError,e:
            QtGui.QMessageBox.information(self, "MMS Info", "Completeaza toate campurile si adauga listele!")

    def sentMails(self, sent):
        self.ui.statusBar.showMessage("Mesaje trimise: %s" % sent)
        
    def finished(self):
        QtGui.QMessageBox.information(self, "MMS Info", "Trimitere completa!")

class WorkingSignal(QtCore.QObject):
    sig = QtCore.pyqtSignal(str)
    
    
class CheckSmtp(QtCore.QThread):
    sigsent  = QtCore.pyqtSignal(int)
    def __init__(self, smtp_list, email_list, parent=None):
        super(CheckSmtp, self).__init__(parent)
        
        self.smtp_list = smtp_list
        self.email_list = email_list
        self.signal = WorkingSignal()
        self.exiting = False
        self.sent = 0
        self.msg = MIMEMultipart()

        self.msg.attach(MIMEText("Test Email from MMS!"))
        self.msg.add_header('Subject', "MMS Test Email")
        self.msg['From'] = "mms@mailsender.com"

    def run(self):


        while self.exiting == False:
            for smtp in self.smtp_list:
                
                smtp = smtp.strip()
                smtp = smtp.split(':')
                
               
                self.msg.replace_header('Subject', "MMS Test Email %s@%s" % (smtp[0], smtp[1]))

               
                try:
                    smtpObj = SMTP(smtp[0], smtp[4])
                
                    smtpObj.set_debuglevel(1)
                    smtpObj.starttls()
  
                    smtpObj.login(smtp[2], smtp[3])

                    to = self.email_list[0]
                
             
                    smtpObj.sendmail(smtp[1], to, self.msg.as_string())
                except Exception, e:
                    print e
                    smtpObj.quit()
                    continue
                smtpObj.quit()
                    
                
                self.sent += 1
                self.sigsent.emit(self.sent)
                time.sleep(1)
            self.exiting == True
            break

        
class SendEmail(QtCore.QThread):

    sigsent  = QtCore.pyqtSignal(int)
    
    def __init__(self, rotatie, smtp_list, email_list, _from_name, body, subject, parent = None):
        super(SendEmail, self).__init__(parent) 
        self.smtp = smtp_list
        self.email_list = email_list
        self._from_name = _from_name
        self.subject = subject
        self.body = body
        self.contor = 0
        self.rotatie = rotatie
        self.msg = MIMEMultipart()
        self.msg['Subject'] = self.subject
        self.msg['From'] = self._from_name
        self.msg.attach(MIMEText(self.body))
        self.exiting = False
        self.signal = WorkingSignal()
        self.rlist = []

    def run(self):
        resend = False
        email_copy = copy.copy(self.email_list)
        while self.exiting == False:
            for s in itertools.cycle(self.smtp):
                s = s.strip()
                s = s.split(':')
                for i in range(self.rotatie):
                    
                    email = email_copy.pop()

                    try:
                        smtp = SMTP(s[0], s[4])
                        smtp.set_debuglevel(1)
                        smtp.starttls()
                        smtp.login(s[2], s[3])
                        smtp.sendmail(s[1], email, self.msg.as_string())
                        if resend:
                            for e in self.rlist:
                                print 'Resending to address: %s' % e
                                smtp.sendmail(s[1], e, self.msg.as_string())
                                self.rlist.remove(e)

                            resend = False
                    except Exception , e:
                        print '----------------------------'
                        print e
                        print '-------------------------------'
                        resend = True
                        self.rlist.append(email)
                        smtp.quit()
                        continue
                    smtp.quit()
                    self.contor += 1
                    self.sigsent.emit(self.contor)
        
            self.exiting == True
            break
           
        
if __name__ == "__main__":
    import sys
    app = QtGui.QApplication(sys.argv)
    MainWindow = Main()
    MainWindow.show()
    sys.exit(app.exec_())
