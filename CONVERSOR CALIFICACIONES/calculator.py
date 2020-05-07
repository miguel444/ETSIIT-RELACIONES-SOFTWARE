#!/usr/bin/env/python3

"""
This file is part of ToR Conversor.

    ToR Conversor is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    ToR Conversor is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with ToR Conversor.  If not, see <https://www.gnu.org/licenses/>.
"""

import csvh
import tor
import csv
import sys
import os
from subprocess import call
from PyQt5 import QtCore, QtWidgets, QtGui
import pylatex

HOME = "España"
UNIV_COLUMN = "Código VICERRECTORADO donde se han cursado los estudios:"

HOME = HOME.strip().upper()
UNIV_COLUMN = UNIV_COLUMN.strip()


def exportCSVToR(personalData, ToR, fileName):
    csv_data = []
    for d in personalData:
        csv_data.append([d, personalData[d]])
    csv_data.append([])
    idv = 1
    for d in ToR:
        csv_data.append(["Bloque:", idv])
        csv_data.append(
            ["", "Asignatura Destino", "Créditos", "Nota Destino", "Sugerencia Origen", "Min. Origen", "Máx. Origen",
             "Min. Destino", "Máx. Destino", "Alias"])
        for subject in ToR[d][0]:
            csv_data.append([""] + subject)
        csv_data.append([])
        csv_data.append(["", "Asignatura Origen", "Créditos"])
        for subject in ToR[d][1]:
            csv_data.append([""] + subject)
        csv_data.append([])
        idv += 1

    csvh.exportRawCSVData(fileName, csv_data)


def readData(fileName, origin, destination):
    eq_data = csvh.importRawCSVData(fileName)

    data = {}
    for d in eq_data:
        d[0] = d[0].strip().upper()
        if str(d[0])[0] == "#":
            pass
        else:
            data[d[0]] = d[1:]

    r = data[destination]
    raw_destination = []
    for d in r:
        raw_destination.append(str(d))
    raw_origin = data[origin]
    return (raw_destination, raw_origin)


def readToR(fileName):
    ToR = csvh.importRawCSVData(fileName)

    subjectData = []
    readSubjects = False
    personalData = {}
    for d in ToR:
        d[0] = d[0].replace("\n", " ").replace("\r", " ").strip()
        if readSubjects:
            subjectData.append(d)
        else:
            if d[0] == "Asignatura":
                readSubjects = True
            elif (d[0] != "" and str(d[0])[0] == "#") or d[0] == "":
                pass
            else:
                personalData[d[0]] = str(d[1]).strip()

    return (personalData, subjectData)


def ls1(path, option):
    """Función que obtiene todos los archivos del directorio pasado como argumento a la función , comprueba si son
    ficheros csv y finalmente devolvemos una lista con el nombre de todos los archivos encontrados.

    :param path: directorio del que vamos a obtener todos los archivos

    :return: lista con los nombres de los diferentes archivos encontrados en el directorio
    """
    if option:
        return [obj for obj in os.listdir(path) if os.path.isfile(os.path.join(path, obj)) and obj[-3:] in ['png']]
    else:
        return [obj for obj in os.listdir(path) if os.path.isfile(os.path.join(path, obj)) and obj[-3:] in ['ods']]


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1050, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)

        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(0, 1, 771, 551))
        self.tableWidget.setRowCount(30)
        self.tableWidget.setColumnCount(30)
        self.tableWidget.setObjectName("tableWidget")

        self.pushButton.setObjectName("pushButton")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 435, 26))
        self.menubar.setObjectName("menubar")
        self.menuOpciones = QtWidgets.QMenu(self.menubar)
        self.menuOpciones.setObjectName("menuOpciones")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionA_adir_ruta_LibreOffice = QtWidgets.QAction(MainWindow)
        self.actionA_adir_ruta_LibreOffice.setObjectName("actionA_adir_ruta_LibreOffice")
        self.actionA_adir_ruta_Latex = QtWidgets.QAction(MainWindow)
        self.actionA_adir_ruta_Latex.setObjectName("actionA_adir_ruta_Latex")
        self.actionImportarAlumnos = QtWidgets.QAction(MainWindow)
        self.actionImportarAlumnos.setObjectName("Importar alumnos")
        self.actionAñadirTablaEquivalencias = QtWidgets.QAction(MainWindow)
        self.menuOpciones.addAction(self.actionA_adir_ruta_LibreOffice)
        self.menuOpciones.addAction(self.actionA_adir_ruta_Latex)
        self.menuOpciones.addAction(self.actionImportarAlumnos)
        self.menuOpciones.addAction(self.actionAñadirTablaEquivalencias)
        self.menubar.addAction(self.menuOpciones.menuAction())

        self.select_country = QtWidgets.QComboBox(self.centralwidget)

        self.select_country.setObjectName("selectCountry")

        self.texto_select = QtWidgets.QLabel("Selecciona el país de destino :")

        self.FolderAlumnos = False
        self.OfficeRoute = False
        self.nameSoffice = "soffice"
        self.nameLatex = "pdflatex"
        self.nameFolderAlumnos = ""
        self.markCheckBox = False
        self.mode_manual = False
        self.personalData = None
        self.Tor = None
        self.loop = None

        self.tabla_equivalencias = "DATA/data.csv"
        self.getCountries()

        self.conversor_individual = QtWidgets.QAction(MainWindow)
        self.menubar.addAction(self.conversor_individual)
        self.conversor_individual.setObjectName("Conversor calificaciones manual")

        self.helpButton = QtWidgets.QAction(MainWindow)
        self.menubar.addAction(self.helpButton)
        self.helpButton.setObjectName("Ayuda")

        self.confirmButton = QtWidgets.QPushButton(self.centralwidget)

        # self.png_files = ls1(os.path.abspath(os.getcwd()),True)
        self.actionA_adir_ruta_LibreOffice.triggered.connect(self.addOfficeRoute)
        self.actionA_adir_ruta_Latex.triggered.connect(self.addLatexRoute)
        self.actionImportarAlumnos.triggered.connect(self.getAlumnos)
        self.actionAñadirTablaEquivalencias.triggered.connect(self.addTable)
        self.pushButton.clicked.connect(self.generate)
        self.helpButton.triggered.connect(self.showHelp)
        self.conversor_individual.triggered.connect(self.controller)

        self.retranslateUi(MainWindow)

        self.checkPDF = QtWidgets.QCheckBox("Comprobar información antes de generar PDF", self.centralwidget)
        self.checkPDF.setObjectName("checkPDF")
        self.checkPDF.stateChanged.connect(self.clickBox)

        self.pushButton.setMaximumHeight(31)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        self.h_layout = QtWidgets.QVBoxLayout()
        self.h_layout.setAlignment(QtCore.Qt.AlignCenter)
        self.h_layout.setSpacing(25)
        self.h_layout.addWidget(self.texto_select)
        self.h_layout.addWidget(self.select_country)
        self.h_layout.addWidget(self.tableWidget)
        self.h_layout.addWidget(self.checkPDF)
        self.h_layout.addWidget(self.confirmButton)
        self.h_layout.addWidget(self.pushButton, QtCore.Qt.AlignHCenter)
        self.centralwidget.setLayout(self.h_layout)

        self.centralwidget.setWindowTitle("ToR Conversor")

        self.pushButton.setStyleSheet("background-color:rgb(220,128,128)")
        self.confirmButton.setStyleSheet("background-color:rgb(220,128,128)")
        self.select_country.hide()
        self.texto_select.hide()
        self.confirmButton.hide()

        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "ToR Conversor"))
        self.pushButton.setText(_translate("MainWindow", "Generar"))
        self.menuOpciones.setTitle(_translate("MainWindow", "Opciones"))
        self.actionA_adir_ruta_LibreOffice.setText(_translate("MainWindow", "Añadir ruta Soffice"))
        self.actionA_adir_ruta_Latex.setText(_translate("MainWindow", "Añadir ruta generador de documentos"))
        self.actionImportarAlumnos.setText((_translate("MainWindow", "Importar alumnos")))
        self.helpButton.setText((_translate("MainWindow", "Ayuda")))
        self.conversor_individual.setText((_translate("MainWindow", "Conversor calificaciones manual")))
        self.confirmButton.setText("Confirmar datos")
        self.actionAñadirTablaEquivalencias.setText("Añadir tabla de equivalencias")

    def clickBox(self, state):

        if state == QtCore.Qt.Checked:
            self.markCheckBox = True
        else:
            self.markCheckBox = False

    def controller(self):
        if not self.mode_manual:
            self.convertSingle()
        else:
            self.tableWidget.clearContents()
            self.tableWidget.setColumnCount(30)
            self.select_country.hide()
            self.checkPDF.show()
            self.texto_select.hide()
            self.conversor_individual.setText("Conversor calificaciones manual")
            self.mode_manual = False
            self.menuOpciones.menuAction().setVisible(True)

    def convertSingle(self):
        self.mode_manual = True
        self.texto_select.show()
        self.select_country.show()
        self.checkPDF.hide()
        self.conversor_individual.setText("Conversor calificaciones automático")
        self.menuOpciones.menuAction().setVisible(False)

        switcher = {
            0: "Asignatura destino",
            1: "ECTS",
            2: "otro",
            3: "ID_BLOQUE",
            4: "Calificación Destino",
            5: "Asignatura origen",
            6: "ECTS",
            7: "LRU",
            8: "ID_BLOQUE",
            9: "Calificación Origen",
            10: "CRLabel"
        }

        self.tableWidget.setColumnCount(11)
        for i in range(self.tableWidget.rowCount()):
            for p in range(self.tableWidget.columnCount()):
                if i == 0:
                    item = QtWidgets.QTableWidgetItem(switcher.get(p))
                    if p == 0 or p == 5:
                        item.setBackground(QtGui.QColor(123, 193, 233))
                    else:
                        item.setBackground(QtGui.QColor(255, 128, 128))
                    self.tableWidget.setItem(i, p, item)

    def getCountries(self):

        with open(self.tabla_equivalencias, encoding='utf-8-sig') as File:
            reader = csv.reader(File, delimiter=';')
            fields_form = []
            for row in reader:
                # Comprobamos que no esté vacía y que no sea comentarios
                if row and not row[0].strip() in (None, "") and row[0][0] != '#':
                    # Comprobamos si no hemos guardado dicho campo de forumulario y le asignamos una lista
                    if not row[0].strip() in fields_form:
                        fields_form.append(row[0].strip())

        self.select_country.addItems(fields_form)

    def showHelp(self):
        confirm = QtWidgets.QMessageBox()
        confirm.setIcon(QtWidgets.QMessageBox.Information)
        confirm.setWindowTitle("Ayuda")
        confirm.setIcon(QtWidgets.QMessageBox.Information)
        confirm.setText(
            "Para obtener las calificaciones finales se deben realizar los siguientes pasos:\n\n 1. Se selecciona la ruta de Soffice "
            "(Opciones -> Añadir ruta Soffice) y se selecciona el archivo .bin o .com (Windows)\n\n 2. (OPCIONAL) Se selecciona la ruta del "
            "generador de documentos (Opciones -> Añadir ruta generador de documentos) y se selecciona el archivo ejecutable .exe"
            "\n\n 3.  Se selecciona la carpeta que contiene los ficheros de los alumnos (Opciones -> importar Alumnos) \n\n"
            "4. (OPCIONAL) Se selecciona la tabla de equivalencias (Opciones -> Añadir tabla de equivalencias) y se selecciona un fichero .csv\n\n"
            "5. Se pulsa GENERAR, se pedirá la tabla de equivalencia de calificaciones y se ejecutará el software")
        confirm.exec()

    def addOfficeRoute(self):
        filename = QtWidgets.QFileDialog()
        filename.setWindowTitle("Añadir ruta Soffice")

        name = filename.getOpenFileName(filter="Todos los archivos(*.*)")
        if name[0]:
            self.OfficeRoute = True
            self.nameSoffice = name[0]
            confirm = QtWidgets.QMessageBox()
            confirm.setIcon(QtWidgets.QMessageBox.Information)
            confirm.setWindowTitle("Añadir ruta Soffice")
            confirm.setText("Convetidor de formato cargado con éxito.")
            confirm.exec()
        else:
            pass

    def addLatexRoute(self):
        filename = QtWidgets.QFileDialog()
        filename.setWindowTitle("Añadir ruta PdfLatex")
        name = filename.getOpenFileName(filter="Todos los archivos(*.*)")
        if name[0]:

            self.nameLatex = name[0]
            confirm = QtWidgets.QMessageBox()
            confirm.setIcon(QtWidgets.QMessageBox.Information)
            confirm.setWindowTitle("Añadir ruta PdfLatex")
            confirm.setText("Generador de documentos cargado con éxito.")
            confirm.exec()
        else:
            pass

    def getAlumnos(self):
        fileName = QtWidgets.QFileDialog()
        fileName.setWindowTitle("Seleccionar carpeta con datos de los alumnos")
        folder = fileName.getExistingDirectory()

        if folder:
            self.nameFolderAlumnos = folder
            self.FolderAlumnos = True
            confirm = QtWidgets.QMessageBox()
            confirm.setIcon(QtWidgets.QMessageBox.Information)
            confirm.setWindowTitle("Importar alumnos")
            confirm.setText("Alumnos importados con éxito.")
            confirm.exec()
        else:
            self.FolderAlumnos = False

    def addTable(self):
        filename = QtWidgets.QFileDialog()
        filename.setWindowTitle("Añadir tabla de equivalencia de calificaciones")
        name = filename.getOpenFileName(filter="Archivo CSV (*.csv)")

        if name[0]:
            self.tabla_equivalencias = name[0]
        else:
            pass

    def generate(self):
        if self.mode_manual:
            self.generate_manual()

        else:

            if self.FolderAlumnos == False:
                confirm = QtWidgets.QMessageBox()
                confirm.setIcon(QtWidgets.QMessageBox.Critical)
                confirm.setWindowTitle("Generar calificaciones finales")
                confirm.setText("ERROR : No se ha añadido la carpeta con los datos de los alumnos")
                confirm.exec()
            elif self.OfficeRoute == False:
                confirm = QtWidgets.QMessageBox()
                confirm.setIcon(QtWidgets.QMessageBox.Critical)
                confirm.setWindowTitle("Generar calificaciones finales")
                confirm.setText("ERROR : No se ha añadido la ruta de soffice")
                confirm.exec()


            else:

                files = ls1(self.nameFolderAlumnos, False)
                folder_primary = "ALUMNOS"
                primary_path = os.path.join(os.path.abspath(os.getcwd()), folder_primary)
                try:
                    os.stat(primary_path)
                except:
                    os.mkdir(primary_path)

                for file in files:
                    try:
                        os.stat(os.path.join(primary_path, file[:-4]))
                    except:
                        os.mkdir(os.path.join(primary_path, file[:-4]))

                    FNULL = open(os.devnull, 'w')
                    folder_path = os.path.join(primary_path, file[:-4])
                    cmd = '"%s" --headless --convert-to csv:"Text - txt - csv (StarCalc)":"59,34,76,1" --outdir "%s" "%s"' % (
                        self.nameSoffice, folder_path, os.path.join(self.nameFolderAlumnos, file))
                    call(cmd)

                    csv_file = os.path.join(folder_path, file[:-3] + "csv")
                    personalData, ToR = readToR(csv_file)

                    destination = personalData[UNIV_COLUMN].upper()

                    american, ToR = tor.parseToR(ToR)

                    # 2. Parse and prepare equivalences table.
                    raw_destination, raw_origin = readData(self.tabla_equivalencias, HOME, destination)

                    x, aliasx, y, aliasy = tor.expandScores(raw_origin, raw_destination, american)

                    # 3. Expand the table to score suggestions for each destination subject
                    ToR = tor.extendToR(ToR, x, aliasx, y, aliasy, american)

                    self.personalData = personalData
                    self.Tor = ToR

                    # Generate debug information

                    if self.markCheckBox:
                        self.checkPDF.hide()
                        self.pushButton.hide()
                        self.confirmButton.show()
                        self.show_info_check(personalData, ToR)
                        loop = QtCore.QEventLoop()
                        self.loop = loop
                        self.confirmButton.clicked.connect(self.check_info_show)
                        self.loop.exec()
                        self.tableWidget.clearContents()

                        # exportCSVToR(personalData, ToR, os.path.join(folder_path, file[:-4] + "--debug_mode.csv"))

                    # 4. Load the LaTeX document
                    f = open(os.path.join(os.path.abspath(os.getcwd()), "TEX/template01.tex"), "r", encoding='utf8')
                    text = f.read()
                    f.close()

                    for d in self.personalData:
                        self.personalData[d] = self.personalData[d].replace("\\", "/").replace("_", "\_").replace("$",
                                                                                                                  "\$")
                        text = text.replace("[[" + str(d) + "]]", self.personalData[d].upper())
                        text = text.replace("{{" + str(d) + "}}", self.personalData[d])

                    # 5. Add the final calification:
                    table = ""

                    for block in self.Tor:
                        table += "\\hline \\hline \n"
                        maxCr = 0
                        fail = -1
                        score = 0
                        n = len(self.Tor[block][0])

                        for i in range(n):
                            maxCr += self.Tor[block][0][i][1]

                        for i in range(n):
                            if self.Tor[block][0][i][3] < 5:
                                fail = self.Tor[block][0][i][3]
                                score = fail
                                break
                            else:
                                score += self.Tor[block][0][i][3] * (self.Tor[block][0][i][1] / maxCr)
                        for i in range(max(n, len(self.Tor[block][1]))):
                            try:
                                sOrig = self.Tor[block][1][i][0]
                            except:
                                sOrig = ""
                            try:
                                sDst = self.Tor[block][0][i][0]
                            except:
                                sDst = ""
                            try:
                                crOrig = self.Tor[block][1][i][1]
                            except:
                                crOrig = ""
                            try:
                                crDst = self.Tor[block][0][i][1]
                            except:
                                crDst = ""
                            try:
                                scoreDst = self.Tor[block][0][i][2]
                            except:
                                scoreDst = ""

                            score = float("{:.1f}".format(score))
                            if score < 0:
                                crLabel = "NO PRESENTADO"
                            elif score < 5:
                                crLabel = "(SUSPENSO)"
                            elif score < 7:
                                crLabel = "(APROBADO)"
                            elif score < 9:
                                crLabel = "(NOTABLE)"
                            elif score < 9.5:
                                crLabel = "(SOBRESALIENTE)"
                            else:
                                crLabel = "(MATRÍCULA)"

                            if score < 0 and sOrig != "":
                                table += " {} & {} &  & {} & {} & {} &  & {} \\\\ \\hline \n".format(sDst, crDst,
                                                                                                     scoreDst,
                                                                                                     sOrig,
                                                                                                     crOrig,
                                                                                                     crLabel)
                            elif sOrig == "":
                                table += " {} & {} &  & {} & {} & {} &  & \\\\ \\hline \n".format(sDst, crDst,
                                                                                                  scoreDst,
                                                                                                  sOrig, crOrig)
                            else:
                                table += " {} & {} &  & {} & {} & {} &  & {:.1f} {} \\\\ \\hline \n".format(sDst,
                                                                                                            crDst,
                                                                                                            scoreDst,
                                                                                                            sOrig,
                                                                                                            crOrig,
                                                                                                            score,
                                                                                                            crLabel)

                    text = text.replace("{{:SUBJECT-LIST:}}", table)
                    f = open(os.path.join(folder_path, file[:-3] + "tex"), "w", encoding='utf8')
                    f.write(text)
                    f.close()

                    cmd = '"%s"  -output-directory="%s" "%s"' % (
                        self.nameLatex, folder_path, os.path.join(folder_path, file[:-3] + "tex"))
                    call(cmd)
                    FNULL.close()

                confirm = QtWidgets.QMessageBox()
                confirm.setIcon(QtWidgets.QMessageBox.Information)
                confirm.setWindowTitle("Generar calificaciones finales")
                confirm.setText("Calificaciones generadas con éxito.")
                confirm.exec()

                self.checkPDF.show()
                self.confirmButton.hide()
                self.pushButton.show()

    def generate_manual(self):

        data = []

        bloques = []
        for i in range(1, self.tableWidget.rowCount()):
            row = []
            for p in range(self.tableWidget.columnCount() - 2):
                item = self.tableWidget.item(i, p)
                if item is not None and item.text():
                    if p == 8:
                        bloques.append(i)
                    row.append(item.text())
                else:
                    row.append("")

            data.append(row)

        american, ToR = tor.parseToR(data)
        raw_destination, raw_origin = readData(self.tabla_equivalencias, HOME, self.select_country.currentText())
        x, aliasx, y, aliasy = tor.expandScores(raw_origin, raw_destination, american)

        # 3. Expand the table to score suggestions for each destination subject
        ToR = tor.extendToR(ToR, x, aliasx, y, aliasy, american)

        for block in ToR:
            maxCr = 0
            fail = -1
            score = 0
            n = len(ToR[block][0])
            for i in range(n):
                maxCr += ToR[block][0][i][1]

            for i in range(n):
                if ToR[block][0][i][3] < 5:
                    fail = ToR[block][0][i][3]
                    score = fail
                    break
                else:
                    score += ToR[block][0][i][3] * (ToR[block][0][i][1] / maxCr)
            for i in range(max(n, len(ToR[block][1]))):

                score = float("{:.1f}".format(score))
                if score < 0:
                    crLabel = "NO PRESENTADO"
                elif score < 5:
                    crLabel = "(SUSPENSO)"
                elif score < 7:
                    crLabel = "(APROBADO)"
                elif score < 9:
                    crLabel = "(NOTABLE)"
                elif score < 9.5:
                    crLabel = "(SOBRESALIENTE)"
                else:
                    crLabel = "(MATRÍCULA)"

            cont = bloques.pop(0)
            self.tableWidget.setItem(cont, 9, QtWidgets.QTableWidgetItem(str(score)))
            self.tableWidget.setItem(cont, 10, QtWidgets.QTableWidgetItem(crLabel))

    def show_info_check(self, personalData, ToR):

        self.tableWidget.setRowCount(0)

        for d in personalData:
            rowPosition = self.tableWidget.rowCount()
            self.tableWidget.insertRow(rowPosition)
            item = QtWidgets.QTableWidgetItem(d)
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 0, item)
            self.tableWidget.setItem(rowPosition, 1, QtWidgets.QTableWidgetItem(personalData[d]))

        self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
        idv = 1
        for d in ToR:
            rowPosition = self.tableWidget.rowCount()
            self.tableWidget.insertRow(rowPosition)

            item = QtWidgets.QTableWidgetItem("Bloque")
            item.setBackground(QtGui.QColor(123, 193, 233))
            self.tableWidget.setItem(rowPosition, 0, item)
            item = QtWidgets.QTableWidgetItem(str(idv))
            item.setBackground(QtGui.QColor(123, 193, 233))
            self.tableWidget.setItem(rowPosition, 1, item)
            for blue in range(2, 10):
                item = QtWidgets.QTableWidgetItem("")
                item.setBackground(QtGui.QColor(123, 193, 233))
                self.tableWidget.setItem(rowPosition, blue, item)

            rowPosition = self.tableWidget.rowCount()
            self.tableWidget.insertRow(rowPosition)
            item = QtWidgets.QTableWidgetItem("Asignatura Destino")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 1, item)

            item = QtWidgets.QTableWidgetItem("Créditos")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 2, item)

            item = QtWidgets.QTableWidgetItem("Nota Destino")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 3, item)

            item = QtWidgets.QTableWidgetItem("Sugerencia Origen")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 4, item)

            item = QtWidgets.QTableWidgetItem("Min. Origen")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 5, item)

            item = QtWidgets.QTableWidgetItem("Max. Origen")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 6, item)

            item = QtWidgets.QTableWidgetItem("min. Destino")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 7, item)

            item = QtWidgets.QTableWidgetItem("Max. Destino")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 8, item)

            item = QtWidgets.QTableWidgetItem("Alias")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 9, item)

            for subject in ToR[d][0]:
                rowPosition = self.tableWidget.rowCount()
                self.tableWidget.insertRow(rowPosition)
                for col in range(len(subject)):
                    self.tableWidget.setItem(rowPosition, col + 1, QtWidgets.QTableWidgetItem(str(subject[col])))

            self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
            rowPosition = self.tableWidget.rowCount()
            self.tableWidget.insertRow(rowPosition)

            item = QtWidgets.QTableWidgetItem("Asignatura Origen")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 1, item)

            item = QtWidgets.QTableWidgetItem("Créditos")
            item.setBackground(QtGui.QColor(228, 195, 195))
            self.tableWidget.setItem(rowPosition, 2, item)

            for subject in ToR[d][1]:
                rowPosition = self.tableWidget.rowCount()
                self.tableWidget.insertRow(rowPosition)
                for col in range(len(subject)):
                    self.tableWidget.setItem(rowPosition, col + 1, QtWidgets.QTableWidgetItem(str(subject[col])))

            self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
            idv += 1

    def check_info_show(self):

        idv = -1

        for d in self.personalData:
            idv += 1
            item = self.tableWidget.item(idv, 1)
            if item is not None and item.text():
                self.personalData[d] = item.text()
            else:
                self.personalData[d] = ""

        for d in self.Tor:
            idv += 3

            for subject in range(len(self.Tor[d][0])):

                idv += 1

                for col in range(len(self.Tor[d][0][subject])):
                    item = self.tableWidget.item(idv, col + 1)
                    if item is not None and item.text():

                        if col == 1:

                            self.Tor[d][0][subject][col] = int(item.text())
                        elif col == 8 or col == 0 or not item.text()[0].isdigit():

                            self.Tor[d][0][subject][col] = item.text()
                        else:
                            self.Tor[d][0][subject][col] = float(item.text())
                    else:
                        self.Tor[d][0][subject][col] = ""

            idv += 2

            for subject in range(len(self.Tor[d][1])):
                idv += 1
                for col in range(len(self.Tor[d][1][subject])):
                    item = self.tableWidget.item(idv, col + 1)
                    if item is not None and item.text():
                        if col == 1:
                            self.Tor[d][1][subject][col] = int(item.text())
                        else:
                            self.Tor[d][1][subject][col] = item.text()
                    else:
                        self.Tor[d][1][subject][col] = ""

        self.loop.exit()


########################################################################################################################
########################################################################################################################


app = QtWidgets.QApplication([])
ventana = Ui_MainWindow()
main_window = QtWidgets.QMainWindow()
ventana.setupUi(main_window)
main_window.show()

sys.exit(app.exec())
