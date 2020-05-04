#!/usr/bin/python3

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
import sys
import os
import subprocess
from PyQt5 import QtCore, QtWidgets


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
        MainWindow.resize(435, 326)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)

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
        self.menuOpciones.addAction(self.actionA_adir_ruta_LibreOffice)
        self.menuOpciones.addAction(self.actionA_adir_ruta_Latex)
        self.menuOpciones.addAction(self.actionImportarAlumnos)
        self.menubar.addAction(self.menuOpciones.menuAction())

        self.OfficeRoute = False
        self.LatexRoute = False
        self.FolderAlumnos = False
        self.nameSoffice = ""
        self.nameLatex = ""
        self.nameFolderAlumnos = ""
        self.markCheckBox = False

        self.helpButton = QtWidgets.QAction(MainWindow)
        self.menubar.addAction(self.helpButton)
        self.helpButton.setObjectName("Ayuda")

        # self.png_files = ls1(os.path.abspath(os.getcwd()),True)
        self.actionA_adir_ruta_LibreOffice.triggered.connect(self.addOfficeRoute)
        self.actionA_adir_ruta_Latex.triggered.connect(self.addLatexRoute)
        self.actionImportarAlumnos.triggered.connect(self.getAlumnos)
        self.pushButton.clicked.connect(self.generate)
        self.helpButton.triggered.connect(self.showHelp)

        self.retranslateUi(MainWindow)

        self.checkPDF = QtWidgets.QCheckBox("Generar documento CSV de comprobación", self.centralwidget)
        self.checkPDF.setObjectName("checkPDF")
        self.checkPDF.stateChanged.connect(self.clickBox)

        self.pushButton.setMaximumSize(1000, 31)

        self.h_layout = QtWidgets.QVBoxLayout()
        self.h_layout.setAlignment(QtCore.Qt.AlignCenter)
        self.h_layout.setSpacing(50)

        self.h_layout.addWidget(self.checkPDF)
        self.h_layout.addWidget(self.pushButton, QtCore.Qt.AlignHCenter)
        self.centralwidget.setLayout(self.h_layout)

        self.centralwidget.setWindowTitle("ToR Conversor")

        self.pushButton.setStyleSheet("background-color:rgb(220,128,128)")

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

    def clickBox(self, state):

        if state == QtCore.Qt.Checked:
            self.markCheckBox = True
        else:
            self.markCheckBox = False

    def showHelp(self):
        confirm = QtWidgets.QMessageBox()
        confirm.setIcon(QtWidgets.QMessageBox.Information)
        confirm.setWindowTitle("Ayuda")
        confirm.setIcon(QtWidgets.QMessageBox.Information)
        confirm.setText(
            "Para obtener las calificaciones finales se deben realizar los siguientes pasos:\n\n 1. Se selecciona la ruta de Soffice "
            "(Opciones -> Añadir ruta Soffice) y se selecciona el archivo .bin o .com (Windows)\n\n 2.  Se selecciona la ruta del "
            "generador de documentos (Opciones -> Añadir ruta generador de documentos) y se selecciona el archivo ejecutable .exe"
            "\n\n 3.  Se selecciona la carpeta que contiene los ficheros de los alumnos (Opciones -> importar Alumnos) \n\n"
            "4. Se pulsa GENERAR, se pedirá la tabla de equivalencia de calificaciones y se ejecutará el software")
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
            self.OfficeRoute = False

    def addLatexRoute(self):
        filename = QtWidgets.QFileDialog()
        filename.setWindowTitle("Añadir ruta PdfLatex")
        name = filename.getOpenFileName(filter="Todos los archivos(*.*)")
        if name[0]:
            self.LatexRoute = True
            self.nameLatex = name[0]
            confirm = QtWidgets.QMessageBox()
            confirm.setIcon(QtWidgets.QMessageBox.Information)
            confirm.setWindowTitle("Añadir ruta PdfLatex")
            confirm.setText("Generador de documentos cargado con éxito.")
            confirm.exec()
        else:
            self.LatexRoute = False

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

    def generate(self):

        if self.OfficeRoute == False:
            confirm = QtWidgets.QMessageBox()
            confirm.setIcon(QtWidgets.QMessageBox.Critical)
            confirm.setWindowTitle("Generar calificaciones finales")
            confirm.setText("ERROR : No se ha añadido la ruta de Soffice.")
            confirm.exec()

        elif self.LatexRoute == False:
            confirm = QtWidgets.QMessageBox()
            confirm.setIcon(QtWidgets.QMessageBox.Critical)
            confirm.setWindowTitle("Generar calificaciones finales")
            confirm.setText("ERROR : No se ha añadido la ruta del generador de documentos.")
            confirm.exec()

        elif self.FolderAlumnos == False:
            confirm = QtWidgets.QMessageBox()
            confirm.setIcon(QtWidgets.QMessageBox.Critical)
            confirm.setWindowTitle("Generar calificaciones finales")
            confirm.setText("ERROR : No se ha añadido la carpeta con los datos de los alumnos")
            confirm.exec()

        else:
            filename = QtWidgets.QFileDialog()
            filename.setWindowTitle("Añadir tabla de equivalencia de calificaciones")
            name = filename.getOpenFileName(filter="Archivo CSV (*.csv)")
            if name[0]:

                files = ls1(self.nameFolderAlumnos, False)
                folder_primary = "ALUMNOS"
                primary_path = os.path.join(os.path.abspath(os.getcwd()), folder_primary)
                try:
                    os.stat(primary_path)
                except:
                    os.mkdir(primary_path)

                FNULL = open(os.devnull, 'w')

                for file in files:
                    try:
                        os.stat(os.path.join(primary_path, file[:-4]))
                    except:
                        os.mkdir(os.path.join(primary_path, file[:-4]))

                    folder_path = os.path.join(primary_path, file[:-4])
                    cmd = '"%s" --headless --convert-to csv:"Text - txt - csv (StarCalc)":"59,34,76,1" --outdir "%s" "%s"' % (
                    self.nameSoffice, folder_path, os.path.join(self.nameFolderAlumnos,file))
                    subprocess.call(cmd, stdout=FNULL)

                    csv_file = os.path.join(folder_path, file[:-3] + "csv")
                    personalData, ToR = readToR(csv_file)

                    destination = personalData[UNIV_COLUMN].upper()

                    american, ToR = tor.parseToR(ToR)

                    # 2. Parse and prepare equivalences table.
                    raw_destination, raw_origin = readData(name[0], HOME, destination)

                    x, aliasx, y, aliasy = tor.expandScores(raw_origin, raw_destination, american)

                    # 3. Expand the table to score suggestions for each destination subject
                    ToR = tor.extendToR(ToR, x, aliasx, y, aliasy, american)

                    # Generate debug information

                    if self.markCheckBox:
                        exportCSVToR(personalData, ToR, os.path.join(folder_path, file[:-4] + "--debug_mode.csv"))

                    # 4. Load the LaTeX document
                    f = open(os.path.join(os.path.abspath(os.getcwd()), "template01.tex"), "r", encoding='utf8')
                    text = f.read()
                    f.close()

                    for d in personalData:
                        personalData[d] = personalData[d].replace("\\", "/").replace("_", "\_").replace("$", "\$")
                        text = text.replace("[[" + str(d) + "]]", personalData[d].upper())
                        text = text.replace("{{" + str(d) + "}}", personalData[d])

                    # 5. Add the final calification:
                    table = ""
                    for block in ToR:
                        table += "\\hline \\hline \n"
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
                            try:
                                sOrig = ToR[block][1][i][0]
                            except:
                                sOrig = ""
                            try:
                                sDst = ToR[block][0][i][0]
                            except:
                                sDst = ""
                            try:
                                crOrig = ToR[block][1][i][1]
                            except:
                                crOrig = ""
                            try:
                                crDst = ToR[block][0][i][1]
                            except:
                                crDst = ""
                            try:
                                scoreDst = ToR[block][0][i][2]
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
                                                                                                     scoreDst, sOrig,
                                                                                                     crOrig, crLabel)
                            elif sOrig == "":
                                table += " {} & {} &  & {} & {} & {} &  & \\\\ \\hline \n".format(sDst, crDst, scoreDst,
                                                                                                  sOrig, crOrig)
                            else:
                                table += " {} & {} &  & {} & {} & {} &  & {:.1f} {} \\\\ \\hline \n".format(sDst, crDst,
                                                                                                            scoreDst,
                                                                                                            sOrig,
                                                                                                            crOrig,
                                                                                                            score,
                                                                                                            crLabel)

                    text = text.replace("{{:SUBJECT-LIST:}}", table)
                    f = open(os.path.join(folder_path, file[:-3] + "tex"), "w", encoding='utf8')
                    f.write(text)
                    f.close()

                    cmd = '"%s" -output-directory="%s" "%s"' % (
                    self.nameLatex, folder_path, os.path.join(folder_path, file[:-3] + "tex"))
                    subprocess.call(cmd, stdout=FNULL)

                confirm = QtWidgets.QMessageBox()
                confirm.setIcon(QtWidgets.QMessageBox.Information)
                confirm.setWindowTitle("Generar calificaciones finales")
                confirm.setText("Calificaciones generadas con éxito.")
                confirm.exec()

            else:
                pass

    ########################################################################################################################
    ########################################################################################################################


app = QtWidgets.QApplication([])
ui = Ui_MainWindow()
MainWindow = QtWidgets.QMainWindow()
ui.setupUi(MainWindow)
MainWindow.show()
sys.exit(app.exec_())
