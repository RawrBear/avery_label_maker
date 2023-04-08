from kivy.clock import Clock
from kivy.app import App
from kivy.uix.widget import Widget
from docxtpl import DocxTemplate
from kivy.properties import ObjectProperty
from plyer import filechooser
import pandas as pd
import time as time
import threading
import global_

# TODO: make PROCESS button only available when the process isn't running
# TODO: Fix styling
# TODO: Error handling: The input boxes shouldn't crash the pogram when there's no input or the user clicks out of the selection screen
# TODO: Split the code
# TODO: Improve and refactor this crappy code


def thread(function):
    def wrap(*args, **kwargs):
        t = threading.Thread(target=function, args=args,
                             kwargs=kwargs, daemon=True)
        t.start()

        return t
    return wrap


class MainGrid(Widget):

    dataList = ObjectProperty(None)
    outputDir = ObjectProperty(None)
    outputConsole = ObjectProperty(None)

    listPath = ""
    outputPath = ""
    processing = False

    def updateConsole(self, *kwargs):
        print("UPDATECONSOLE SAYS: ", global_.message)
        self.ids.outputConsole.text = global_.message

    def btnOpenFile(self):
        self.listPath = filechooser.open_file(title="Choose your list..", filters=[
            ("Excel", "*.xlsx")])
        self.ids.dataList.text = self.listPath[0]

    def btnSaveFile(self):
        self.outputPath = filechooser.choose_dir(title="Where do you want to save the output?", filters=[
            ("All Files", "*.*")])
        self.ids.outputDir.text = self.outputPath[0]

    def start_console_thread(self):
        print("Start console thread")
        threading.Thread(target=self.console_thread).start()

    def console_thread(self):
        print("Console thread running")
        self.console_clock = Clock.schedule_interval(self.updateConsole, 1)

    def stop_console_clock(self):
        self.console_clock.cancel()

    def btnGo(self):
        runProcess(self, self.listPath, self.outputPath)
        self.processing == True


# Opens the template file
doc = DocxTemplate("template.docx")

# Sets the main context to be an empty dictionary
main_context = {}

# Is used to track the sheets.
# It logs every time the save function is called and sets that as the sheet number
counter = 0

# Used to store the index of the current row
# This is used in process method
personIndex = 0


class Person:
    def __init__(self, name, address, city, state, zip):
        self.name = name
        self.address = address
        self.city = city
        self.state = state
        self.zip = zip

        global personIndex
        personIndex += 1
        print("personIndex is: ", personIndex)

        # Call the function to process the data
        self.process(personIndex)

        # Reset personIndex for the next sheet of labels
        if personIndex >= 20:
            personIndex = 0

    def process(self, personIndex):
        #  Write the context
        context = {
            "name" + str(personIndex): self.name,
            "address" + str(personIndex): self.address,
            "city" + str(personIndex): self.city,
            "state" + str(personIndex): self.state,
            "zip" + str(personIndex): self.zip
        }
        # Push the context to the main_context
        main_context.update(context)

# Save the sheet
    def saveFile(self, outputFolder):
        global counter
        counter = counter + 1
        print("Counter is:", counter)

        doc.render(main_context)
        fileName = "sheet" + str(counter) + ".docx"
        fileSaveLocation = outputFolder + "/" + fileName
        print(fileSaveLocation)
        doc.save(fileSaveLocation)
        print("CONSOLE IS:", self.ids.outputConsole.text)
        # self.ids.outputConsole.text = "Sheet Saved" + "\n"
        print("SHEET SAVED")
        main_context.clear()
        global_.message = "Sheet" + str(counter) + " Saved"
        print("GLOBAL MESSAGE IS: ", global_.message)
        # threading.Thread(target=MainGrid.updateConsole(self)).start()


@ thread
def runProcess(self, excelFile, outputFolder):

    # get dataframe from excel file
    # FILE NOT FOUND ERROR
    excelFile = excelFile[0]
    outputFolder = outputFolder[0]

    print("FOLDER IS: ", outputFolder)
    df = pd.read_excel(excelFile, usecols=[
        "NAME", "ADDRESS", "CITY", "STATE", "ZIP"])

    # Get length of dataframe
    datasetLen = len(df.index)

    #  Loop through dataframe
    for index, row in df.iterrows():
        # print("Index is: ", index)

        # Check when hit every 20 in the dataset
        if index + 1 in range(0, datasetLen, 20):
            print("In Range: ", index + 1)

            # # Send data to the class
            person = Person(row["NAME"], row["ADDRESS"],
                            row["CITY"], row["STATE"], row["ZIP"])
            doc.render(main_context)
            Person.saveFile(self, outputFolder)
            #  Print the done message
            print("DONE WITH PAGE")
            time.sleep(5)
            continue

    # Send data to the class
        person = Person(row["NAME"], row["ADDRESS"],
                        row["CITY"], row["STATE"], row["ZIP"])

        # Check if we are at the end of the dataset
        if index + 1 >= datasetLen:
            # Send data to the class
            person = Person(row["NAME"], row["ADDRESS"],
                            row["CITY"], row["STATE"], row["ZIP"])
            doc.render(main_context)
            Person.saveFile(self, outputFolder)
    #  Print the done message
            print("DONE")
            global_.message = "Processing finished."
            MainGrid.processing = False
            MainGrid.stop_console_clock(self)
            break


class LabelMaker(App):
    def build(self):

        return MainGrid()


# processDataThread = threading.Thread(target=runProcess)
# updateConsoleThread = threading.Thread(target=MainGrid.updateConsole)
if __name__ == '__main__':
    LabelMaker().run()

    # processDataThread.start()
    # threading.Thread(target=LabelMaker.run())
