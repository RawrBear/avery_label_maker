
from docxtpl import DocxTemplate

from plyer import filechooser
import pandas as pd
import time as time
import threading
import global_


# TODO: Remove all KIVY crap and replace with CTkinter
# TODO: Fix styling
# TODO: Split the code
# TODO: Improve and refactor this crappy code


def thread(function):
    def wrap(*args, **kwargs):
        t = threading.Thread(target=function, args=args,
                             kwargs=kwargs, daemon=True)
        t.start()

        return t
    return wrap


class MainGrid():

    listPath = ""
    outputPath = ""
    processing = False

    def updateConsole(self, *kwargs):
        self.ids.outputConsole.text = global_.message

        print("UPDATECONSOLE SAYS: ", global_.message)

    def btnOpenFile(self):
        self.listPath = filechooser.open_file(title="Choose your list..", filters=[
            ("Excel", "*.xlsx")])
        try:
            self.ids.dataList.text = self.listPath[0]

        except:

            self.ids.outputConsole.text = "Please Choose A File..."
            print("ERROR: CLICKED OUT OF THE SELECTED FILE")

    def btnSaveFile(self):
        self.outputPath = filechooser.choose_dir(title="Where do you want to save the output?", filters=[
            ("All Files", "*.*")])

        try:
            self.ids.outputDir.text = self.outputPath[0]
        except:
            self.ids.outputConsole.text = "Please Choose A Directory..."
            print(self.outputPath)
            print("ERROR: CLICKED OUT OF THE SELECTED DIRECTORY")

    def start_console_thread(self):
        print("Start console thread")
        threading.Thread(target=self.console_thread).start()

    def console_thread(self):
        print("Console thread running")
        self.console_clock = Clock.schedule_interval(self.updateConsole, 0.4)

    def stop_console_clock(self):
        self.console_clock.cancel()

    def btnGo(self):
        # if listpath and outputpath are not empty
        if self.listPath != "" and self.outputPath != "":
            runProcess(self, self.listPath, self.outputPath)
            self.processing == True
            global_.message = "Processing..."
        else:
            self.ids.outputConsole.text = "Please select a list and output directory"


# Opens the template file
doc = DocxTemplate("template.docx")

# Sets the main context to be an empty dictionary
main_context = {}


# Used to store the index of the current row
# This is used in process method
personIndex = 0


class Person:
    # global_.global_.counter = 0
    # Is used to track the sheets.
    # It logs every time the save function is called and sets that as the sheet number
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
        global_.counter
        global_.counter = global_.counter + 1
        print("global_.counter is:", global_.counter)

        doc.render(main_context)
        fileName = "sheet" + str(global_.counter) + ".docx"
        fileSaveLocation = outputFolder + "/" + fileName
        print(fileSaveLocation)
        doc.save(fileSaveLocation)
        print("SHEET SAVED")
        main_context.clear()
        global_.message = "Sheet" + str(global_.counter) + " Saved"


@ thread
def runProcess(self, excelFile, outputFolder):

    # get dataframe from excel file
    try:
        excelFile = excelFile[0]
    except IndexError:
        print("No excel file selected")

    try:
        outputFolder = outputFolder[0]
    except IndexError:
        print("No output folder selected")

    # Handle pandas error if the excel file is empty or has the wrong column names
    try:
        try:
            df = pd.read_excel(excelFile, usecols=[
                "NAME", "ADDRESS", "CITY", "STATE", "ZIP"])
            # df.columns returns a list of column names
            # loops thrrough each and check for True. 'name' in df

            # Get length of dataframe
            datasetLen = len(df.index)
        except:
            print("ERROR: Column names do not match")

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
                time.sleep(3)
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
                global_.counter = 0
                time.sleep(5)
                MainGrid.stop_console_clock(self)
                break
    except:
        print("ERROR")


class LabelMaker(App):
    def build(self):
        self.icon = 'label.png'
        self.title = 'Awesome Label Processor!'

        return MainGrid()


# processDataThread = threading.Thread(target=runProcess)
# updateConsoleThread = threading.Thread(target=MainGrid.updateConsole)
if __name__ == '__main__':
    LabelMaker().run()

    # processDataThread.start()
    # threading.Thread(target=LabelMaker.run())
