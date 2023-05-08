import pandas as pd
import global_
from docxtpl import DocxTemplate

# TODO: There is a problem opening the excel sheet from outside of the exe dir. Find a fix

# Open the template file
doc = DocxTemplate("template2.docx")
# Save data


def saveFile():
    # +1 the global counter
    global_.counter = global_.counter + 1

    # put data from main_context into template
    doc.render(global_.main_context)

    # Sset the file name
    fileName = "sheet" + str(global_.counter) + ".docx"

    # Set the save location
    fileSaveLocation = global_.save_location + "/" + fileName

    # Save the file
    doc.save(fileSaveLocation)
    # print(fileSaveLocation)

    # global_.message = "Sheet" + str(global_.counter) + " Saved"
    global_.main_context.clear()
    print("SAVED")

    # Send message to console
    global_.message = "Sheet" + str(global_.counter) + " Saved"


def pushContext(row):
    #  Write the context
    context = {
        "name" + str(global_.personIndex): row["NAME"],
        "address" + str(global_.personIndex): row["ADDRESS"],
        "city" + str(global_.personIndex): row["CITY"] + ",",
        "state" + str(global_.personIndex): row["STATE"],
        "zip" + str(global_.personIndex): row["ZIP"],
    }

    # Push the context to the main_context
    global_.main_context.update(context)
    # print("GOBAL CONTEXT IS: ", global_.main_context)
    print(global_.personIndex)
    global_.personIndex = global_.personIndex + 1


def runProcess():
    ### VARS ###

    # Open excel file
    excelFile = global_.list_location
    # print(global_.list_location)

    # Set the dataframe
    df = pd.read_excel(excelFile, usecols=["NAME", "ADDRESS", "CITY", "STATE", "ZIP"])

    # Get the length of the dataset
    datasetLen = len(df.index)

    # loop through it. It starts at index 0 but row 2 of the sheet by default
    for index, row in df.iterrows():
        pushContext(row)

        """
        When you hit 20, save the context.
        These numbers are important! You need -1 as the start because the index starts at 0. 
        The second is the dataset length.
        Last is the step. This needs to be 20 because of the -1. 
        """
        if index in range(-1, datasetLen, 20):
            print("20 HIT. Name is: ", row["NAME"])
            pushContext(row)
            saveFile()
            # Reset the counter
            global_.personIndex = 0

        # TODO: Add another if statement for when we reach the end of the dataframe
        if index + 1 >= datasetLen:
            print("END OF DATAFRAME")
            # pushContext(row)
            saveFile()
            global_.message = "No more data"
