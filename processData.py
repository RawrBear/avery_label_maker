from docxtpl import DocxTemplate
import global_

import time as time

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

        if personIndex >= 20:
            # Reset personIndex for the next sheet of labels
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
        print("SELF IS:", self)
        global_.message = fileName + " has been saved. \n\n"
        print("GLOBAL IS: ", global_.message)
        print("SHEET SAVED")
        main_context.clear()
