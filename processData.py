from docxtpl import DocxTemplate


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

        if personIndex == 21:
            # Reset personIndex for the next sheet of labels
            personIndex = 0

# Call the function to process the data
        self.process(personIndex)

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

# Save the file
    def saveFile():
        global counter
        counter = counter + 1
        print("Counter is:", counter)

        doc.render(main_context)
        doc.save("sheet"+str(counter)+".docx")

        print("SHEET SAVED")
        main_context.clear()
