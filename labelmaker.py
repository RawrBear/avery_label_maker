from docxtpl import DocxTemplate
import pandas as pd

# Opens the template file
doc = DocxTemplate("template.docx")

# Sets the main context to be an empty dictionary
main_context = {}

# Person class


class Person:
    def __init__(self, name, address, city, state, zip, index):
        self.name = name
        self.address = address
        self.city = city
        self.state = state
        self.zip = zip

# Call the function to process the data
        self.process(index)

    def process(self, index):
        #  Write the context
        context = {
            "name" + str(index): self.name,
            "address" + str(index): self.address,
            "city" + str(index): self.city,
            "state" + str(index): self.state,
            "zip" + str(index): self.zip
        }
        # Push the context to the main_context
        main_context.update(context)

# Save the file
    def saveFile():
        counter = 0
        counter = counter + 1
        doc.save("sheet"+str(counter)+".docx")


# get dataframe from excel file
df = pd.read_excel("birthday_list.xlsx", usecols=[
                   "NAME", "ADDRESS", "CITY", "STATE", "ZIP"])

# Get length of dataframe
datasetLen = len(df.index)

#  Loop through dataframe
for index, row in df.iterrows():
    print("Index is: ", index)

# Send data to the class
    person = Person(row["NAME"], row["ADDRESS"],
                    row["CITY"], row["STATE"], row["ZIP"], index)

    # Check if we are at the end of the dataset
    if index + 1 >= datasetLen:
        doc.render(main_context)
        Person.saveFile()
#  Print the done message
        print("DONE")
        break

# TODO Test with bigger dataset
# TODO Change dataset to fake data before making git repo
# TODO Make git repo
