# from docxtpl import DocxTemplate
# import time as time
# import global_
# import pandas as pd
# from gui import updateConsole


# # Opens the template file
# doc = DocxTemplate("template.docx")

# # Sets the main context to be an empty dictionary
# main_context = {}

# # Is used to track the sheets.
# # It logs every time the save function is called and sets that as the sheet number
# counter = 0

# # Used to store the index of the current row
# # This is used in process method
# personIndex = 0


# class Person:
#     def __init__(self, name, address, city, state, zip):
#         self.name = name
#         self.address = address
#         self.city = city
#         self.state = state
#         self.zip = zip

#         global personIndex
#         personIndex += 1
#         print("personIndex is: ", personIndex)

#         # Call the function to process the data
#         self.process(personIndex)

#         # Reset personIndex for the next sheet of labels
#         if personIndex >= 20:
#             personIndex = 0

#     def process(self, personIndex):
#         #  Write the context
#         context = {
#             "name" + str(personIndex): self.name,
#             "address" + str(personIndex): self.address,
#             "city" + str(personIndex): self.city,
#             "state" + str(personIndex): self.state,
#             "zip" + str(personIndex): self.zip
#         }
#         # Push the context to the main_context
#         main_context.update(context)

# # Save the sheet
#     def saveFile(self, outputFolder):
#         global counter
#         counter = counter + 1
#         print("Counter is:", counter)

#         doc.render(main_context)
#         fileName = "sheet" + str(counter) + ".docx"
#         fileSaveLocation = outputFolder + "/" + fileName
#         print(fileSaveLocation)
#         doc.save(fileSaveLocation)
#         print("CONSOLE IS:", self.ids.outputConsole.text)
#         # self.ids.outputConsole.text = "Sheet Saved" + "\n"
#         print("SHEET SAVED")
#         main_context.clear()
#         global_.message = "Sheet" + str(counter) + " Saved"
#         print("GLOBAL MESSAGE IS: ", global_.message)
#         updateConsole()


# def runProcess(self, excelFile, outputFolder):

#     # get dataframe from excel file
#     # FILE NOT FOUND ERROR
#     excelFile = excelFile[0]
#     outputFolder = outputFolder[0]

#     print("FOLDER IS: ", outputFolder)
#     df = pd.read_excel(excelFile, usecols=[
#         "NAME", "ADDRESS", "CITY", "STATE", "ZIP"])

#     # Get length of dataframe
#     datasetLen = len(df.index)

#     #  Loop through dataframe
#     for index, row in df.iterrows():
#         # print("Index is: ", index)

#         # Check when hit every 20 in the dataset
#         if index + 1 in range(0, datasetLen, 20):
#             print("In Range: ", index + 1)

#             # # Send data to the class
#             person = Person(row["NAME"], row["ADDRESS"],
#                             row["CITY"], row["STATE"], row["ZIP"])
#             doc.render(main_context)
#             Person.saveFile(self, outputFolder)
#             #  Print the done message
#             print("DONE WITH PAGE")
#             time.sleep(5)
#             continue

#     # Send data to the class
#         person = Person(row["NAME"], row["ADDRESS"],
#                         row["CITY"], row["STATE"], row["ZIP"])

#         # Check if we are at the end of the dataset
#         if index + 1 >= datasetLen:
#             # Send data to the class
#             person = Person(row["NAME"], row["ADDRESS"],
#                             row["CITY"], row["STATE"], row["ZIP"])
#             doc.render(main_context)
#             Person.saveFile(self, outputFolder)
#     #  Print the done message
#             print("DONE")
#             break
