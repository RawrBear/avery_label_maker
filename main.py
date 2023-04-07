# import pandas as pd
# from processData import *


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
