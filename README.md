## **DESCRIPTION**

This project is a small python program that takes data from an Excel file, and populates a Word document with the Avery label template.
It uses Jinja2 to generate the template and Excel as the dataset.

**NOTES:**
Please note that the formatting of some of the labls can be off if the address is long. There's no way to deal with this from the program as Python doesn't "see" the label limits. Please check the sheets after generating and fix any formatting errors manually. 

## **HOW TO USE**

### If running the exe


### If running from the command line:
1. Clone the repository
2. Edit the Excel file to include your data. Alternatively, you can use your file but the column names must be: NAME ADDRESS CITY STATE ZIP. Other columns will be ignored and won't cause problems if left in. (needs to be verified)
3. Open the cloned directory in your command line of choice.
4. Create your virtual env if you're using one.
5. Run "pip install" to install dependencies.
6. Run "labelmaker.py" to open the GUI.
7. You will now see "sheet1", "sheet2" etc in the directory. 

The program is set to create a new file for every 20 people in you dataset. This is for easier printing as Word sometimes screws up the formatting if you add more pages to the template. 


**TODO**
* Create cross-platform executables 
* Fix problem with not being able to process another sheet without restarting the app