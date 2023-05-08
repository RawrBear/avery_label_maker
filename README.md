**DESCRIPTION**

This project is a small python program that takes data from an Excel file, and populates a Word document with the Avery label template.
It uses Jinja2 to generate the template and Excel as the dataset.


**HOW TO USE**

If running from the command line:
1. Clone the repository
2. Edit the Excel file to include your data. Alternatively, you can use your file but the column names must be: NAME ADDRESS CITY STATE ZIP. Other columns will be ignored and wont cause problems if left in.
3. Open the cloned directory in your command line of choice.
4. Create your virtual env if you;re using one.
5. Run "pip install" to install dependencies.
6. Open the "labelmaker.py" file and edit the variables for DATASET and TEMPLATE to reflect the name and location of your Excel and Word templates.
7. Run "labelmaker.py" to generate the template.
8. You will now see "sheet1", "sheet2" etc in the directory. 

The program is set to create a new file for every 20 people in you dataset. This is for easier printing as Word sometimes screws up the formatting if you add more pages to the template. 

**TODO**
* Create cross-platform executables 