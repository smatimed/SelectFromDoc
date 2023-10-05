# SelectFromDoc
This program lets you select data from documents (Excel, csv, txt) or clipboard using SQL language.


# Cloning the repository
--> Clone the repository using the command below :
git clone https://github.com/smatimed/SelectFromDoc.git

--> Move into the directory where we have the project files :
cd SelectFromDoc

--> Create a virtual environment :

Install pipenv first:
pip install pipenv

Create the virtual environment:
set pipenv_venv_in_project=1
pipenv install
The requirements will be automatically installed

--> Activate the virtual environment :
.venv\scripts\activate


# Running the application
--> To run the application, use :
python SelectFromDoc.py


# Using the application
- select a document (Excel, csv or fixed-width text) by clicking on the button "..."
- or copy data to clipboard, in this case click on the button "From clipboard" (after having copied data)
- write a SELECT request ('from doc' is mandatory)
- click on "Execute" button
- to save the result, choose a format (Csv, Excel, Html or Text) and click on "Export" button
- to get help about SQlite select click on the button "SQL help".
