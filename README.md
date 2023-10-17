# SelectFromDoc
This program lets you select data from documents (Excel, CSV, JSON, Text, XML) or clipboard using SQL language, with the possibility to visualize data in charts.


## Cloning the repository
- Clone the repository using the command below :
```
git clone https://github.com/smatimed/SelectFromDoc.git
```

- Move into the directory where we have the project files :
```
cd SelectFromDoc
```

- Create a virtual environment :
Install pipenv first:
```
pip install pipenv
```
Create the virtual environment:
```
set pipenv_venv_in_project=1
pipenv install
```
The requirements, stored in the file 'Pipfile', (pandas, pandasql, openpyxl, lxml, matplotlib) will be automatically installed.

- Activate the virtual environment :
```
.venv\scripts\activate
```


## Running the application
- To run the application, use :
```
python SelectFromDoc.py
```


## Using the application
- select a document (Excel, CSV, JSON, Fixed-Width Text or XML) by clicking on the button "..."
- or copy data to clipboard, in this case click on the button "From clipboard" (after having copied data)
- write a SELECT request ('from doc' is mandatory)
- click on "Execute" button
- to save the result, choose a format (CSV, Excel, JSON, Html, Text or XML) and click on "Export" button
- to visualize a chart from data click on "chart toolbar" to display the related toolbar
  then give these informations:
  - Type of the chart (Area, Bar, Barh, Line, Pie, Scatter)
  - x-axis column: number of the column to use in X axis
  - y-axis column(s): number(s) of the column(s) to use in Y axis, separated by ',' if there is more than one
  - Title
  - x-label: Label for X axis (empty = no label)
  - y-label: Label for Y axis (empty = no label)
  - Legend: Legend for column(s) used for Y axis, separated by ',' if there is more than one (empty = no legend)
  then click on the button "Visualization"
- to get help about SQlite select click on the button "SQL help".


## Screenshots
### main window
![mainwindow](https://github.com/smatimed/SelectFromDoc/blob/main/screenshots/main-screen.png?raw=true)
