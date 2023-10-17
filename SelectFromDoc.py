from xmlrpc.client import boolean
import pandas as pd
from pandas import read_fwf, read_json, read_xml, read_clipboard
from pandasql import sqldf
import openpyxl

import matplotlib.pyplot as plt
import numpy as np

import tkinter as tk
from tkinter import *
from tkinter import messagebox
from tkinter import ttk   # To override the basic Tk widgets, the import of "from tkinter import ttk" should follow the Tk import "from tkinter import *"
# from PIL import Image   # PIL for image processing
# import glob   # glob for iterating through files of the given folder in the OS
from tkinter import filedialog
import ntpath

from time import time
from datetime import datetime

from os.path import exists

from pathlib import Path

root = Tk()
root_width = 980
root_height = 600
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
root.geometry(f'{root_width}x{root_height}+{(screen_width-root_width) // 2}+{(screen_height-root_height) // 2}')

root.title('Select From Document (Excel, Csv, Json, Text, Xml) or Clipboard')
# root.iconbitmap(path.abspath(path.join(path.dirname(__file__), 'SelectFromDoc.ico')))
root.iconbitmap('SelectFromDoc.ico')


class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        self.tooltip_window = tk.Toplevel(self.widget)
        x, y, _, _ = self.widget.bbox("insert")
        x = self.widget.winfo_rootx() + x
        y = self.widget.winfo_rooty() + y + 25
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip_window, text=self.text, justify='left', background="#f8f8b1", relief='solid', borderwidth=1)   # old: ffffe0
        label.pack(ipadx=1)

    def hide_tooltip(self, event):
        if hasattr(self, 'tooltip_window'):
            self.tooltip_window.destroy()


def on_resize(event):
    # Update the size of 'frame_sql_resultat' when the window is resized
    numRow_Frame_Sql_Resultat = 4
    root.grid_rowconfigure(numRow_Frame_Sql_Resultat, weight=1)
    root.grid_columnconfigure(0, weight=1)


def on_keypress(event):
    # F8
    if event.keysym =="F8":
        if boutonExecuter["state"] != "disabled":
            Executer()
    # F1
    elif event.keysym =="F1":
        ouvrir_Sql_Aide()
    # Ctrl + Q
    elif event.keysym == 'q' and event.state & 0x4:
        root.quit()
    # else:
    #     print(f'Key pressed: {event.keysym}')

    # if event.keysym == 's' and event.state & 0x4:
    #     print("CTRL+S pressed")
    # if event.keysym == 's' and event.state & 0x8:
    #     print("Alt+S pressed")    
    # if event.keysym == 'S' and event.state & 0x1:
    #     print("Shift+S pressed")

root.bind("<KeyPress>",on_keypress)



def ouvrir_Sql_Aide():

    def on_resize(event):
        Sql_Aide.grid_rowconfigure(0, weight=1)
        Sql_Aide.grid_columnconfigure(0, weight=1)

    def changerWordWrap():
        if wordWrap.get() == 1:
            text_aide['wrap'] = WORD
        else:
            text_aide['wrap'] = 'none'
    
    Sql_Aide = tk.Toplevel()
    Sql_Aide.iconbitmap('SelectFromDoc.ico')
    Sql_Aide.title("SQL help")
    win_width = 550
    win_height = 370
    Sql_Aide.config(width=win_width, height=win_height)
    Sql_Aide.geometry(f'{win_width}x{win_height}+{(screen_width-win_width) // 2}+{(screen_height-win_height) // 2}')

    frame_aide = ttk.Frame(Sql_Aide)
    frame_aide.grid(row=0, column=0, sticky='WENS')

    Sql_Aide.grid_columnconfigure(0, weight=1)   # makes the widgets with sticky='WE' fill all the width

    scrollbar_x = ttk.Scrollbar(frame_aide, orient=HORIZONTAL)
    scrollbar_x.pack(side=BOTTOM, fill=X)
    scrollbar_y = ttk.Scrollbar(frame_aide)
    scrollbar_y.pack(side=RIGHT, fill=Y)

    text_aide = Text(frame_aide, xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set, wrap='none', height=20)# , width=30)
    text_aide.pack(expand=True, fill=BOTH)

    scrollbar_x.config(command=text_aide.xview)
    scrollbar_y.config(command=text_aide.yview)


    frame_enbas = ttk.Frame(Sql_Aide)
    frame_enbas.grid(row=1, column=0, sticky='WE')
    ttk.Style().configure("Custom.TFrame", background=bg_color_default)
    frame_enbas.configure(style="Custom.TFrame")

    wordWrap = IntVar()
    option_wordwrap = ttk.Checkbutton(frame_enbas, text="Word wrap", variable=wordWrap, onvalue=1, offvalue=0, width=10, command=changerWordWrap)
    ttk.Style().configure("Custom.TCheckbutton", foreground=fg_color_default_Label, background=bg_color_default)
    option_wordwrap.configure(style="Custom.TCheckbutton")
    option_wordwrap.pack(side=LEFT)

    button_close = ttk.Button(frame_enbas, text="Close", command=Sql_Aide.destroy)
    button_close.pack(side=RIGHT, padx=5, pady=5)

    # wordWrap = IntVar()
    # ttk.Checkbutton(Sql_Aide, text="Word wrap", variable=wordWrap, onvalue=1, offvalue=0, width=10, command=changerWordWrap).grid(row=1, sticky=W, padx=5)

    # button_close = ttk.Button(Sql_Aide, text="Close", command=Sql_Aide.destroy)
    # button_close.grid(row=1, sticky=E, padx=5, pady=5)

    # Bind the window resize event to the on_resize function
    Sql_Aide.bind('<Configure>', on_resize)


    # --- Texte de l'aide
    text_aide.insert(END, "Source (mainly): sqlite.org")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\nContent:")
    text_aide.insert(END, "\n-------")
    text_aide.insert(END, "\n  1. SELECT simple syntax")
    text_aide.insert(END, "\n  2. Built-In Scalar SQL Functions")
    text_aide.insert(END, "\n  3. List of built-in aggregate functions")
    text_aide.insert(END, "\n  4. Date And Time Functions")
    text_aide.insert(END, "\n  5. SQLite Keywords")
    text_aide.insert(END, "\n  ")
    text_aide.insert(END, "\n  ")
    text_aide.insert(END, "\n1. SELECT simple syntax")
    text_aide.insert(END, "\n   --------------------")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n  SELECT [ALL | DISTINCT] column_list ")
    text_aide.insert(END, "\n  FROM table_list")
    text_aide.insert(END, "\n    JOIN table ON join_condition")
    text_aide.insert(END, "\n  WHERE row_filter")
    text_aide.insert(END, "\n  ORDER BY column [ASC | DESC]")
    text_aide.insert(END, "\n  LIMIT count OFFSET offset")
    text_aide.insert(END, "\n  GROUP BY column")
    text_aide.insert(END, "\n  HAVING group_filter;")
    text_aide.insert(END, "\n  ")
    text_aide.insert(END, "\n  Logical operators: AND, BETWEEN, EXISTS, IN, NOT, LIKE, GLOB, NOT, OR, IS NULL, IS, IS NOT, ||, UNIQUE")
    text_aide.insert(END, "\n  ")
    text_aide.insert(END, "\n  Connections: INNER JOIN, LEFT [OUTER] JOIN, RIGHT [OUTER] JOIN, FULL [OUTER] JOIN, CROSS JOIN")
    text_aide.insert(END, "\n  ")
    text_aide.insert(END, "\n  Compound-operator: UNION, UNION ALL, INTERSECT, EXCEPT")
    text_aide.insert(END, "\n  ")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n2. Built-In Scalar SQL Functions")
    text_aide.insert(END, "\n   -----------------------------")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- abs(X)")
    text_aide.insert(END, "\n   The abs(X) function returns the absolute value of the numeric argument X. Abs(X) returns NULL if X is NULL. Abs(X) returns 0.0 if X is a string or blob that cannot be converted to a numeric value. If X is the integer ­9223372036854775808 then abs(X) throws an integer overflow error since there is no equivalent positive 64­bit two complement value.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- changes()")
    text_aide.insert(END, "\n  The changes() function returns the number of database rows that were changed or inserted or deleted by the most recently completed INSERT, DELETE, or UPDATE statement, exclusive of statements in lower­level triggers. The changes() SQL function is a wrapper around the sqlite3_changes64() C/C++ function and hence follows the same rules for counting changes.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- char(X1,X2,...,XN)")
    text_aide.insert(END, "\n  The char(X1,X2,...,XN) function returns a string composed of characters having the unicode code point values of integers X1 through XN, respectively.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- coalesce(X,Y,...)")
    text_aide.insert(END, "\n  The coalesce() function returns a copy of its first non­NULL argument, or NULL if all arguments are NULL. Coalesce() must have at least 2 arguments.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- format(FORMAT,...)")
    text_aide.insert(END, "\n  The format(FORMAT,...) SQL function works like the sqlite3_mprintf() C­language function and the printf() function from the standard C library. The first argument is a format string that specifies how to construct the output string using values taken from subsequent arguments. If the FORMAT argument is missing or NULL then the result is NULL. The %n format is silently ignored and does not consume an argument. The %p format is an alias for %X. The %z format is interchangeable with %s. If there are too few arguments in the argument list, missing arguments are assumed to have a NULL value, which is translated into 0 or 0.0 for numeric formats or an empty string for %s. See the built­in printf() documentation for additional information.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- glob(X,Y)")
    text_aide.insert(END, "\n  The glob(X,Y) function is equivalent to the expression \"Y GLOB X\". Note that the X and Y arguments are reversed in the glob() function relative to the infix GLOB operator. Y is the string and X is the pattern. So, for example, the following expressions are equivalent:")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n  name GLOB '*helium*'")
    text_aide.insert(END, "\n  glob('*helium*',name)")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n  If the sqlite3_create_function() interface is used to override the glob(X,Y) function with an alternative implementation then the GLOB operator will invoke the alternative implementation.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- hex(X)")
    text_aide.insert(END, "\n  The hex() function interprets its argument as a BLOB and returns a string which is the upper­case hexadecimal rendering of the content of that blob.")
    text_aide.insert(END, "\n  If the argument X in \"hex(X)\" is an integer or floating point number, then \"interprets its argument as a BLOB\" means that the binary number is first converted into a UTF8 text representation, then that text is interpreted as a BLOB. Hence, \"hex(12345678)\" renders as \"3132333435363738\" not the binary representation of the integer value \"0000000000BC614E\".")
    text_aide.insert(END, "\n  See also: unhex()")
    text_aide.insert(END, "\n  ")
    text_aide.insert(END, "\n- ifnull(X,Y)")
    text_aide.insert(END, "\n  The ifnull() function returns a copy of its first non­NULL argument, or NULL if both arguments are NULL. Ifnull() must have exactly 2 arguments. The ifnull() function is equivalent to coalesce() with two arguments.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- iif(X,Y,Z)")
    text_aide.insert(END, "\n  The iif(X,Y,Z) function returns the value Y if X is true, and Z otherwise. The iif(X,Y,Z) function is logically equivalent to and generates the same bytecode as the CASE expression \"CASE WHEN X THEN Y ELSE Z END\".")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- instr(X,Y)")
    text_aide.insert(END, "\n  The instr(X,Y) function finds the first occurrence of string Y within string X and returns the number of prior characters plus 1, or 0 if Y is nowhere found within X. Or, if X and Y are both BLOBs, then instr(X,Y) returns one more than the number bytes prior to the first occurrence of Y, or 0 if Y does not occur anywhere within X. If both arguments X and Y to instr(X,Y) are non­NULL and are not BLOBs then both are interpreted as strings. If either X or Y are NULL in instr(X,Y) then the result is NULL.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- last_insert_rowid()")
    text_aide.insert(END, "\n   The last_insert_rowid() function returns the ROWID of the last row insert from the database connection which invoked the function. The last_insert_rowid() SQL function is a wrapper around the sqlite3_last_insert_rowid() C/C++ interface function.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- length(X)")
    text_aide.insert(END, "\n  For a string value X, the length(X) function returns the number of characters (not bytes) in X prior to the first NUL character. Since SQLite strings do not normally contain NUL characters, the length(X) function will usually return the total number of characters in the string X. For a blob value X, length(X) returns the number of bytes in the blob. If X is NULL then length(X) is NULL. If X is numeric then length(X) returns the length of a string representation of X.")
    text_aide.insert(END, "\n  Note that for strings, the length(X) function returns the character length of the string, not the byte length. The character length is the number of characters in the string. The character length is always different from the byte length for UTF­16 strings, and can be different from the byte length for UTF­8 strings if the string contains multi­byte characters. Use the octet_length() function to find the byte length of a string.")
    text_aide.insert(END, "\n  For BLOB values, length(X) always returns the byte­length of the BLOB.")
    text_aide.insert(END, "\n  For string values, length(X) must read the entire string into memory in order to compute the character length. But for BLOB values, that is not necessary as SQLite knows how many bytes are in the BLOB. Hence, for multi­megabyte values, the length(X) function is usually much faster for BLOBs than for strings, since it does not need to load the value into memory.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- like(X,Y)")
    text_aide.insert(END, "\n- like(X,Y,Z)")
    text_aide.insert(END, "\n  The like() function is used to implement the \"Y LIKE X [ESCAPE Z]\" expression. If the optional ESCAPE clause is present, then the like() function is invoked with three arguments. Otherwise, it is invoked with two arguments only. Note that the X and Y parameters are reversed in the like() function relative to the infix LIKE operator. X is the pattern and Y is the string to match against that pattern. Hence, the following expressions are equivalent:")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n  name LIKE '%neon%'")
    text_aide.insert(END, "\n  like('%neon%',name)")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n  The sqlite3_create_function() interface can be used to override the like() function and thereby change the operation of the LIKE operator. When overriding the like() function, it may be important to override both the two and three argument versions of the like() function. Otherwise, different code may be called to implement the LIKE operator depending on whether or not an ESCAPE clause was specified.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- likelihood(X,Y)")
    text_aide.insert(END, "\n  The likelihood(X,Y) function returns argument X unchanged. The value Y in likelihood(X,Y) must be a floating point constant between 0.0 and 1.0, inclusive. The likelihood(X) function is a no­op that the code generator optimizes away so that it consumes no CPU cycles during run­time (that is, during calls to sqlite3_step()). The purpose of the likelihood(X,Y) function is to provide a hint to the query planner that the argument X is a boolean that is true with a probability of approximately Y. The unlikely(X) function is short­hand for likelihood(X,0.0625). The likely(X) function is short­hand for likelihood(X,0.9375).")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- likely(X)")
    text_aide.insert(END, "\n  The likely(X) function returns the argument X unchanged. The likely(X) function is a no­op that the code generator optimizes away so that it consumes no CPU cycles at run­time (that is, during calls to sqlite3_step()). The purpose of the likely(X) function is to provide a hint to the query planner that the argument X is a boolean value that is usually true. The likely(X) function is equivalent to likelihood(X,0.9375). See also: unlikely(X).")
    text_aide.insert(END, "\n  ")
    text_aide.insert(END, "\n- load_extension(X)")
    text_aide.insert(END, "\n  load_extension(X,Y)")
    text_aide.insert(END, "\n  The load_extension(X,Y) function loads SQLite extensions out of the shared library file named X using the entry point Y. The result of load_extension() is always a NULL. If Y is omitted then the default entry point name is used. The load_extension() function raises an exception if the extension fails to load or initialize correctly.")
    text_aide.insert(END, "\n  The load_extension() function will fail if the extension attempts to modify or delete an SQL function or collating sequence. The extension can add new functions or collating sequences, but cannot modify or delete existing functions or collating sequences because those functions and/or collating sequences might be used elsewhere in the currently running SQL statement. To load an extension that changes or deletes functions or collating sequences, use the sqlite3_load_extension() C­language API.")
    text_aide.insert(END, "\n  For security reasons, extension loading is disabled by default and must be enabled by a prior call to sqlite3_enable_load_extension().")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- lower(X)")
    text_aide.insert(END, "\n  The lower(X) function returns a copy of string X with all ASCII characters converted to lower case. The default built­in lower() function works for ASCII characters only. To do case conversions on non­ASCII characters, load the ICU extension.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- ltrim(X)")
    text_aide.insert(END, "\n  ltrim(X,Y)")
    text_aide.insert(END, "\n  The ltrim(X,Y) function returns a string formed by removing any and all characters that appear in Y from the left side of X. If the Y argument is omitted, ltrim(X) removes spaces from the left side of X.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- max(X,Y,...)")
    text_aide.insert(END, "\n  The multi­argument max() function returns the argument with the maximum value, or return NULL if any argument is NULL. The multi­argument max() function searches its arguments from left to right for an argument that defines a collating function and uses that collating function for all string comparisons. If none of the arguments to max() define a collating function, then the BINARY collating function is used. Note that max() is a simple function when it has 2 or more arguments but operates as an aggregate function if given only a single argument.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- min(X,Y,...)")
    text_aide.insert(END, "\n  The multi­argument min() function returns the argument with the minimum value. The multi­argument min() function searches its arguments from left to right for an argument that defines a collating function and uses that collating function for all string comparisons. If none of the arguments to min() define a collating function, then the BINARY collating function is used. Note that min() is a simple function when it has 2 or more arguments but operates as an aggregate function if given only a single argument.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- nullif(X,Y)")
    text_aide.insert(END, "\n  The nullif(X,Y) function returns its first argument if the arguments are different and NULL if the arguments are the same. The nullif(X,Y) function searches its arguments from left to right for an argument that defines a collating function and uses that collating function for all string comparisons. If neither argument to nullif() defines a collating function then the BINARY collating function is used.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- octet_length(X)")
    text_aide.insert(END, "\n  The octet_length(X) function returns the number of bytes in the encoding of text string X. If X is NULL then octet_length(X) returns NULL. If X is a BLOB value, then octet_length(X) is the same as length(X). If X is a numeric value, then octet_length(X) returns the number of bytes in a text rendering of that number.")
    text_aide.insert(END, "\n  Because octet_length(X) returns the number of bytes in X, not the number of characters, the value returned depends on the database encoding. The octet_length() function can return different answers for the same input string if the database encoding is UTF16 instead of UTF8.")
    text_aide.insert(END, "\n  If argument X is a table column and the value is of type text or blob, then octet_length(X) avoids reading the content of X from disk, as the byte length can be computed from metadata. Thus, octet_length(X) is efficient even if X is column containing a multi­megabyte text or blob value.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- printf(FORMAT,...)")
    text_aide.insert(END, "\n  The printf() SQL function is an alias for the format() SQL function. The format() SQL function was original named printf(). But the name was later changed to format() for compatibility with other database engines. The original printf() name is retained as an alias so as not to break any legacy code.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- quote(X)")
    text_aide.insert(END, "\n  The quote(X) function returns the text of an SQL literal which is the value of its argument suitable for inclusion into an SQL statement. Strings are surrounded by single­quotes with escapes on interior quotes as needed. BLOBs are encoded as hexadecimal literals. Strings with embedded NUL characters cannot be represented as string literals in SQL and hence the returned string literal is truncated prior to the first NUL.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- random()")
    text_aide.insert(END, "\n  The random() function returns a pseudo­random integer between ­9223372036854775808 and +9223372036854775807.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- randomblob(N)")
    text_aide.insert(END, "\n  The randomblob(N) function return an N­byte blob containing pseudo­random bytes. If N is less than 1 then a 1­byte random blob is returned.")
    text_aide.insert(END, "\n  Hint: applications can generate globally unique identifiers using this function together with hex() and/or lower() like this:")
    text_aide.insert(END, "\n    hex(randomblob(16))")
    text_aide.insert(END, "\n	lower(hex(randomblob(16)))")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- replace(X,Y,Z)")
    text_aide.insert(END, "\n  The replace(X,Y,Z) function returns a string formed by substituting string Z for every occurrence of string Y in string X. The BINARY collating sequence is used for comparisons. If Y is an empty string then return X unchanged. If Z is not initially a string, it is cast to a UTF­8 string prior to processing.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- round(X)")
    text_aide.insert(END, "\n  round(X,Y)")
    text_aide.insert(END, "\n  The round(X,Y) function returns a floating­point value X rounded to Y digits to the right of the decimal point. If the Y argument is omitted or negative, it is taken to be 0.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- rtrim(X)")
    text_aide.insert(END, "\n  rtrim(X,Y)")
    text_aide.insert(END, "\n  The rtrim(X,Y) function returns a string formed by removing any and all characters that appear in Y from the right side of X. If the Y argument is omitted, rtrim(X) removes spaces from the right side of X.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- sign(X)")
    text_aide.insert(END, "\n  The sign(X) function returns ­1, 0, or +1 if the argument X is a numeric value that is negative, zero, or positive, respectively. If the argument to sign(X) is NULL or is a string or blob that cannot be losslessly converted into a number, then sign(X) returns NULL.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- soundex(X)")
    text_aide.insert(END, "\n  The soundex(X) function returns a string that is the soundex encoding of the string X. The string \"?000\" is returned if the argument is NULL or contains no ASCII alphabetic characters. This function is omitted from SQLite by default. It is only available if the SQLITE_SOUNDEX compile-time option is used when SQLite is built.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- sqlite_compileoption_get(N)")
    text_aide.insert(END, "\n  The sqlite_compileoption_get() SQL function is a wrapper around the sqlite3_compileoption_get() C/C++ function. This routine returns the N-th compile-time option used to build SQLite or NULL if N is out of range. See also the compile_options pragma.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- sqlite_compileoption_used(X)")
    text_aide.insert(END, "\n  The sqlite_compileoption_used() SQL function is a wrapper around the sqlite3_compileoption_used() C/C++ function. When the argument X to sqlite_compileoption_used(X) is a string which is the name of a compile-time option, this routine returns true (1) or false (0) depending on whether or not that option was used during the build.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- sqlite_offset(X)")
    text_aide.insert(END, "\n  The sqlite_offset(X) function returns the byte offset in the database file for the beginning of the record from which value would be read. If X is not a column in an ordinary table, then sqlite_offset(X) returns NULL. The value returned by sqlite_offset(X) might reference either the original table or an index, depending on the query. If the value X would normally be extracted from an index, the sqlite_offset(X) returns the offset to the corresponding index record. If the value X would be extracted from the original table, then sqlite_offset(X) returns the offset to the table record.")
    text_aide.insert(END, "\n  The sqlite_offset(X) SQL function is only available if SQLite is built using the -DSQLITE_ENABLE_OFFSET_SQL_FUNC compile-time option.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- sqlite_source_id()")
    text_aide.insert(END, "\n  The sqlite_source_id() function returns a string that identifies the specific version of the source code that was used to build the SQLite library. The string returned by sqlite_source_id() is the date and time that the source code was checked in followed by the SHA3-256 hash for that check-in. This function is an SQL wrapper around the sqlite3_sourceid() C interface.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- sqlite_version()")
    text_aide.insert(END, "\n  The sqlite_version() function returns the version string for the SQLite library that is running. This function is an SQL wrapper around the sqlite3_libversion() C-interface.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- substr(X,Y,Z)")
    text_aide.insert(END, "\n  substr(X,Y)")
    text_aide.insert(END, "\n  substring(X,Y,Z)")
    text_aide.insert(END, "\n  substring(X,Y)")
    text_aide.insert(END, "\n  The substr(X,Y,Z) function returns a substring of input string X that begins with the Y-th character and which is Z characters long. If Z is omitted then substr(X,Y) returns all characters through the end of the string X beginning with the Y-th. The left-most character of X is number 1. If Y is negative then the first character of the substring is found by counting from the right rather than the left. If Z is negative then the abs(Z) characters preceding the Y-th character are returned. If X is a string then characters indices refer to actual UTF-8 characters. If X is a BLOB then the indices refer to bytes.")
    text_aide.insert(END, "\n  \"substring()\" is an alias for \"substr()\" beginning with SQLite version 3.34.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- total_changes()")
    text_aide.insert(END, "\n  The total_changes() function returns the number of row changes caused by INSERT, UPDATE or DELETE statements since the current database connection was opened. This function is a wrapper around the sqlite3_total_changes64() C/C++ interface.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- trim(X)")
    text_aide.insert(END, "\n  trim(X,Y)")
    text_aide.insert(END, "\n  The trim(X,Y) function returns a string formed by removing any and all characters that appear in Y from both ends of X. If the Y argument is omitted, trim(X) removes spaces from both ends of X.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- typeof(X)")
    text_aide.insert(END, "\n  The typeof(X) function returns a string that indicates the datatype of the expression X: \"null\", \"integer\", \"real\", \"text\", or \"blob\".")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- unhex(X)")
    text_aide.insert(END, "\n  unhex(X,Y)")
    text_aide.insert(END, "\n  The unhex(X,Y) function returns a BLOB value which is the decoding of the hexadecimal string X. If X contains any characters that are not hexadecimal digits and which are not in Y, then unhex(X,Y) returns NULL. If Y is omitted, it is understood to be an empty string and hence X must be a pure hexadecimal string. All hexadecimal digits in X must occur in pairs, with both digits of each pair beginning immediately adjacent to one another, or else unhex(X,Y) returns NULL. If either parameter X or Y is NULL, then unhex(X,Y) returns NULL. The X input may contain an arbitrary mix of upper and lower case hexadecimal digits. Hexadecimal digits in Y have no affect on the translation of X. Only characters in Y that are not hexadecimal digits are ignored in X.")
    text_aide.insert(END, "\n  See also: hex()")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- unicode(X)")
    text_aide.insert(END, "\n  The unicode(X) function returns the numeric unicode code point corresponding to the first character of the string X. If the argument to unicode(X) is not a string then the result is undefined.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- unlikely(X)")
    text_aide.insert(END, "\n  The unlikely(X) function returns the argument X unchanged. The unlikely(X) function is a no-op that the code generator optimizes away so that it consumes no CPU cycles at run-time (that is, during calls to sqlite3_step()). The purpose of the unlikely(X) function is to provide a hint to the query planner that the argument X is a boolean value that is usually not true. The unlikely(X) function is equivalent to likelihood(X, 0.0625).")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- upper(X)")
    text_aide.insert(END, "\n  The upper(X) function returns a copy of input string X in which all lower-case ASCII characters are converted to their upper-case equivalent.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- zeroblob(N)")
    text_aide.insert(END, "\n  The zeroblob(N) function returns a BLOB consisting of N bytes of 0x00. SQLite manages these zeroblobs very efficiently. Zeroblobs can be used to reserve space for a BLOB that is later written using incremental BLOB I/O. This SQL function is implemented using the sqlite3_result_zeroblob() routine from the C/C++ interface.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n3. List of built-in aggregate functions")
    text_aide.insert(END, "\n   ------------------------------------")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- avg(X)")
    text_aide.insert(END, "\n  The avg() function returns the average value of all non-NULL X within a group. String and BLOB values that do not look like numbers are interpreted as 0. The result of avg() is always a floating point value whenever there is at least one non-NULL input even if all inputs are integers. The result of avg() is NULL if there are no non-NULL inputs. The result of avg() is computed as total()/count() so all of the constraints that apply to total() also apply to avg().")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- count(X)")
    text_aide.insert(END, "\n  count(*)")
    text_aide.insert(END, "\n  The count(X) function returns a count of the number of times that X is not NULL in a group. The count(*) function (with no arguments) returns the total number of rows in the group.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- group_concat(X)")
    text_aide.insert(END, "\n  group_concat(X,Y)")
    text_aide.insert(END, "\n  The group_concat() function returns a string which is the concatenation of all non-NULL values of X. If parameter Y is present then it is used as the separator between instances of X. A comma (',') is used as the separator if Y is omitted. The order of the concatenated elements is arbitrary.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- max(X)")
    text_aide.insert(END, "\n  The max() aggregate function returns the maximum value of all values in the group. The maximum value is the value that would be returned last in an ORDER BY on the same column. Aggregate max() returns NULL if and only if there are no non-NULL values in the group.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- min(X)")
    text_aide.insert(END, "\n  The min() aggregate function returns the minimum non-NULL value of all values in the group. The minimum value is the first non-NULL value that would appear in an ORDER BY of the column. Aggregate min() returns NULL if and only if there are no non-NULL values in the group.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- sum(X)")
    text_aide.insert(END, "\n  total(X)")
    text_aide.insert(END, "\n  The sum() and total() aggregate functions return the sum of all non-NULL values in the group. If there are no non-NULL input rows then sum() returns NULL but total() returns 0.0. NULL is not normally a helpful result for the sum of no rows but the SQL standard requires it and most other SQL database engines implement sum() that way so SQLite does it in the same way in order to be compatible. The non-standard total() function is provided as a convenient way to work around this design problem in the SQL language.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n4. Date And Time Functions")
    text_aide.insert(END, "\n   -----------------------")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- date(timevalue, modifier, modifier, ...)")
    text_aide.insert(END, "\n  The date() function returns the date as text in this format: YYYYMMDD.")
    text_aide.insert(END, "\n  ")
    text_aide.insert(END, "\n- time(timevalue, modifier, modifier, ...)")
    text_aide.insert(END, "\n  The time() function returns the time as text in this format: HH:MM:SS.")
    text_aide.insert(END, "\n  ")
    text_aide.insert(END, "\n- datetime(timevalue, modifier, modifier, ...)")
    text_aide.insert(END, "\n  The datetime() function returns the date and time as text in this formats: YYYYMMDD HH:MM:SS.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- julianday(timevalue, modifier, modifier, ...)")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- unixepoch(timevalue, modifier, modifier, ...)")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n- strftime(format, timevalue, modifier, modifier, ...)")
    text_aide.insert(END, "\n  The strftime() routine returns the date formatted according to the format string specified as the first argument. The format string supports the most common substitutions found in the strftime() function from the standard C library plus two new substitutions, %f and %J. The following is a complete list of valid strftime() substitutions:")
    text_aide.insert(END, "\n    %d day of month: 00")
    text_aide.insert(END, "\n    %f fractional seconds: SS.SSS")
    text_aide.insert(END, "\n    %H hour: 0024")
    text_aide.insert(END, "\n    %j day of year: 001366")
    text_aide.insert(END, "\n    %J Julian day number (fractional)")
    text_aide.insert(END, "\n    %m month: 0112")
    text_aide.insert(END, "\n    %M minute: 0059")
    text_aide.insert(END, "\n    %s seconds since 19700101")
    text_aide.insert(END, "\n    %S seconds: 0059")
    text_aide.insert(END, "\n    %w day of week 06")
    text_aide.insert(END, "\n    with Sunday==0")
    text_aide.insert(END, "\n    %W week of year: 0053")
    text_aide.insert(END, "\n    %Y year: 00009999")
    text_aide.insert(END, "\n    %% %")
    text_aide.insert(END, "\n	")
    text_aide.insert(END, "\n- timediff(timevalue, timevalue)")
    text_aide.insert(END, "\n  The timediff(A,B) routine returns a string that describes the amount of time that must be added to B in order to reach time A. The format of the timediff() result is designed to be humanreadable. The format is: (+|)YYYYMMDD HH:MM:SS.SSS")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n  The first six date and time functions take an optional time value as an argument, followed by zero or more modifiers. The strftime() function also takes a format string as its first argument. The timediff() function takes exactly two arguments which are both time values.")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n")
    text_aide.insert(END, "\n5. SQLite Keywords")
    text_aide.insert(END, "\n   ---------------")
    text_aide.insert(END, "\n  ABORT, ACTION, ADD, AFTER, ALL, ALTER, ALWAYS, ANALYZE,")
    text_aide.insert(END, "\n  AND, AS, ASC, ATTACH, AUTOINCREMENT, BEFORE, BEGIN,")
    text_aide.insert(END, "\n  BETWEEN, BY, CASCADE, CASE, CAST, CHECK, COLLATE, COLUMN,")
    text_aide.insert(END, "\n  COMMIT, CONFLICT, CONSTRAINT, CREATE, CROSS, CURRENT,")
    text_aide.insert(END, "\n  CURRENT_DATE, CURRENT_TIME, CURRENT_TIMESTAMP, DATABASE,")
    text_aide.insert(END, "\n  DEFAULT, DEFERRABLE, DEFERRED, DELETE, DESC, DETACH,")
    text_aide.insert(END, "\n  DISTINCT, DO, DROP, EACH, ELSE, END, ESCAPE, EXCEPT,")
    text_aide.insert(END, "\n  EXCLUDE, EXCLUSIVE, EXISTS, EXPLAIN, FAIL, FILTER, FIRST,")
    text_aide.insert(END, "\n  FOLLOWING, FOR, FOREIGN, FROM, FULL, GENERATED, GLOB,")
    text_aide.insert(END, "\n  GROUP, GROUPS, HAVING, IF, IGNORE, IMMEDIATE, IN, INDEX,")
    text_aide.insert(END, "\n  INDEXED, INITIALLY, INNER, INSERT, INSTEAD, INTERSECT,")
    text_aide.insert(END, "\n  INTO, IS, ISNULL, JOIN, KEY, LAST, LEFT, LIKE, LIMIT,")
    text_aide.insert(END, "\n  MATCH, MATERIALIZED, NATURAL, NO, NOT, NOTHING, NOTNULL,")
    text_aide.insert(END, "\n  NULL, NULLS, OF, OFFSET, ON, OR, ORDER, OTHERS, OUTER,")
    text_aide.insert(END, "\n  OVER, PARTITION, PLAN, PRAGMA, PRECEDING, PRIMARY, QUERY,")
    text_aide.insert(END, "\n  RAISE, RANGE, RECURSIVE, REFERENCES, REGEXP, REINDEX,")
    text_aide.insert(END, "\n  RELEASE, RENAME, REPLACE, RESTRICT, RETURNING, RIGHT,")
    text_aide.insert(END, "\n  ROLLBACK, ROW, ROWS, SAVEPOINT, SELECT, SET, TABLE,")
    text_aide.insert(END, "\n  TEMP, TEMPORARY, THEN, TIES, TO, TRANSACTION, TRIGGER,")
    text_aide.insert(END, "\n  UNBOUNDED, UNION, UNIQUE, UPDATE, USING, VACUUM, VALUES,")
    text_aide.insert(END, "\n  VIEW, VIRTUAL, WHEN, WHERE, WINDOW, WITH, WITHOUT")



def rafraichir_affichage():
    root.update()
    root.update_idletasks()


def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)


def my_read_excel(excelDocument:str):
    """Replace pandas read_excel which we cannot make it read a text column in Excel which is like a number as text.
    """
    wb = openpyxl.load_workbook(excelDocument)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        row_data = []
        for cell in row:
            row_data.append(cell)
        data.append(row_data)

    df_result = pd.DataFrame(data[1:], columns=data[0])
    return df_result


def my_read_csv(csvDocument:str, separator:str=';'):
    """Replace pandas read_csv which we cannot make it read a text column in Csv which is like a number as text.
    """
    data = []
    with open(csvDocument) as f:
        for ligne in f:
            data.append(ligne.rstrip().split(separator))   # rstrip() : to remove the final '\n'

    df_result = pd.DataFrame(data[1:], columns=data[0])
    return df_result


def browse():
    global docSource, doc
    result = filedialog.askopenfilename(title="Select a document", filetypes=(("All supported formats", ("*.csv", "*.json", "*.txt", "*.xlsx", "*.xls", "*.xml")), ("CSV file", "*.csv"), ("Excel doc", ("*.xlsx", "*.xls")), ("JSON file", "*.json"), ("Text file (fixed width)", "*.txt"), ("XML file", "*.xml"), ("All Files", "*.*")))

    if result != '':
        extension = Path(result).suffix.upper()
        if extension in ('.CSV', '.JSON', '.TXT', '.XLSX', '.XLS', '.XML'):
            boutonExecuter["state"] = "disabled"
            requete_sql.delete("1.0","end")
            requete_sql.insert(END,"\n   Wait ...")
            rafraichir_affichage()

            docSource.set(result)

            try:
                # --- open doc
                # Excel
                if extension == '.XLSX' or extension == '.XLS':
                    doc = my_read_excel(docSource.get())
                # CSV
                elif extension == '.CSV':
                    doc = my_read_csv(docSource.get())
                # Text
                elif extension == '.TXT':
                    doc = read_fwf(docSource.get())
                # XML
                elif extension == '.XML':
                    doc = read_xml(docSource.get())
                # JSON
                elif extension == '.JSON':
                    doc = read_json(docSource.get())
                
                requete_sql.delete("1.0","end")

                # if there is a SQL file with the same name we open it automatically
                sql_file = result.upper().replace(extension,'.sql')
                if exists(sql_file):
                    requete_sql.insert(END, '-- select '+', '.join(doc.columns)+'\n')
                    # --- lire SQL file
                    with open(sql_file) as f:
                        for ligne in f:
                            requete_sql.insert(END,ligne)
                else:
                    requete_sql.insert(END, 'select '+', '.join(doc.columns)+'\nfrom doc')

                boutonExecuter["state"] = "normal"
                boutonExporterResultat["state"] = "disabled"
                option_displayGraphToolbar["state"] = "disabled"
                boutonVisualisation["state"] = "disabled"
            except Exception as ErrRead:
                messagebox.showerror('Reading error', ErrRead)
        else:
            messagebox.showerror('Format error',f"This format '{extension}' is not supported.")


def sourceFromClipboard():
    global docSource, doc

    try:
        boutonExecuter["state"] = "disabled"
        requete_sql.delete("1.0","end")
        requete_sql.insert(END,"\n   Wait ...")
        rafraichir_affichage()

        docSource.set('clipboard.source')

        # --- open clipboard
        doc = read_clipboard()

        requete_sql.delete("1.0","end")
        requete_sql.insert(END, 'select '+', '.join(doc.columns)+'\nfrom doc')

        boutonExecuter["state"] = "normal"
        boutonExporterResultat["state"] = "disabled"
        option_displayGraphToolbar["state"] = "disabled"
        boutonVisualisation["state"] = "disabled"

    except pd.errors.EmptyDataError:
        requete_sql.delete("1.0","end")
        messagebox.showerror('Clipboard error',"Clipboard is empty or does not contain readable data.")
    except Exception as Err:
        requete_sql.delete("1.0","end")
        messagebox.showerror('Clipboard error', Err)


def Executer():
    global requete_sql, sql_resultat
    global df
    sql_resultat.delete("1.0","end")

    # print('***',root.width)

    try:
        debutExec = time()
        vTempsExec.set(f"Start: {datetime.now().strftime('%H:%M:%S')}")
        rafraichir_affichage()
        df = sqldf(requete_sql.get("1.0","end"), globals())

        try:
            largeurs = df.astype(str).apply(lambda col: col.str.len()).max()

            # --- Titres
            ligne_valeur = ligne2_valeur = ''
            for ind, colonne in enumerate(df.columns):
                if len(df) > 0:
                    # if len(colonne) > largeurs[ind]:   # si le nom de la colonne est plus long que sa valeur
                    if len(colonne) > largeurs.iloc[ind]:   # si le nom de la colonne est plus long que sa valeur
                        # largeurs[ind] = len(colonne)
                        largeurs.iloc[ind] = len(colonne)
                else:
                    # largeurs[ind] = len(colonne)
                    largeurs.iloc[ind] = len(colonne)
                # ligne_valeur += colonne.ljust(largeurs[ind]) + ' '
                ligne_valeur += colonne.ljust(largeurs.iloc[ind]) + ' '
                # ligne2_valeur += ''.ljust(largeurs[ind],'-') + ' '
                ligne2_valeur += ''.ljust(largeurs.iloc[ind],'-') + ' '
            sql_resultat.insert(END, ligne_valeur+'\n')
            sql_resultat.insert(END, ligne2_valeur+'\n')


            # --- Donnees

            # determiner l'alignement des colonnes
            if len(df) > 0:
                alignements = ['G' for i in range(len(df.columns))]   # initialiser par defaut a GAUCHE
                for iColonne in range(len(df.columns)):
                    if type(df.values[0][iColonne]) is int or type(df.values[0][iColonne]) is float:
                        alignements[iColonne] = 'D'   # à DROITE

            for index, row in df.iterrows():
                ligne_valeur = ''
                # for Colonne in df.columns:
                for iColonne in range(len(df.columns)):
                    if alignements[iColonne] == 'G':
                        ligne_valeur += str(row[df.columns[iColonne]]).ljust(largeurs.iloc[iColonne]) +' '
                    else:
                        ligne_valeur += str(row[df.columns[iColonne]]).rjust(largeurs.iloc[iColonne]) + ' '
                sql_resultat.insert(END, ligne_valeur+'\n')

            # !!! VERY VERY SLOW
            # for ligne in range(len(df)): 
            #     ligne_valeur = ''
            #     for iColonne in range(len(df.columns)):
            #         # if alignements[iColonne] == 'G':
            #         ligne_valeur += str(df.values[ligne][iColonne]).ljust(largeurs[iColonne])
            #         # else:
            #             # ligne_valeur += str(df.values[ligne][iColonne]).rjust(largeurs[iColonne])
            #         # ligne_valeur += ('\n' if iColonne == len(df.columns)-1 else ' ')
            #     sql_resultat.insert(END, ligne_valeur+'\n')

            dureeExec = time() - debutExec
            # vTempsExec.set(f"Durée: {dureeExec:.2f} s , nb.enreg: {len(df)}")
            s = 's' if len(df) > 1 else ''
            vTempsExec.set(f"{len(df)} record{s}  (in {dureeExec:.2f} s)")

            boutonExporterResultat["state"] = "normal"
            option_displayGraphToolbar["state"] = "normal"

            xAxis.set('')
            yAxis.set('')
            Titre.set('')
            XLabel.set('')
            YLabel.set('')
            Legend.set('')


            # --- enregistrer la requête dans un fichier TEMP
            with open('last_request.sql','w') as f:
                print(requete_sql.get("1.0","end"),file=f)

        except Exception as ErrPython:
            messagebox.showerror('Python error', ErrPython)

    except Exception as ErrSql:
        messagebox.showerror('SQL error', ErrSql)


def Exporter():
    global exportFormat, df, docSource

    extension = Path(docSource.get()).suffix.upper()

    lOkExportation = True

    try:
        # --- CSV
        if exportFormat.get() == 'CSV':
            dest = docSource.get().upper().replace(extension,'_export.csv')
            df.to_csv(dest, sep=';')
        # --- EXCEL
        elif exportFormat.get() == 'Excel':
            dest = docSource.get().upper().replace(extension,'_export.xlsx')
            df.to_excel(dest)
        # --- HTML
        elif exportFormat.get() == 'Html':
            dest = docSource.get().upper().replace(extension,'_export.html')
            df.to_html(dest)
        # --- TEXT
        elif exportFormat.get() == 'Text':
            dest = docSource.get().upper().replace(extension,'_export.txt')
            df.to_string(dest)
        # --- JSON
        elif exportFormat.get() == 'JSON':
            dest = docSource.get().upper().replace(extension,'_export.json')
            df.to_json(dest)
        # --- XML
        elif exportFormat.get() == 'XML':
            dest = docSource.get().upper().replace(extension,'_export.xml')
            df.to_xml(dest, index=False)
        else:
            lOkExportation = False
            messagebox.showwarning("Exportation",'Unknown format for exportation.')

        if lOkExportation:
            messagebox.showinfo("Exportation",f"Exportation done to '{dest}'.")
    except Exception as ErrExport:
        messagebox.showerror('Exportation error', ErrExport)


def displayGraph(df: pd.DataFrame, graphType: str, title: str, xAxisNum: str, yAxisNums: str, xLabel: str = None, yLabel: str = None, yLegendLabels: str = None, saveFigure: bool = True, formatSavedFigure:str='png', bar_width=0.4):

    # plt.figure(figsize=(10, 6))
    plt.figure()
    # fig, ax = plt.subplots()

    columnNameUsed_for_xAxis = df.columns[int(xAxisNum)-1]

    if xLabel is None:
        xLabel = columnNameUsed_for_xAxis

    x_positions = np.arange(len(df[columnNameUsed_for_xAxis]))  # Create equally spaced x positions

    columnsNamesUsed_for_yAxis = [df.columns[int(x)-1] for x in yAxisNums.split(',') if x!='']

    lGenerate_yLabel = False
    if yLabel is None:
        lGenerate_yLabel = True
    
    lGenerate_yLegendLabels = False
    if yLegendLabels is None:
        lGenerate_yLegendLabels = True
    
    if lGenerate_yLabel or lGenerate_yLegendLabels:
        valForLabel = ''
        for colName in columnsNamesUsed_for_yAxis:
            valForLabel += ('' if valForLabel == '' else ',') + colName
        if lGenerate_yLabel:
            yLabel = valForLabel
        if lGenerate_yLegendLabels:
            yLegendLabels = valForLabel
        
    columnsLegenUsed_for_yAxis = [x for x in yLegendLabels.split(',')]

    colors = (('lightblue','blue'),('lightcoral','coral'),('lightgreen','green'),('thistle','purple'),('lightseagreen','seagreen'),('burlywood','brown'),('pink','hotpink'),('orangered','orange'),('silver','gray'),('violet','mediumvioletred'))

    for i, colName in enumerate(columnsNamesUsed_for_yAxis):
        valLabel = columnsLegenUsed_for_yAxis[i] if i < len(columnsLegenUsed_for_yAxis) else ''
        match graphType:
            case 'Area':
                plt.fill_between(x_positions, df[colName], color=colors[i][0], alpha=0.4)
                plt.plot(x_positions, df[colName], color=colors[i][1], alpha=0.6, label=valLabel)
            case 'Bar':
                plt.bar(x_positions + i*bar_width, df[colName], width=bar_width, label=valLabel)
            case 'Barh':
                plt.barh(x_positions + i*bar_width, df[colName], height=bar_width, label=valLabel)
            case 'Line':
                plt.plot(x_positions + i*bar_width, df[colName], label=valLabel)
            case 'Pie':
                plt.pie(df[colName], explode=tuple([0.1 for x in range(len(df[columnNameUsed_for_xAxis]))]), labels=df[columnNameUsed_for_xAxis], autopct='%1.1f%%', shadow=True, startangle=90)
                # plt.pie(df[colName], explode=tuple([0.1 if x==0 else 0 for x in range(len(df[columnNameUsed_for_xAxis]))]), labels=df[columnNameUsed_for_xAxis], autopct='%1.1f%%', shadow=True, startangle=90)
                # plt.pie(df[colName], labels=df[columnNameUsed_for_xAxis], autopct='%1.1f%%', shadow=True, startangle=90)
                # plt.pie(df[colName], labels=x_positions, autopct='%1.1f%%', shadow=True, startangle=90)
            case 'Scatter':
                plt.scatter(x_positions + i*bar_width, df[colName], label=valLabel)
        # if i >= len(columnsLegenUsed_for_yAxis) Then the Legend is empty or doesn't contain sufficient elements

    """
    # --- Bar Chart
    plt.bar(x_positions, df['NOMBRE'], width=bar_width, label='Nbre')
    plt.bar(x_positions + bar_width, df['MONTANT'], width=bar_width, label='Mont')

    # --- Bar Horizental Chart
    plt.barh(x_positions, df['NOMBRE'], height=bar_width, label='Nbre')
    plt.barh(x_positions + bar_width, df['MONTANT'], height=bar_width, label='Mont')

    # --- Line Chart
    plt.plot(x_positions, df['NOMBRE'], label='Nbre')
    plt.plot(x_positions + bar_width, df['MONTANT'], label='Mont')

    # --- Scatter Plot
    plt.scatter(x_positions, df['NOMBRE'], label='Nbre')
    plt.scatter(x_positions + bar_width, df['MONTANT'], label='Mont')

    # --- Pie Chart
    plt.pie(df['NOMBRE'], labels=df['CATEGORIE'], autopct='%1.1f%%')

    # --- Area Chart
    plt.fill_between(x_positions, df['NOMBRE'], color="skyblue", alpha=0.4)
    plt.plot(x_positions, df['NOMBRE'], color="Slateblue", alpha=0.6, label='Nbre')

    plt.fill_between(x_positions, df['MONTANT'], color="lightgreen", alpha=0.4)
    plt.plot(x_positions, df['MONTANT'], color="green", alpha=0.6, label='Mont')
    """

    if graphType != 'Barh':
        plt.xlabel(xLabel)
        if graphType != 'Pie':
            plt.ylabel(yLabel)
    else:
        # For 'Barh' type we invert between axes
        plt.xlabel(yLabel)
        plt.ylabel(xLabel)

    if graphType != 'Pie':
        if graphType != 'Barh':
            plt.xticks(x_positions, df[columnNameUsed_for_xAxis])  # Set x-axis labels back to the original non-numeric values
        else:
            plt.yticks(x_positions, df[columnNameUsed_for_xAxis])  # Set x-axis labels back to the original non-numeric values
    plt.title(title)

    # if yLegendLabels != '' and graphType != 'Pie':
    if yLegendLabels != '':
        if graphType != 'Pie':
            plt.legend()
        else:
            plt.legend(loc='upper left', bbox_to_anchor=(-0.3, 1))

    if graphType != 'Pie':
        # Rotate x-axis labels for better readability if needed
        if graphType != 'Barh':
            plt.xticks(rotation=45)
        else:
            plt.yticks(rotation=45)

        plt.tight_layout()

    fig = plt.gcf()
    fig_width, fig_height = fig.get_size_inches()
    pos_x = int((screen_width - fig_width * fig.dpi) / 2)
    pos_y = int((screen_height - fig_height * fig.dpi) / 2)

    # Set the position
    plt.get_current_fig_manager().window.wm_geometry(f"+{pos_x}+{pos_y}")

    if saveFigure:
        fig.savefig(f'last_figure.{formatSavedFigure}')

    plt.show()


def changerDisplayGraphToolbar():
    global frame_2bis, frame_2bis_row
    if displayGraphToolbar.get() == 1:
        frame_2bis.grid(row=frame_2bis_row, column=0, columnspan=2, sticky='WE')
    else:
        frame_2bis.grid_remove()


def isOk_value_xAxis(valeur):
    global df
    if valeur.isdigit():
        if int(valeur) <= len(df.columns) and valeur not in yAxis.get().split(','):
            return True
        else:
            return False
    else:
        return False

def on_validate_xAxis(valeur):
    if valeur == '':
        lOkValeur = True
    else:
        lOkValeur = isOk_value_xAxis(valeur)

    if valeur != '' and lOkValeur and isOk_value_yAxis(yAxis.get()):
        boutonVisualisation["state"] = "normal"
        # set x-Label
        XLabel.set(df.columns[int(valeur)-1])
    else:
        boutonVisualisation["state"] = "disabled"

    return lOkValeur


validate_xAxis = root.register(on_validate_xAxis)


def isOk_value_yAxis(valeur):
    global df
    lAllOk = True
    if valeur != '':
        for val in valeur.split(','):
            if val != '':
                if not val.isdigit() or int(val) > len(df.columns) or val == xAxis.get():
                    lAllOk = False
    return lAllOk

def on_validate_yAxis(valeur):
    if valeur == '':
        lOkValeur = True
    else:
        lOkValeur = isOk_value_yAxis(valeur)
    
    if isOk_value_xAxis(xAxis.get()) and valeur != '' and lOkValeur:
        boutonVisualisation["state"] = "normal"
        # set y-Label & Legend
        val_y_label = ''
        for colName in [df.columns[int(x)-1] for x in valeur.split(',') if x!='']:
            val_y_label += ('' if val_y_label == '' else ',') + colName
        YLabel.set(val_y_label)
        Legend.set(val_y_label)
    else:
        boutonVisualisation["state"] = "disabled"

    return lOkValeur

validate_yAxis = root.register(on_validate_yAxis)


# * ===========================================================================
# * ===========================================================================
# * ===========================================================================


global doc

current_row = 0

root.grid_columnconfigure(0, weight=1)   # makes the widgets with sticky='WE' fill all the width

# Colors
bg_color_default = "#b7cbf3"
bg_color_default_GraphToolbar = "#31579D" #"#6b84b6"
fg_color_default_Label = "#31579D"
fg_color_default_Label_GraphToolbar = "#b7cbf3"
fg_color_default_Label_ShortcutKey = "maroon"
fg_color_default_Label_Copyright = "gray"
bg_color_default_Button = "#31579D"
fg_color_default_Button = "#31579D"

# --- Default styles
style = ttk.Style()
style.configure('TFrame', background=bg_color_default)
style.configure('GraphToolbar.TFrame', background=bg_color_default_GraphToolbar)
style.configure('TLabel', foreground=fg_color_default_Label, background=bg_color_default, font=("Helvetica", 10))
style.configure('GraphToolbar.TLabel', foreground=fg_color_default_Label_GraphToolbar, background=bg_color_default_GraphToolbar, font=("Helvetica", 10))
style.configure('TButton', foreground=fg_color_default_Button, background=bg_color_default_Button, font=("Helvetica", 10))

# * --------- Frame 1 : doc source
# --- docSource
frame_1 = ttk.Frame(root)
frame_1.grid(row=current_row, column=0, columnspan=5, sticky='WE')

lblPath = ttk.Label(frame_1, text="Source Document:")
# lblPath.configure(foreground=fg_color_default_Label, background=bg_color_default)
lblPath.pack(side=LEFT, padx=5, pady=10)

docSource = StringVar()
entryPath = ttk.Entry(frame_1, textvariable=docSource, width=70)
entryPath.pack(side=LEFT, padx=5)
boutonPath = ttk.Button(frame_1, text='...', width=3, command=browse)
Tooltip(boutonPath, "Click here to select a document")
boutonPath.pack(side=LEFT)


boutonFromClipboard = ttk.Button(frame_1, text='From clipboard', command=sourceFromClipboard)
Tooltip(boutonFromClipboard, "Click here to use data from clipboard")
boutonFromClipboard.pack(side=LEFT, padx=15)

boutonSqlAide = ttk.Button(frame_1, text='SQL help', width=8, command=ouvrir_Sql_Aide)
boutonSqlAide.pack(side=RIGHT, padx=5)

current_row += 1


# * --------- Frame 2 : requete SQL
# --- requete_sql
frame_requete_sql = ttk.Frame(root)
frame_requete_sql.grid(row=current_row, column=0, columnspan=5, sticky='WE')
current_row += 1

scrollbar_requete_sql_x = ttk.Scrollbar(frame_requete_sql, orient= HORIZONTAL)
scrollbar_requete_sql_x.pack(side=BOTTOM, fill=X)
scrollbar_requete_sql_y = ttk.Scrollbar(frame_requete_sql)
scrollbar_requete_sql_y.pack(side=RIGHT, fill=Y)

requete_sql = Text(frame_requete_sql, xscrollcommand=scrollbar_requete_sql_x.set, yscrollcommand=scrollbar_requete_sql_y.set, wrap='none', height=8)   #  , width=80
requete_sql.pack(expand=True, fill=BOTH)

scrollbar_requete_sql_x.config(command=requete_sql.xview)
scrollbar_requete_sql_y.config(command=requete_sql.yview)


# * --------- Frame 3 : bouton Executer
frame_2 = ttk.Frame(root)
frame_2.grid(row=current_row, column=0, columnspan=5, sticky='WE')

# --- bouton Executer
lblF8 = ttk.Label(frame_2,text="F8")
# ttk.Style().configure("Bold.TLabel", font=("TkDefaultFont", 9, "bold"))
style.configure("Bold.TLabel", font=("TkDefaultFont", 9, "bold"), foreground=fg_color_default_Label_ShortcutKey, background=bg_color_default)
lblF8.configure(style="Bold.TLabel")
lblF8.pack(side=LEFT, padx=7)
boutonExecuter = ttk.Button(frame_2, text="  Execute  ", state="disabled", command=Executer)
boutonExecuter.pack(side=LEFT, padx=5, pady=10)

# --- Temps d'execution
vTempsExec = StringVar()
lblTempsExec = ttk.Label(frame_2, textvariable=vTempsExec)
# lblTempsExec.configure(foreground=fg_color_default_Label, background=bg_color_default)
lblTempsExec.pack(side=LEFT, padx=10, pady=10)

# --- bouton Quit
boutonFermer = ttk.Button(frame_2, text="Quit", width=6, command=root.quit)
boutonFermer.pack(side=RIGHT, padx=10)

# --- option "Graph"
displayGraphToolbar = IntVar()
option_displayGraphToolbar = ttk.Checkbutton(frame_2, text="Chart toolbar", state="disabled", variable=displayGraphToolbar, onvalue=1, offvalue=0, width=12, command=changerDisplayGraphToolbar)
ttk.Style().configure("Custom.TCheckbutton", foreground=fg_color_default_Label, background=bg_color_default)
option_displayGraphToolbar.configure(style="Custom.TCheckbutton")
Tooltip(option_displayGraphToolbar, "Display chart toolbar to visualize data")
option_displayGraphToolbar.pack(side=RIGHT,padx=20)

# --- Export
boutonExporterResultat = ttk.Button(frame_2, text="Export", state="disabled", command=Exporter)
boutonExporterResultat.pack(side=RIGHT, padx=10)

exportFormat = StringVar()
exportFormat.set("Excel")
combo_exportFormat = ttk.Combobox(frame_2, textvariable = exportFormat, width=7, values=('CSV', 'Excel', 'JSON', 'Html', 'Text', 'XML'), state="readonly")
combo_exportFormat.pack(side=RIGHT)

lblExport = ttk.Label(frame_2,text="Export format:")
# lblExport.configure(foreground=fg_color_default_Label, background=bg_color_default)
lblExport.pack(side=RIGHT, padx=10, pady=10)   # , bg=bg_color_default

current_row += 1


# * --------- Frame 4 : Graph
frame_2bis = ttk.Frame(root, style='GraphToolbar.TFrame')
frame_2bis_row = current_row
frame_2bis.grid(row=frame_2bis_row, column=0, columnspan=2, sticky='WE')
frame_2bis.grid_remove()

# --- Graph type
lblGraphType = ttk.Label(frame_2bis, text="Type:")
lblGraphType.configure(style='GraphToolbar.TLabel')
lblGraphType.pack(side=LEFT, padx=5, pady=6)

graphType = StringVar()
graphType.set("Bar")
combo_graphType = ttk.Combobox(frame_2bis, textvariable=graphType, width=6, values=('Area', 'Bar', 'Barh', 'Line', 'Pie', 'Scatter'), state="readonly")
combo_graphType.pack(side=LEFT)

# --- x-axis
lblXAxis = ttk.Label(frame_2bis, text="x-axis column:")
lblXAxis.configure(style='GraphToolbar.TLabel')
lblXAxis.pack(side=LEFT, padx=5)

# xAxis = StringVar(value=1)
xAxis = StringVar()
entryXAxis = ttk.Entry(frame_2bis, textvariable=xAxis, width=3, validate="key", validatecommand=(validate_xAxis, '%P'))
Tooltip(entryXAxis, "Number of the column to use in X axis")
entryXAxis.pack(side=LEFT) #, padx=5)

# --- y-axis
lblYAxis = ttk.Label(frame_2bis, text="y-axis columns (,):")
lblYAxis.configure(style='GraphToolbar.TLabel')
lblYAxis.pack(side=LEFT, padx=5) #, pady=10)

# yAxis = StringVar(value=2)
yAxis = StringVar()
entryYAxis = ttk.Entry(frame_2bis, textvariable=yAxis, width=6, validate="key", validatecommand=(validate_yAxis, '%P'))
Tooltip(entryYAxis, "Number(s) of the column(s) to use in Y axis\nseparated by ',' if there is more than one")
entryYAxis.pack(side=LEFT) #, padx=5)

# --- Title
lblTitre = ttk.Label(frame_2bis, text="Title:")
lblTitre.configure(style='GraphToolbar.TLabel')
lblTitre.pack(side=LEFT, padx=5)

Titre = StringVar()
entryTitre = ttk.Entry(frame_2bis, textvariable=Titre, width=10)
Tooltip(entryTitre, "Title of the chart\n(empty = no title)")
entryTitre.pack(side=LEFT) #, padx=5)

# --- x-Label
lblXLabel = ttk.Label(frame_2bis, text="x-label:")
lblXLabel.configure(style='GraphToolbar.TLabel')
lblXLabel.pack(side=LEFT, padx=5)

XLabel = StringVar()
entryXLabel = ttk.Entry(frame_2bis, textvariable=XLabel, width=10)
Tooltip(entryXLabel, "Label for X axis\n(empty = no label)")
entryXLabel.pack(side=LEFT) #, padx=5)

# --- y-Label
lblYLabel = ttk.Label(frame_2bis, text="y-label:")
lblYLabel.configure(style='GraphToolbar.TLabel')
lblYLabel.pack(side=LEFT, padx=5)

YLabel = StringVar()
entryYLabel = ttk.Entry(frame_2bis, textvariable=YLabel, width=10)
Tooltip(entryYLabel, "Label for Y axis\n(empty = no label)")
entryYLabel.pack(side=LEFT) #, padx=5)

# --- Legend
lblLegend = ttk.Label(frame_2bis, text="Legend (,):")
lblLegend.configure(style='GraphToolbar.TLabel')
lblLegend.pack(side=LEFT, padx=5)

Legend = StringVar()
entryLegend = ttk.Entry(frame_2bis, textvariable=Legend, width=10)
Tooltip(entryLegend, "Legend for column(s) used for Y axis\nseparated by ',' if there is more than one\n(empty = no legend)")
entryLegend.pack(side=LEFT) #, padx=5)


# --- Visualization button
boutonVisualisation = ttk.Button(frame_2bis, text='Visualization', state="disabled", width=12, command= lambda: displayGraph(df, graphType.get(), Titre.get(), xAxis.get(), yAxis.get(), XLabel.get(), YLabel.get(), Legend.get()))
# boutonVisualisation = ttk.Button(frame_2bis, text='Graph', width=6, state="disabled", command=displayGraph)
boutonVisualisation.pack(side=LEFT, padx=5)

current_row += 1


# * --------- Frame 5 : Résultat
# --- sql_resultat
frame_sql_resultat = ttk.Frame(root)
# frame_sql_resultat = ttk.Frame(root, borderwidth=2, relief=GROOVE, height=18)
frame_sql_resultat.grid(row=current_row, column=0, columnspan=4, sticky="WENS")
current_row += 1

scrollbar_sql_resultat_x = Scrollbar(frame_sql_resultat, orient= HORIZONTAL)
scrollbar_sql_resultat_x.pack(side=BOTTOM, fill=X)
scrollbar_sql_resultat_y = Scrollbar(frame_sql_resultat)
scrollbar_sql_resultat_y.pack(side=RIGHT, fill=Y)

sql_resultat = Text(frame_sql_resultat, yscrollcommand= scrollbar_sql_resultat_y.set,
xscrollcommand = scrollbar_sql_resultat_x.set, wrap='none') #, height=15) #, width=80)
sql_resultat.pack(expand=True, fill=BOTH)

scrollbar_sql_resultat_x.config(command=sql_resultat.xview)
scrollbar_sql_resultat_y.config(command=sql_resultat.yview)


# * --------- Frame 6 : Export
# frame_3 = ttk.Frame(root)
# frame_3.grid(row=current_row, column=0, columnspan=5, sticky='WE')

# # --- Export
# lblExport = ttk.Label(frame_3,text="Export format:")
# # lblExport.configure(foreground=fg_color_default_Label, background=bg_color_default)
# lblExport.pack(side=LEFT, padx=10, pady=10)   # , bg=bg_color_default

# exportFormat = StringVar()
# exportFormat.set("Excel")
# combo_exportFormat = ttk.Combobox(frame_3, textvariable = exportFormat, width=7, values=('CSV', 'Excel', 'JSON', 'Html', 'Text', 'XML'), state="readonly")
# combo_exportFormat.pack(side=LEFT)

# boutonExporterResultat = ttk.Button(frame_3, text="Export", state="disabled", command=Exporter)
# boutonExporterResultat.pack(side=LEFT, padx=10)

# # --- bouton Quit
# boutonFermer = ttk.Button(frame_3, text="Quit", width=6, command=root.quit)
# boutonFermer.pack(side=RIGHT, padx=10)

# current_row += 1

# # --- separator
# separator = ttk.Frame(root, height=2)
# separator.grid(row=current_row, column=0)
# current_row += 1

# --- copyright label
m_smati = (77, 111, 104, 97, 109, 101, 100, 32, 83, 77, 65, 84, 73)
lbl_copyright = ttk.Label(root, text="By " + ''.join((chr(i) for i in m_smati)) + " , September 2023", anchor=CENTER) #, height=2)  #, fg="#808080"
lbl_copyright.configure(foreground=fg_color_default_Label_Copyright) #, background=bg_color_default)
lbl_copyright.grid(row=current_row, column=0, columnspan=5, ipady=10, sticky="WENS")
current_row += 1

# Bind the window resize event to the on_resize function
root.bind('<Configure>', on_resize)

root.mainloop()
