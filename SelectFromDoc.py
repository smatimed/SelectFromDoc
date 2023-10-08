from pandas import read_fwf, read_json, read_xml, read_clipboard
import pandas as pd
from pandasql import sqldf
import openpyxl

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
root_width = 800
root_height = 600
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
root.geometry(f'{root_width}x{root_height}+{(screen_width-root_width) // 2}+{(screen_height-root_height) // 2}')

root.title('SQL from Document (Excel, Csv, Json, Text, Xml) or Clipboard')
# root.iconbitmap(path.abspath(path.join(path.dirname(__file__), 'SelectFromDoc.ico')))
root.iconbitmap('SelectFromDoc.ico')

def on_resize(event):
    # Update the size of 'frame_sql_resultat' when the window is resized
    root.grid_rowconfigure(3, weight=1)
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
                    if len(colonne) > largeurs[ind]:   # si le nom de la colonne est plus long que sa valeur
                        largeurs[ind] = len(colonne)
                else:
                    largeurs[ind] = len(colonne)
                ligne_valeur += colonne.ljust(largeurs[ind]) + ' '
                ligne2_valeur += ''.ljust(largeurs[ind],'-') + ' '
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
                        ligne_valeur += str(row[df.columns[iColonne]]).ljust(largeurs[iColonne]) +' '
                    else:
                        ligne_valeur += str(row[df.columns[iColonne]]).rjust(largeurs[iColonne]) + ' '
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

# ===========================================================================


global doc

current_row = 0

root.grid_columnconfigure(0, weight=1)   # makes the widgets with sticky='WE' fill all the width

# Colors
bg_color_default = "#b7cbf3"
fg_color_default_Label = "#31579D"
fg_color_default_Label_ShortcutKey = "maroon"
fg_color_default_Label_Copyright = "gray"
bg_color_default_Button = "#31579D"
fg_color_default_Button = "#31579D"

# --- Default styles
style = ttk.Style()
style.configure('TFrame', background=bg_color_default)
style.configure('TLabel', foreground=fg_color_default_Label, background=bg_color_default, font=("Helvetica", 10))
style.configure('TButton', foreground=fg_color_default_Button, background=bg_color_default_Button, font=("Helvetica", 10))

# --- docSource
frame_1 = ttk.Frame(root)
frame_1.grid(row=current_row, column=0, columnspan=5, sticky='WE')

lbl_path = ttk.Label(frame_1, text="Source Document:")
# lbl_path.configure(foreground=fg_color_default_Label, background=bg_color_default)
lbl_path.pack(side=LEFT, padx=5, pady=10)

docSource = StringVar()
ent_path = ttk.Entry(frame_1, textvariable=docSource, width=70)
ent_path.pack(side=LEFT, padx=5)
but_path = ttk.Button(frame_1, text='...', width=3, command=browse)
but_path.pack(side=LEFT)

boutonFromClipboard = ttk.Button(frame_1, text='From clipboard', command=sourceFromClipboard)
boutonFromClipboard.pack(side=LEFT, padx=15)

boutonSqlAide = ttk.Button(frame_1, text='SQL help', width=8, command=ouvrir_Sql_Aide)
boutonSqlAide.pack(side=RIGHT, padx=5)

current_row += 1


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


# --------- Frame
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

current_row += 1


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

# --------- Frame
frame_3 = ttk.Frame(root)
frame_3.grid(row=current_row, column=0, columnspan=5, sticky='WE')

# --- Export
lblExport = ttk.Label(frame_3,text="Export format:")
# lblExport.configure(foreground=fg_color_default_Label, background=bg_color_default)
lblExport.pack(side=LEFT, padx=10, pady=10)   # , bg=bg_color_default

exportFormat = StringVar()
exportFormat.set("Excel")
combo_exportFormat = ttk.Combobox(frame_3, textvariable = exportFormat, width=7, values=('CSV', 'Excel', 'JSON', 'Html', 'Text', 'XML'), state="readonly")
combo_exportFormat.pack(side=LEFT)

boutonExporterResultat = ttk.Button(frame_3, text="Export", state="disabled", command=Exporter)
boutonExporterResultat.pack(side=LEFT, padx=10)

# --- bouton Quit
boutonFermer = ttk.Button(frame_3, text="Quit", command=root.quit)
boutonFermer.pack(side=RIGHT, padx=10)

current_row += 1

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
