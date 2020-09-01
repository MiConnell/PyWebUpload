# PyWebUpload
This is a locally-hosted web upload tool that is heavily reliant on the great work of Plotly's Dash [Web Upload Component](https://dash.plotly.com/dash-core-components/upload).

This reads from a spreadsheet that is either selected from the built-in file picker or dragged and dropped to the website.

Currently set up for SQL Server, but credentials/relevant libraries can be edited to work with any other relational database.

## Needed Script Edits

### Credentials

### Spreadsheet Columns

This can be as many columns as needed, define variables and use python's `zip` function

ex.

    for n, d, c, s in zip(uploadSheet(name), uploadSheet(date), uploadSheet(city), uploadSheet(state):
         ...

### SQL Statement

The `UPDATE` / `INSERT` statement will need to be altered in the file as well.

ex.

    """
    UPDATE CUSTOMERS
        SET DATE = '{str(b)}',
        CITY = '{str(b)}',
        STATE = '{str(b)}'
    WHERE NAME = '{str(b)}'
    """

### Change Drive for Error Log (optional)
Currently `out = "Drive/Error Log.txt"` will need to be updated to a valid drive for error logging


