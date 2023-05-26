# Connect via MSSQL
Will be integrating MSSQL into Repfunc in the future. In the meantime, refer to Connect Example.pyscript for how to connect using pyodbc.
```python

import hetpy as hp
import pyodbc
import pandas as pd

# Set functions
def ms_connect(user,passw,server_name,db,driver):
    cnxn = pyodbc.connect('DRIVER=' + driver + ';SERVER=' + server_name + ';DATABASE=' + db + ';UID=' + user+ ';PWD=' + passw)
    return cnxn

print('Enter Credentials...')
user = hp.inputter('Enter Username: ','str')
passw = hp.input_pass()

print('Connecting...')
cnxn = ms_connect(user,passw,'YOUR_SERVER_HERE','YOUR_DB_HERE','SQL Server')

cnxn.execute('SELECT TOP 5 * INTO #TMP_Table FROM YOUR_DB_HERE.dbo.Table')

df = pd.read_sql('SELECT * FROM #TMP_Table',cnxn)
```
