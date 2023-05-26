import hetpy as hp
import pyodbc
import pandas as pd

# Set functions
def ms_connector(user,passw,server_name,db,driver):
    session = pyodbc.connect('DRIVER=' + driver + ';SERVER=' + server_name + ';DATABASE=' + db + ';UID=' + user+ ';PWD=' + passw)
    return session

print('Enter Credentials...')
user = hp.inputter('Enter Username: ','str')
passw = hp.input_pass()

print('Connecting...')
cnxn = ms_connector(user,passw,'YOUR_SERVER_HERE','YOUR_DB_HERE','SQL Server')

cnxn.execute('SELECT TOP 5 * INTO #TMP_Table FROM YOUR_DB_HERE.dbo.Table')

df = pd.read_sql('SELECT * FROM #TMP_Table',session)


