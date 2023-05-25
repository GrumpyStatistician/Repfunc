# Repfunc
Basic Reporting functions for DA &amp; reporting

## Available Functions
<b>inputter(question,dtype,options=None,file_type='.xlsx',file_dtype=str) - </b>
Inputter function integrates error handling into the base input() function. Can be utilized with command line programs that require simple inputs. 
Currently accepts str, int, date, and file as dtype. Options arg supported by int and str dtype only.  File dtype arg returns all .xlsx file types within directory, will possibly add more options in the future. Added sheetname arg to allow selection from dataframe with multiple tabs, default is 0.
``` python
#str with options
inputter('Enter Auth type (LDAP,TD2): ','str',['ldap','td2'])
#int with options
inputter('What is the meaning of life?: ','int',42) 
#date
inputter('Please enter start date (svc_from_dt): ','date')
#file
inputter('Please select from the following options:','file')
```
<b>input_pass(txt='Enter Your Password: ') - </b>
Input_pass function uses getpass module to mask your password in most applications (warning will appear if not applicable). Text defaults to 'Enter Your Password: ', just enter in your text if you want something else.
``` python
input_pass('Whatever you want your prompt to be here')
```

<b>connector(user,passw,auth,sys,drive, ver='2.0',log=False,meth="odbc") - </b>Connection function utilized connect to Teradata. Requires entry of username, password, and auth type (LDAP,TD2). Other arg are defaulted to most commonly used Teradata server.</br>
``` python
connector(user,passw,auth)
```
<b>create_wb(filename,add=None,engine='xlsxriter',ind=False) - </b>Generates excel workbook object and writer. Add arg allows you to add in a description to your excel file, type in whatever string you want or type 'date' to add in today's date</br>
``` python
workbook,writer = hp.create_wb('Data')
```
<b>create_tab(df,writer,workbook,tab_name='sheet1',col=15,row=15,ind= False) - </b> Creates Excel tab from pandas dataframe. Defaults can be updated to reflect desired format.
``` python
create_tab(df,writer,workbook,'summary tab')
```
<b>create_table(df, table_name, infer_dtype=None) - </b>Generates sql create table based on pandas dataframe column names, character lenth with optional data type inference (defaults to varchar). Added optional manual primary key setting, recommend using this on any table that you're planning on keeping on the server.
``` python
create_table(df,'db.Table_Test',primary_index='some_key')
```
<b>insert_df(df, table_name) - </b>Generates sql insert based on pandas dataframe column names
``` python
insert_df(df,'db.Table_Test')
```
<b>df_load(insert,df,session,chunk=False) - </b>Loads data into Teradata database using insert sql, pandas dataframe, and session object. Change chunk to True when working with 10k rows or greater.
``` python
df_load(insert,df,session,chunk=True)
```
<b>emailer_head(subject,to,cc='',disp='N') - </b>Creates message object that contains email subject and sendee args with optional CC. Disp arg enables email popup when set to 'Y'.
``` python
emailer_head('Here You go','Sendee@gmail.com')
```
<b>emailer_body(body,message,attach=None,disp='N') - </b>For adding email body and attachments to message object. Set attach arg to filename to attach to email.
``` python
emailer_body('''
<p>Good Afternoon Dr.Seldon,</p>

<p>Refer to the attached data. Password is PASS.</p>

<p>Best,</p>

<p>The Dude</p>
''',message)
```
<b>protect_wb(input_wb,output_wb,passw,date_include='y') - </b> Adds password protection to Excel workbook with option to add dates to updated file name.
 ``` python
 protect_wb('Stuff.xlsx','Final','PASS')
  ```
 <b>create_folder(folder_name) - </b>Creates a folder within the current working directory. folder_name (str) arg is name of folder to be created.
 ``` python
 create_folder('202203')
