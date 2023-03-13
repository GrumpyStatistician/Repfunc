import teradata 
import pandas as pd
import os
from datetime import datetime
import numpy as np
import win32com.client as client
from math import ceil
from getpass import getpass

# Create Connection String (teradata only)
def connector(user,passw,auth,sys,drive,ver='2.0',log=False,meth="odbc"):
    udaExec = teradata.UdaExec(appName='connectit', version=ver, logConsole=log)
    session = udaExec.connect(method=meth, system=sys,username=user, password=passw, authentication=auth,driver=drive)
    return session

# Handle Inputs for basic ad hoc cmd based reporting
def inputter(question,dtype,options=None,sheetname=0,file_type='.xlsx',file_dtype=str):
    while True:
            if dtype.lower()=='file':
                excel_files = [f for f in os.listdir() if f.endswith(f'{file_type}')]
                print(question)
                for i in excel_files:
                    index = excel_files.index(i)
                    print(f'{index} {i}')
                answer= input("Answer: ")
                try:
                    answer = int(answer)
                    file = excel_files[answer]
                    answer2 = input(f'{file} - Is this correct?(Y/N): ')
                    if answer2.lower() =="y":
                        print('Reading file to dataframe...')
                        df = pd.read_excel(file, sheet_name=sheetname,dtype=file_dtype,keep_default_na=False,engine='openpyxl')
                        return df
                    else:
                        print('Please select a valid file.')
                        continue
                except Exception as e:
                    print(e)
                    continue   
            else:
                x = input(question)
                try:
                    if dtype == 'str' and options==None:
                        return x
                    elif dtype =='str' and options is not None:

                        choices = [x.lower() for x in options]
                        if x.lower() in (choices):
                            return x
                        else:
                            print('Please enter a valid option.')
                            continue
                    elif dtype == 'int' and options==None:
                        x=int(x)
                        return x
                    elif dtype == 'int' and options is not None:
                        x=int(x)
                        if x in options:
                            return x
                        else:
                            print('Restarting.')
                            continue
                    elif dtype.lower() == 'date':
                        datetime.strptime(x, '%Y-%m-%d')
                        answer = input(f'{x} - Is this correct?(Y/N): ')
                        if answer.lower() == "y":
                            return x
                        else:
                            print('Please enter a valid date.')
                            continue
                    else:
                        print('Please selete a valid dtype.')
                    break
                except Exception as e:
                    print(e)
                    break
                    
# Mask Password
def input_pass(txt='Enter Your Password: '):
    x = getpass(prompt=txt,stream=None)
    return x
                    
# Create simple workbook using pandas xlsxwriter              
def create_wb(filename,add=None,engine='xlsxriter',ind=False):
    if add is None:
        writer = pd.ExcelWriter(f'{filename}.xlsx', engine='xlsxwriter')
        
    elif add.lower() == 'date':
        now = datetime.now()
        date = now.strftime('%Y-%m-%d')
        writer = pd.ExcelWriter(f'{filename}-{date}.xlsx', engine='xlsxwriter')
    else:
        writer = pd.ExcelWriter(f'{filename}-{add}.xlsx', engine='xlsxwriter')
    workbook = writer.book
    return workbook,writer
# Create tabs for workbooks 

def create_tab(df,writer,workbook,tab_name='sheet1',col=15,row=15,ind= False):
    df_len=len(df.columns)
    max_col = chr(ord('@') + df_len)
    row_all = workbook.add_format({'align': 'center'})
    column_all = workbook.add_format({'align': 'left'})
    df.to_excel(writer, sheet_name=tab_name, index=ind)
    worksheet = writer.sheets[tab_name]
    worksheet.set_row(0, row, row_all)
    worksheet.set_column(f'A:{max_col}', col, column_all) 
    return worksheet
# Create table formatting from df

def create_table(df, table_name,primary_index=None, infer_dtype=False):

    # use numpy to vectorize
    
    measurer = np.vectorize(len)
    col_len = measurer(df.values.astype(str)).max(axis=0)

    # get column names
    
    col_lst_init = [i for i in df.columns]
    col_lst1 = [i.replace(" ","_") for i in col_lst_init] #add underscore with columns with blanks
    col_lst2 = [i.replace("'","") for i in col_lst1] #remove any single quotes
    col_lst = [i.replace("-","_") for i in col_lst2] #replace dashes with underscore
    col_index = [ (f'{i}') for i in col_lst]
    col_str = '"'+'","'.join(col_index)+'"' #string with comma & added quotes
    
    # get table info
    
    if infer_dtype == True:
        dtype_list = []
        for i in df:
                dtype= df[f'{i}'].infer_objects().dtypes
                dtype_list.append(dtype)
        final_type = []
        for d,c in zip(dtype_list,col_len):
                #can add more conditions to handle different dtypes
                if d == "int64" or d == "int":
                        final_type.append('int')
                elif d == "float" or d == "float64":
                        final_type.append('float')
                else: #probably not worth handling datetimes(teradata can interpret)
                    final_type.append(f'varchar({c+5})')
        info_lst = [(f'"{x}" {y},') for x,y in zip(col_lst,final_type)] 
    else:
        info_lst = [(f'"{x}" varchar({y+5}),') for x,y in zip(col_lst,col_len)]
    info_lst[-1] = info_lst[-1][:-1] #remove comma
    info_str = "\n".join(info_lst) #string with new lines
    if primary_index is not None:
        col_str= primary_index
    final= f'''
    CREATE TABLE {table_name}
    (
    {info_str}
    )
    PRIMARY INDEX ({col_str})
    ;
    '''
    return final
# Create insert formatting from df

def insert_df(df, table_name):
    
    # use numpy to vectorize
    
    measurer = np.vectorize(len)
    col_len = measurer(df.values.astype(str)).max(axis=0)

    # get column names
    
    col_lst_init = [i for i in df.columns]
    col_lst1 = [i.replace(" ","_") for i in col_lst_init] #add underscore with columns with blanks
    col_lst2 = [i.replace("'","") for i in col_lst1] #remove any single quotes
    col_lst = [i.replace("-","_") for i in col_lst2] #replace dashes with underscore
    col_index = [ (f'{i}') for i in col_lst]
    col_str = '"'+'","'.join(col_index)+'"' #string with comma & added quotes
    info_lst = [(f'"{x}" varchar({y}),') for x,y in zip(col_lst,col_len)]
    info_lst[-1] = info_lst[-1][:-1] #remove comma
    info_str = "\n".join(info_lst) #string with new lines
    #for insert
    qst_lst = [('?') for i in info_lst]
    qst_str = ','.join(qst_lst)

    final= f'''
    INSERT INTO {table_name}
    ({col_str})
    VALUES
    ({qst_str})
    '''
    return final

def df_load(insert,df,session,chunk=False):
    try:
        if chunk==True:
            df_len = len(df.index)
            chunk = ceil(df_len/10000)
            chunks_df = np.array_split(df,chunk)
            for i,_ in enumerate(chunks_df):
                data = [tuple(x) for x in chunks_df[i].to_records(index=False)]
                session.executemany(insert,data,batch=True)
            print('Chunk loading complete.')
        elif chunk==False:
            data = [tuple(x) for x in df.to_records(index=False)]
            session.executemany(insert,data,batch=True)
            print('Load complete.')
        else:
            print('Please set chunk option to True or False')
    except Exception as e:
        print(e)

def params(var_names):
    params={}
    for i in var_names:
        params[i] = globals()[i]
    params_df = pd.DataFrame(params,index=['Parameter'])
    params_df= params_df.T
    params_str= params_df.to_string()
    note = open(f"Parameters-{add_date}.txt","w")
    note.write(params_str)
    note.close()
    
def emailer_head(subject,to,cc='',disp='N'):
    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = to
    message.CC = cc
    message.Subject = subject
    if disp.lower()=='y':
        message.Display()
    return message
    
def emailer_body(body,message,attach=None,disp='N'):
    now=datetime.now()
    year,month,day  = now.year,now.strftime('%m'),now.strftime('%d')
    wd = os.getcwd()
    message.HTMLBody =f"""{body}"""
    if disp.lower()=='y':
        message.Display()
    if attach == None:    
        message.Saveas(f"{wd}//{message.subject}_{year}{month}{day}.msg")
    if attach is not None:
        message.Attachments.Add(fr'{wd}\\{attach}')
        message.Saveas(f"{wd}//{message.subject}_{year}{month}{day}.msg")

def protect_wb(input_wb,output_wb,passw,date_include='y'):
    now=datetime.now()
    year,month,day  = now.year,now.strftime('%m'),now.strftime('%d')
    wd = os.getcwd()
    win = client.gencache.EnsureDispatch("Excel.Application")
    win.DisplayAlerts = False
    wb = win.Workbooks.Open(f"{wd}\\{input_wb}")
    wb.Visible = False
    if date_include.lower() == 'y':
        file =f'{output_wb} - {year}{month}{day}'
        wb.SaveAs(f"{wd}\\{file}",51,passw)
    elif date_include.lower() == 'n':
        file =f'{output_wb}'
        wb.SaveAs(f"{wd}\\{file}",51,passw)
    elif date_include.lower() not in ['y','n']:
        print('date_include arg must be set to y or n')
        
    wb.Close()
    win.Quit()
    return file
