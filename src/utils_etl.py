import os
import pandas as pd
import numpy as np
import datetime
import logging

def new_filename(name: str):
    """Generates a new filename based on the current date and time.
    =======
    params
    name: str. first part of the file's name
     """
    now = datetime.datetime.now()
    return name + '_' + f"{now.strftime('%Y%m%d_%H%M')}.xlsx"


def get_last_two_files(directory: str):
    """ Get the last two files created in a given directory3
    =============
    Params:
    directory:str. Directory with the files of interest
    ===========
    Returns:
    
    last two files path: list. [last file path, second to last path]
    """
    ##list all files in the directory
    files = [os.path.join(directory, f) for f in os.listdir(directory) if os.path.isfile(os.path.join(directory,f))]
    if not files:
        print("No files found in:", directory)
        return []

    ## sort files by creation time (newest first)
    files.sort(key=os.path.getctime, reverse=True)

    #return the last two
    logging.info("Lists used for diagnostics: %s", files[:2])
    #print('Lists to be use:', files[:2])
    return files[:2]



def online_offline_process(folder_connectplus_downloads_path: str, table_all_stores_path: str, final_storage_path: str, N: int):
    """  This function takes the last two lists downloaded from connect + and the fixed table of all stores.
    Returns a table with all the columns in table_all_stores plus five columns. four columns including the status and time
    of last connections from the two list and one final column with a final diagnostic. 
    ==========================
    Params:
    folder_connectplus_downloads_path: str. 
        Path to the folder where the lists downloaded from connection plus are stored.
    table_all_stores_path: str. 
        Path where the fixed list of all stores considered is stored.
    final_storage_path: 
        Path of the directory where the final results will be stored. 
    N: int. 
        Number of hours withouth activity before download to be considered CommLoss
    ========================
    Returns:
    None. 
        Creates the resulting list in the folder final_storage_path.
    """

    ### gets the paths of the last two lists downloaded in the given directory
    ## they will be .xlsx files with the format the connect plues yields 
    last_two_lists = get_last_two_files(folder_connectplus_downloads_path)

    ## second to last list in directory
    df1 = pd.read_excel(last_two_lists[1], engine='openpyxl', header=4) 
    ## last list in directory
    df2 = pd.read_excel(last_two_lists[0], engine='openpyxl', header=4)

    ##### simplifies names of the columns ###
    ### this could be a small function####
    parts1 = df1.columns.str.split('/')
    df1.columns = parts1.str[4] + "/" + parts1.str[5].str.split(':').str[1]
    df1.rename(columns={df1.columns[0]: 'Time'}, inplace=True)
    
    parts2 = df2.columns.str.split('/')
    df2.columns = parts2.str[4] + "/" + parts1.str[5].str.split(':').str[1]
    df2.rename(columns={df2.columns[0]: 'Time'}, inplace=True)

    ####### considering the last N hours observed in each list######
    ### this could also be a small function
    ## N is a given parameter
    
    df1['Time']  = pd.to_datetime(df1['Time'])
    last_time1 = df1['Time'].iloc[-1]
    cutoff1 = last_time1 - pd.Timedelta(hours=N)  ###cutoff time
    last_times_df1 = df1[ (df1['Time'] > cutoff1) & (df1['Time'] <= last_time1)]
    
    df2['Time']  = pd.to_datetime(df2['Time'])
    last_time2 = df2['Time'].iloc[-1]
    cutoff2 = last_time2 - pd.Timedelta(hours=N)  ###cutoff time
    last_times_df2 = df2[ (df2['Time'] > cutoff2) & (df2['Time'] <= last_time2)]

    ##### Online and Offline stores according to criteria (given by N)#######
    ### this could be a small function###

    offline_stores1 = list(last_times_df1.isnull().all()[last_times_df1.isnull().all()].index)
    ###note that it will offline if there are not data at all
    ##it will categorized a store as online if there are some kind of annotation, even if it is not 'On' or 'Off'
    online_stores1 = [element for element in list(df1.columns) if element not in offline_stores1]
    online_stores1.remove('Time')

    offline_stores2 = list(last_times_df2.isnull().all()[last_times_df2.isnull().all()].index)
    online_stores2 = [element for element in list(df2.columns) if element not in offline_stores2]
    online_stores2.remove('Time')


    ########## Adding last time on record (if available on the lists downloaded) for offline stores

    off1 = []
    for item in offline_stores1:
        last_non_null_timestamp = df1.dropna(subset= item).iloc[-1]['Time']
        off1.append([item, last_non_null_timestamp])

    off2 = []
    for item in offline_stores2:
        last_non_null_timestamp = df2.dropna(subset=item).iloc[-1]['Time']
        off2.append([item, last_non_null_timestamp])

    ##### Separating the names into column. so we have [Control System, Unit, status,  last_time_online] for offline stores
    #### and [Control System, Unit, status] for online stores

    offline_list1 = [
        [item[0].split('/')[0].strip(), ##first part: store
        item[0].split('/')[1].strip(), ## second part: unit
        'Offline', ## third part: status
        item[1] ### fourth part: time last connection
    ] for item in off1]

    online_list1 = [
        [item.split('/')[0].strip(), #first part: store
        item.split('/')[1].strip(), #second part: unit
        'Online' ###third part: status     
    ] for item in online_stores1 ]


    offline_list2 = [
        [item[0].split('/')[0].strip(),
        item[0].split('/')[1].strip(),
        'Offline',
        item[1]
    ] for item in off2]

    online_list2 = [
        [item.split('/')[0].strip(),
         item.split('/')[1].strip(),
        'Online'
    ] for item in online_stores2  ]

    ####  creating data frames with the online offline stores####

    offline1_df = pd.DataFrame(offline_list1, columns=['Control System', 'Unit', 'status_1', 'last_time_online_1'])
    offline2_df = pd.DataFrame(offline_list2, columns=['Control System', 'Unit', 'status_2', 'last_time_online_2'])
    online1_df = pd.DataFrame(online_list1, columns=['Control System', 'Unit', 'status_1'])
    online2_df = pd.DataFrame(online_list2, columns=['Control System', 'Unit', 'status_2'])

    #### mergin data frames into a single one with the results of the two processess

    process1 = pd.concat([offline1_df, online1_df], ignore_index=True)
    process2 = pd.concat([offline2_df, online2_df], ignore_index=True)
    full_process = process1.merge(process2, how='outer', on=['Control System', 'Unit'])

    #### Creates a data frame with the fixed list of all available stores########
    #### uses the connect + format, so ingores the first 4 lines of the .xlsx file

    all_stores = pd.read_excel(table_all_stores_path, engine='openpyxl', header=4)

    ###### merge the dataframe of all stores with the dataframe of the full_process

    result = all_stores.merge(full_process, how='left', on=['Control System', 'Unit'])

    #### for stores that do not appear in the downloaded lists, the statuts is automatically Offline

    result['status_1'] = result['status_1'].fillna('Offline')
    result['status_2'] = result['status_2'].fillna('Offline')

    #######################################################################
    ######## logic to do the final diagnostic #####

    ###to make sure they are time stamps
    result['last_time_online_1'] = pd.to_datetime(result['last_time_online_1'])
    result['last_time_online_2'] = pd.to_datetime(result['last_time_online_2'])

    ### conditions
    conditions = [
        #condition 1 (offline_time1 < offline_time2 & status_1 = offline & status_2 = offline): Intermitencia
        (result['last_time_online_1'] < result['last_time_online_2']) & (result['status_1'] == 'Offline') & (result['status_2'] == 'Offline'),
        #condition 2 (offline_time1 == offline_time2 & status_1 = offline & status_2 = offline): CommLoss
        (result['last_time_online_1'] == result['last_time_online_2']) & (result['status_1'] == 'Offline') & (result['status_2'] == 'Offline'),
        #condition 3 (online & Offline): Intermitencia  
        (result['status_1'] == 'Online') & (result['status_2'] == 'Offline'),
        #condition 4 (offline & online): online
        (result['status_1'] == 'Offline') & (result['status_2'] == 'Online'),
        ##condition 5 (offline & offline. no timestamps): CommLoss
        (result['last_time_online_1'].isna() & result['last_time_online_2'].isna()) & (result['status_1'] == 'Offline') & (result['status_2'] == 'Offline'),
        ## condition 6 (online & online) :  Online
        (result['status_1'] == 'Online') & (result['status_2'] == 'Online')
    ]

    diagnostics = [
        'Intermitencia', 
        'CommLoss',
        'Intermitencia', 
        'Online',
        'CommLoss', 
        'Online'
    ]

    result['diagnostico'] = np.select(conditions, diagnostics, default='Error/Unknown')

    ##### saving an excel file with the results in the provided directory for storage
    name_date = new_filename('diagnostics')
    result.to_excel(os.path.join(final_storage_path,name_date))
    logging.info("Succesfully created diagnostics file at: %s", os.path.join(final_storage_path, name_date))
    #print('Succesfully created diagnostics file at:', final_storage_path + name_date)
    return None


    






















































