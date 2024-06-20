from typing import List

import warnings

import pandas as pd
import numpy as np
import openpyxl
import os
from datetime import datetime
from pyxlsb import convert_date

warnings.filterwarnings("ignore", category=FutureWarning)

pd.set_option('display.max_columns', None)

def print_time():
    print("===",datetime.now())

def calculate_distance(group):
    group = group.copy()  # Avoid changing the original DataFrame
    # print("Calculating distance")
    group['Jarak (Kilometer)'] = np.sqrt(
        (group['Long ISR ENoB'] - group['Long ISR ENoB'].shift())**2 +
        (group['Lat ISR ENoB'] - group['Lat ISR ENoB'].shift())**2
    ) * 110.6
    
    # =(SQRT((Y2-AA2)^2+(Z2-AB2)^2)*110.6)*1000
    
    # print("Filling NaN with the next non-null value")
    group['Jarak (Kilometer)'] = group['Jarak (Kilometer)'].fillna(method='bfill')  # Fill NaN with the next non-null value

    return group

def read_csv_with_address_fix(file_path):
    # Read the CSV file while specifying the maximum number of columns
    df = pd.read_csv(file_path, header=None)

    # Define the correct column names
    columns = ['site_id', 'site_name', 'longitude', 'latitude', 'Lat (D)', 'Lat (M)', 'Lat (S)', 'S/N',
               'Long (D)', 'Long (M)', 'Long (S)', 'E/W', 'site_address']

    # Handle cases where site_address expands into multiple columns
    df.columns = columns[:len(df.columns)]
    
    # Combine extra columns back into site_address
    if len(df.columns) > len(columns):
        df['site_address'] = df.iloc[:, len(columns)-1:].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)
        df = df.iloc[:, :len(columns)]

    return df

def safe_convert_date(value):
    try:
        return convert_date(value)
    except (TypeError, ValueError):
        return pd.NaT

def read_csv_with_custom_delimiter(file_path):
    # Define the columns
    columns = ['site_id', 'site_name', 'longitude', 'latitude', 'Lat (D)', 'Lat (M)', 'Lat (S)', 'S/N',
               'Long (D)', 'Long (M)', 'Long (S)', 'E/W', 'site_address']

    data = []

    # Read the file
    with open(file_path, 'r', encoding='cp1252') as file:
        lines = file.readlines()

        for line in lines:
            # Skip lines that start with any of the column names (assuming they are headers)
            if line.startswith(tuple(columns)):
                continue
            
            # Split the line by commas only up to the 12th element (index 11 for 'E/W' column)
            parts = line.strip().split(',', 12)

            # If the line doesn't have enough parts, it's malformed and should be skipped
            if len(parts) < 13:
                continue

            # Combine any extra parts into the 'site_address' column
            if len(parts) > 13:
                parts[12] = ','.join(parts[12:])
                parts = parts[:13]

            data.append(parts)

    # Create DataFrame
    df = pd.DataFrame(data, columns=columns)

    return df

def getComplyUncomply(input_folder: List[str]):
    print("\n========= START")
    print_time()
    # Check if folder is empty or not
    if not input_folder:
        print("Error: Input folder list is empty.")
        return "Error: Input folder list is empty."
    
    file_paths = input_folder
    
    # ================================================
    # Input data NPE
    # ================================================
    
    print("\n=== Reading NPE data...")
    print_time()
    df_NPE = pd.DataFrame()

    for file_path in file_paths:
        # Check if the file name starts with "NPE"
        if os.path.basename(file_path).startswith("NPE"):
            try:
                # Read Excel file using pyxlsb engine
                # self.data: pd.DataFrame = pd.read_excel(self.file, sheet_name=self.sheet, engine='pyxlsb', header=0)
                # self.data["test"] = self.data.apply(lambda x: convert_date(x.SomeStupidDate), axis=1)
                df_temp = pd.read_excel(file_path, sheet_name='Detail ISR', header=0, engine='pyxlsb', parse_dates=['ISR Expired Date'])
                df_NPE = pd.concat([df_NPE, df_temp], ignore_index=True)
                print(f"Successfully read data from file: {file_path}")
                # Perform further operations with df_NPE if needed
            except Exception as e:
                print(f"Error reading file '{file_path}': {str(e)}")
        else:
            print(f"Skipping file: {file_path} (does not start with 'NPE')")
            
    # =====================================================================================
    # df_ENoB
    # =====================================================================================
    
    df_ENoB = pd.DataFrame()
    
    print("\n=== Reading ENoB data...")
    print_time()

    for file_path in file_paths:
        if os.path.basename(file_path).startswith("ENoB"):
            try:
                # Read Excel file using pyxlsb engine
                df_temp = read_csv_with_custom_delimiter(file_path)
                df_ENoB = pd.concat([df_ENoB, df_temp], ignore_index=True)
                print(f"Successfully read data from file: {file_path}")
                # Perform further operations with df_NPE if needed
            except Exception as e:
                print(f"Error reading file '{file_path}': {str(e)}")
        else:
            print(f"Skipping file: {file_path} (does not start with 'NPE')")
            
    # ================================================
    # Input data Modif
    # ================================================
    
    print("\n=== Reading Mod data...")
    print_time()
    df_Mod = pd.DataFrame()

    for file_path in file_paths:
        # Check if the file name starts with "NPE"
        if os.path.basename(file_path).startswith("Modif"):
            try:
                # Read Excel file using pyxlsb engine
                # self.data: pd.DataFrame = pd.read_excel(self.file, sheet_name=self.sheet, engine='pyxlsb', header=0)
                # self.data["test"] = self.data.apply(lambda x: convert_date(x.SomeStupidDate), axis=1)
                df_temp = pd.read_excel(file_path, sheet_name='Detail ISR', header=0, engine='pyxlsb')
                df_Mod = pd.concat([df_Mod, df_temp], ignore_index=True)
                print(f"Successfully read data from file: {file_path}")
                # Perform further operations with df_NPE if needed
            except Exception as e:
                print(f"Error reading file '{file_path}': {str(e)}")
        else:
            print(f"Skipping file: {file_path} (does not start with 'Modif')")

    # =====================================================================================
    # df_Postel
    # =====================================================================================
    
    df_postel = pd.DataFrame()
    
    print("\n=== Reading Postel data...")
    print_time()

    for file_path in file_paths:
        # Check if the file name starts with "NPE"
        if os.path.basename(file_path).startswith("Postel"):
            try:
                # Read Excel file using pyxlsb engine
                df_temp = pd.read_excel(file_path, header=0)
                df_postel = pd.concat([df_postel, df_temp], ignore_index=True)
                print(f"Successfully read data from file: {file_path}")
                # Perform further operations with df_NPE if needed
            except Exception as e:
                print(f"Error reading file '{file_path}': {str(e)}")
        else:
            print(f"Skipping file: {file_path} (does not start with 'NPE')")

    # ================================================
    # df_result
    # ================================================
    
    print("\n=== Creating result dataframe...")
    print_time()

    columns_to_select = ['No. RR#site ID', 
                'APPL_ID/site ID',
                'No. RR',
                'REF ID (NE)', 
                'REF ID (FE)',
                'New Link ID',
                'STN_NAME', 
                'Height ASL (m)',
                'Long ISR', 
                'Lat ISR',
                'Region', 
                'PROVINCE', 
                'Site_District',
                'Site_City', 
                'STN_ADDR', 
                'Radio Type', 
                'EQ_Code', 
                'Output_Power_W',
                'BWIDTH (MHz)', 
                'Tx', 
                'Rx',
                'Antenna Type', 
                'ANT_Code', 
                'HGT_ANT',
                'Pol',
                'PM2 NPE',
                'ISR Expired Date', 
                'Month',
                'Freq', 
                'Radio',
                'Zona', 
                'Site ID (MDRS)', 
                'Plus Code',
                'Status ISR']

    alternate_names = {
        'PM2 NPE': ['PM2'],
        # Add alternate names for other columns as needed
    }

    df_result = pd.DataFrame()

    for col in columns_to_select:
        try:
            # Attempt to select the specified column from df_NPE
            df_result[col] = df_NPE[col]
        except KeyError as e:
            # Handle the KeyError if the column doesn't exist in df_NPE
            original_column_name = e.args[0]
            
            # Check alternate names if the original column doesn't exist
            found_alternate = False
            for alt_name in alternate_names.get(original_column_name, []):
                if alt_name in df_NPE.columns:
                    df_result[original_column_name] = df_NPE[alt_name]
                    print(f"Successfully selected column '{original_column_name}' using alternate name '{alt_name}'.")
                    found_alternate = True
                    break
            
            if not found_alternate:
                print(f"Error: Column '{original_column_name}' and its alternates do not exist in the DataFrame.")
    
    # Splitting APPL_ID/site ID
    df_result[['Output_Aplikasi','Output_Station_ID']] = df_result['APPL_ID/site ID'].str.split('/',expand=True)
    
    print("\n=== Drop column where Status ISR is Not Active#ISR Terminated or ISR Not Active#10 years...")
    print_time()
    values_to_drop = ["ISR Not Active#ISR Terminated", "ISR Not Active#10 years"]
    df_result = df_result[~df_result['Status ISR'].isin(values_to_drop)]

    print("\n=== Converting ISR Expired Date...")
    print_time()
    # Converting date
    
    df_result['ISR Expired Date (old)'] = df_result['ISR Expired Date']
    
    # =====================================================================================
    # df_postel
    # =====================================================================================
    
    print("\n=== Creating unique_id...")
    print_time()
    df_postel.info()
    df_postel['SITE_ID'] = df_postel['SITE_ID'].astype(str).str.lstrip('0')
    df_postel['unique_id'] = df_postel['NO_RR_SIMS'].astype(str) + '#' + df_postel['SITE_ID'].astype(str)
    df_result['unique_id'] = df_result['No. RR#site ID'].astype(str)
    df_postel.info()
    
    # Input value from Postel data
    print("\n=== Input value from Postel data...")
    print_time()
    
    df_result['Long ISR Postel'] = df_result['unique_id'].map(df_postel.set_index('unique_id')['SID_LONG'])
    df_result['Lat ISR Postel'] = df_result['unique_id'].map(df_postel.set_index('unique_id')['SID_LAT'])

    # df_result[['Output_Aplikasi', 'Long ISR', 'Lat ISR', 'Long ISR Postel', 'Lat ISR Postel']].head()
    
    # If not null, fill with B (Postel) column
    df_result.loc[df_result['Long ISR Postel'].notna(), 'Output 1 Long'] = df_result['Long ISR Postel']
    df_result.loc[df_result['Lat ISR Postel'].notna(), 'Output 1 Lat'] = df_result['Lat ISR Postel']

    # If null, fill with A (NPE) column
    df_result.loc[df_result['Long ISR Postel'].isna(), 'Output 1 Long'] = df_result['Long ISR']
    df_result.loc[df_result['Lat ISR Postel'].isna(), 'Output 1 Lat'] = df_result['Lat ISR']
    
    # =========================
    # df_result.loc[df_result['Long ISR Postel'].isna(), 'Output 1 Long'] = 12345
    # df_result.loc[df_result['Lat ISR Postel'].isna(), 'Output 1 Lat'] = 12345
    # =========================

    df_ENoB['longitude'] = df_ENoB['longitude'].astype(float)
    df_ENoB['latitude'] = df_ENoB['latitude'].astype(float)
        
    # Input value from ENoB data
    df_result['Long ISR ENoB'] = None
    df_result['Long ISR ENoB'] = df_result['REF ID (NE)'].map(df_ENoB.set_index('site_id')['longitude'])
    
    df_result['Lat ISR ENoB'] = None
    df_result['Lat ISR ENoB'] = df_result['REF ID (NE)'].map(df_ENoB.set_index('site_id')['latitude'])

    print("\n=== Get output 2 long lat...")
    print_time()
    
    # removing duplicates
    duplicates = df_Mod[df_Mod.duplicated(subset=['APPL_ID/site ID'], keep=False)]
    print(f"Duplicate entries in 'APPL_ID/site ID':\n{len(duplicates)}")
    df_Mod = df_Mod.drop_duplicates(subset=['APPL_ID/site ID'])
    
    # Proceed with the mapping after ensuring 'APPL_ID/site ID' is unique
    df_result = df_result.reset_index(drop=True)
    df_Mod = df_Mod.reset_index(drop=True)
    
    df_result['Long ENodeB'] = df_result['APPL_ID/site ID'].map(df_Mod.set_index('APPL_ID/site ID')['Long ENodeB'])
    df_result['Lat EnodeB'] = df_result['APPL_ID/site ID'].map(df_Mod.set_index('APPL_ID/site ID')['Lat EnodeB'])
    
    # If not null, fill with B (Output 2) column
    df_result.loc[df_result['Long ISR ENoB'].notna(), 'Output 2 Long'] = df_result['Long ISR ENoB']
    df_result.loc[df_result['Lat ISR ENoB'].notna(), 'Output 2 Lat'] = df_result['Lat ISR ENoB']
    
    # If null
    df_result.loc[(df_result['Long ISR ENoB'].isnull()) | (df_result['Long ISR ENoB'] == 0), 'Output 2 Long'] = df_result['Long ENodeB']
    df_result.loc[(df_result['Lat ISR ENoB'].isnull()) | (df_result['Lat ISR ENoB'] == 0), 'Output 2 Lat'] = df_result['Lat EnodeB']
    
    # Last chance
    df_result.loc[(df_result['Output 2 Long'].isnull()) | (df_result['Output 2 Long'] == 0), 'Output 2 Long'] = df_result['Output 1 Long']
    df_result.loc[(df_result['Output 2 Lat'].isnull()) | (df_result['Output 2 Lat'] == 0), 'Output 2 Lat'] = df_result['Output 1 Lat']
    
    # df_result.loc[df_result['Long ISR ENoB'].isna(), 'Long ISR ENoB'] = df_result['Output 1 Long']
    # df_result.loc[df_result['Lat ISR ENoB'].isna(), 'Lat ISR ENoB'] = df_result['Output 1 Lat']
    
    # Jarak Geser (Meter)
    print("\n=== Calculate distance difference (Meter)...")
    print_time()
    df_result = df_result.reset_index(drop=True)
    df_result['Jarak Geser (Meter)'] = np.sqrt(
        (df_result['Output 1 Long'] - df_result['Output 2 Long'])**2 + 
        (df_result['Output 1 Lat'] - df_result['Output 2 Lat'])**2
    ) * 110.6 * 1000
    
    print("\n=== Calculating site distance...")
    print_time()
    df_result = df_result.reset_index(drop=True)
    # df_result = df_result.groupby(['Output_Aplikasi', 'No. RR']).apply(calculate_distance).reset_index(drop=True)
    
    df_result['Jarak (Kilometer)'] = np.sqrt(
        (df_result['Output 2 Long'] - df_result.groupby(['Output_Aplikasi', 'No. RR'])['Output 2 Long'].shift())**2 +
        (df_result['Output 2 Lat'] - df_result.groupby(['Output_Aplikasi', 'No. RR'])['Output 2 Lat'].shift())**2
    ) * 110.6

    # Fill NaN values with the next non-null value within each group
    df_result['Jarak (Kilometer)'] = df_result.groupby(['Output_Aplikasi', 'No. RR'])['Jarak (Kilometer)'].fillna(method='bfill')

    # Reset index if needed
    df_result = df_result.reset_index(drop=True)  
    
    # Apply Com0
    print("\n=== Determining compliance status...")
    print_time()
    conditions = [
        df_result['Jarak Geser (Meter)'] <= 20,
        (df_result['Jarak Geser (Meter)'] > 20) & (df_result['Jarak Geser (Meter)'] <= 50),
        (df_result['Jarak Geser (Meter)'] > 50) & (df_result['Jarak Geser (Meter)'] <= 100),
        df_result['Jarak Geser (Meter)'] > 100
    ]

    choices = ['OK (Match)', 'P2', 'P1', 'P0']

    # Apply the conditions using numpy.select
    df_result['Com0'] = np.select(conditions, choices, default='')
    
    # Sort values
    df_result = df_result.sort_values(by=['Output_Aplikasi', 'Output_Station_ID'])
    
    # Com1
    df_result['Com1'] = df_result.groupby(['Output_Aplikasi', 'No. RR'])['Com0'].transform(lambda x: ''.join(x))

    """
    =IF(OR(AK2="-",CB2="00G",CB2="05G",CB2="18G"),"TBC",
    IF(AND(CB2="32G",AK2>0),"Comply",
    IF(AND(CB2="23G",AK2>0.2),"Comply",
    IF(AND(CB2="11G",AK2>2.5),"Comply",
    IF(AND(CB2="13G",AK2>2.5),"Comply",
    IF(AND(CB2="15G",AK2>2.5),"Comply",
    IF(AND(CB2="07G",AK2>8),"Comply",
    IF(AND(CB2="08G",AK2>8),"Comply",
    IF(AND(CB2="04G",AK2>20),"Comply",
    IF(AND(CB2="06G",AK2>20),"Comply",
    "Un-Comply"))))))))))
    """

    conditions = [
        (df_result['Jarak (Kilometer)'] == '-') | (df_result['Freq'].isin(['00G', '05G', '18G'])),
        (df_result['Freq'] == '32G') & (df_result['Jarak (Kilometer)'].astype(str).astype(float) >= 0),
        (df_result['Freq'] == '23G') & (df_result['Jarak (Kilometer)'].astype(str).astype(float) > 0.2),
        (df_result['Freq'] == '11G') & (df_result['Jarak (Kilometer)'].astype(str).astype(float) > 2.5),
        (df_result['Freq'] == '13G') & (df_result['Jarak (Kilometer)'].astype(str).astype(float) > 2.5),
        (df_result['Freq'] == '15G') & (df_result['Jarak (Kilometer)'].astype(str).astype(float) > 2.5),
        (df_result['Freq'] == '07G') & (df_result['Jarak (Kilometer)'].astype(str).astype(float) > 8),
        (df_result['Freq'] == '08G') & (df_result['Jarak (Kilometer)'].astype(str).astype(float) > 8),
        (df_result['Freq'] == '04G') & (df_result['Jarak (Kilometer)'].astype(str).astype(float) > 20),
        (df_result['Freq'] == '06G') & (df_result['Jarak (Kilometer)'].astype(str).astype(float) > 20)
    ]

    # Define choices based on the given formula
    choices = [
        'TBC',
        'Comply',
        'Comply',
        'Comply',
        'Comply',
        'Comply',
        'Comply',
        'Comply',
        'Comply',
        'Comply'
    ]

    # Use numpy.select to apply the conditions and choices
    df_result['Compliance_Status'] = np.select(conditions, choices, default='Un-Comply')
    
    # print("\n=== Getting df_result info...")
    # print(df_result.info())
    
    # get today - due date
    print("\n=== Calculating today - due date...")
    print_time()
    df_result['ISR Expired Date (postel)'] = pd.to_datetime(df_result['unique_id'].map(df_postel.set_index('unique_id')['MASA_LAKU_BHP']), 
                                                            format='%d-%b-%y').dt.strftime('%Y/%m/%d')
    df_result['ISR Expired Date (postel)'] = pd.to_datetime(df_result['ISR Expired Date (postel)'], 
                                                            format='%Y/%m/%d').dt.date
    
    df_result['ISR Expired Date (new)'] = pd.to_datetime(df_result['ISR Expired Date (old)'], origin='1899-12-30', unit='D').dt.strftime('%Y/%m/%d')
    df_result['ISR Expired Date (new)'] = pd.to_datetime(df_result['ISR Expired Date (new)'],
                                                         format='%Y/%m/%d').dt.date
    
    df_result['ISR Expired Date (core)'] = df_result['ISR Expired Date (postel)'].fillna(df_result['ISR Expired Date (new)'])
    
    df_result['Today date'] =  pd.to_datetime(pd.Timestamp.now().normalize().strftime('%Y/%m/%d'),
                                              format='%Y/%m/%d')
    df_result['Today date'] = pd.to_datetime(df_result['Today date'],
                                             format='%Y/%m/%d').dt.date
    
    df_result['Today - due date'] = (pd.to_datetime(df_result['Today date']) - pd.to_datetime(df_result['ISR Expired Date (core)'])).dt.days
    
    print(df_result.info())
    # df_result['Today - due date'] = (pd.Timestamp.now() - pd.to_datetime(df_result['ISR Expired Date'])).dt.days
    # df_result['Today - due date'] = (datetime.today().date() - pd.to_datetime(df_result['ISR Expired Date']).dt.date).dt.days * -1
    
    print("\n=== Get Modif/no")
    print_time()
    # conditions = [
    #     df_result['Today - due date'] > 10,
    #     df_result['Today - due date'] < -91
    # ]

    # choices = [
    #     'Bisa Modif',
    #     'Bisa Modif'
    # ]

    # df_result['Modif/no'] = np.select(conditions, choices, default='Tidak Bisa Dimodif')
    
    # Initialize the column with the default value
    df_result['Modif/no'] = 'Tidak Bisa Dimodif'

    # Apply conditions using boolean indexing
    df_result.loc[df_result['Today - due date'] > 10, 'Modif/no'] = 'Bisa Modif'
    df_result.loc[df_result['Today - due date'] < -91, 'Modif/no'] = 'Bisa Modif'
    df_result.loc[df_result['Today - due date'].isna(), 'Modif/no'] = 'Check by aplikasi'
    
    # =====================================================================================
    # Finishing
    # =====================================================================================
    
    first_file_path = file_paths[0]
    directory = os.path.dirname(first_file_path)
    current_datetime = datetime.now().strftime("%Y%m%d_%H%M")

    output_file_path = os.path.join(directory, f"Result_UComply_{current_datetime}.xlsx")
    df_result.to_excel(output_file_path, index=False)
    print(output_file_path)
    
    df_postel.to_csv(os.path.join(directory, f"Result_UComply_Postel_{current_datetime}.csv"), index=False)
    
    print("\n========= END")
    print_time()
    return output_file_path