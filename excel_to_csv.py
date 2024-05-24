import pandas as pd
from datetime import datetime
import sys

# Suppress all warnings
import warnings
warnings.filterwarnings("ignore")
    
def process_file(df):
    # Create Index column
    df['Index'] = df.index + 1
    
    # Split file to odd and even
    df_odd = df[df['Index'] % 2 != 0]
    df_even = df[df['Index'] % 2 == 0]  
    
    # Drop column
    df_even.drop(columns=['b'], inplace=True) # genap
    df_odd.drop(columns=['b'], inplace=True) # ganjil
    
    # Duplicate column
    column_to_duplicate = ['Antenna_Type',
                       'Equipment_type',
                       'Antenna_Height_AGL_m',
                       'Transmitter_Polarization',
                       'Feeding_Loss_dB',
                       'Connector_Loss_dB',
                       'Branch_Loss_dB']
    
    duplicate_column(df=df_odd, list=column_to_duplicate)
    duplicate_column(df=df_even, list=column_to_duplicate)
    
    # Rearrange
    df_odd = df_odd[["Service_ID", 
                    "SubService_ID", 
                    "RR", 
                    "Reference name", 
                    "App ID", 
                    "Action Type", 
                    "Client_ID", 
                    "Station_Name", 
                    "Height_ASL_m", 
                    "LONG_DEG", "LONG_DIR_IND_E_W", "LONG_MIN", "LONG_SEC", 
                    "LAT_DEG", "LAT_DIR_IND_N_S", "LAT_MIN", "LAT_SEC", 
                    "Area_of_Service", 
                    "Site_District", 
                    "Site_City", 
                    "Site_Province", 
                    "Antenna_Type", 
                    "Equipment_type", 
                    "Equipment_Output_Power_W", 
                    "Antenna_Height_AGL_m", 
                    "Transmitter_Polarization", 
                    "Feeding_Loss_dB", 
                    "Connector_Loss_dB", 
                    "Branch_Loss_dB", 
                    "Transmitter_Freq_MHz", 
                    "Bandwidth_Hz", 
                    "Antenna_Type - Copy", 
                    "Equipment_type - Copy", 
                    "Antenna_Height_AGL_m - Copy", 
                    "Transmitter_Polarization - Copy", 
                    "Feeding_Loss_dB - Copy", 
                    "Connector_Loss_dB - Copy", 
                    "Branch_Loss_dB - Copy", 
                    "Receiver_Freq_MHz", 
                    "Bandwidth_Hz.1"]]
    df_even = df_even[["Service_ID", 
                        "SubService_ID", 
                        "RR", 
                        "Reference name", 
                        "App ID", 
                        "Action Type", 
                        "Client_ID", 
                        "Station_Name", 
                        "Height_ASL_m", 
                        "LONG_DEG", 
                        "LONG_DIR_IND_E_W", 
                        "LONG_MIN", 
                        "LONG_SEC", 
                        "LAT_DEG", 
                        "LAT_DIR_IND_N_S", 
                        "LAT_MIN", 
                        "LAT_SEC", 
                        "Area_of_Service", 
                        "Site_District", 
                        "Site_City", 
                        "Site_Province", 
                        "Antenna_Type", 
                        "Equipment_type", 
                        "Equipment_Output_Power_W", 
                        "Antenna_Height_AGL_m", 
                        "Transmitter_Polarization", 
                        "Feeding_Loss_dB", 
                        "Connector_Loss_dB", 
                        "Branch_Loss_dB", 
                        "Transmitter_Freq_MHz", 
                        "Bandwidth_Hz", 
                        "Antenna_Type - Copy", 
                        "Equipment_type - Copy", 
                        "Antenna_Height_AGL_m - Copy", 
                        "Transmitter_Polarization - Copy", 
                        "Feeding_Loss_dB - Copy", 
                        "Connector_Loss_dB - Copy", 
                        "Branch_Loss_dB - Copy", 
                        "Receiver_Freq_MHz", 
                        "Bandwidth_Hz.1"]]
    
    # Rename columns
    df_odd.rename(columns={"Service_ID": "<SV_SV_ID>", 
                            "SubService_ID": "<SS_SS_ID>", 
                            "RR": "<AP_NAME>", 
                            "Reference name": "<AP_NAME_REFERENCE>", 
                            "App ID": "<AP_REF_NUMBER>", 
                            "Action Type": "<AP_ACTION_TYPE>", 
                            "Client_ID": "<AD_MAN_NUMBER>", 
                            "Station_Name": "<TCS_NAME>", 
                            "Height_ASL_m": "<SID_H_NN>", 
                            "LONG_DEG": "<SID_LONG_DEG>", 
                            "LONG_DIR_IND_E_W": "<SID_LONG_E_W>", 
                            "LONG_MIN": "<SID_LONG_MIN>", 
                            "LONG_SEC": "<SID_LONG_SEC>", 
                            "LAT_DEG": "<SID_LAT_DEG>", 
                            "LAT_DIR_IND_N_S": "<SID_LAT_N_S>", 
                            "LAT_MIN": "<SID_LAT_MIN>", 
                            "LAT_SEC": "<SID_LAT_SEC>", 
                            "Area_of_Service": "<AD_STREET>", 
                            "Site_District": "<AD_CITY>", 
                            "Site_City": "<AD_DISTRICT>", 
                            "Site_Province": "<AD_COUNTY>", 
                            "Antenna_Type": "<TRANSMITTER   EAN_ANT_IDENT=", 
                            "Equipment_type": "<EQP_EQUIP_IDENT>", 
                            "Equipment_Output_Power_W": "<ETX_EQ_OUTPUT>", 
                            "Antenna_Height_AGL_m": "<EAC_AN_H>", 
                            "Transmitter_Polarization": "<EAC_AN_POL>", 
                            "Feeding_Loss_dB": "<EAC_FEEDING_LOSS>", 
                            "Connector_Loss_dB": "<EAC_CONNECTOR_LOSS>", 
                            "Branch_Loss_dB": "<EAC_BRANCH_LOSS>", 
                            "Transmitter_Freq_MHz": "<EFL_FREQ>x", 
                            "Bandwidth_Hz": "<EFL_RF_BWIDTH>x", 
                            "Antenna_Type - Copy": "<RECEIVER   EAN_ANT_IDENT=", 
                            "Equipment_type - Copy": "<EQP_EQUIP_IDENT2>", 
                            "Antenna_Height_AGL_m - Copy": "<EAC_AN_H2>", 
                            "Transmitter_Polarization - Copy": "<EAC_AN_POL2>", 
                            "Feeding_Loss_dB - Copy": "<EAC_FEEDING_LOSS2>", 
                            "Connector_Loss_dB - Copy": "<EAC_CONNECTOR_LOSS2>", 
                            "Branch_Loss_dB - Copy": "<EAC_BRANCH_LOSS2>", 
                            "Receiver_Freq_MHz": "<EFL_FREQ2>", 
                            "Bandwidth_Hz.1": "<EFL_RF_BWIDTH2>"}, 
                  inplace=True)
    df_even.rename(columns={"Station_Name": "<TCS_NAME2>", 
                        "Height_ASL_m": "<SID_H_NN2>", 
                        "LONG_DEG": "<SID_LONG_DEG2>", 
                        "LONG_DIR_IND_E_W": "<SID_LONG_E_W2>", 
                        "LONG_MIN": "<SID_LONG_MIN2>", 
                        "LONG_SEC": "<SID_LONG_SEC2>", 
                        "LAT_DEG": "<SID_LAT_DEG2>", 
                        "LAT_DIR_IND_N_S": "<SID_LAT_N_S2>", 
                        "LAT_MIN": "<SID_LAT_MIN2>", 
                        "LAT_SEC": "<SID_LAT_SEC2>", 
                        "Area_of_Service": "<AD_STREET2>", 
                        "Site_District": "<AD_CITY2>", 
                        "Site_City": "<AD_DISTRICT2>", 
                        "Site_Province": "<AD_COUNTY2>", 
                        "Antenna_Type": "<TRANSMITTER   EAN_ANT_IDENT2=", 
                        "Equipment_type": "<EQP_EQUIP_IDENT3>", 
                        "Equipment_Output_Power_W": "<ETX_EQ_OUTPUT4>", 
                        "Antenna_Height_AGL_m": "<EAC_AN_H4>", 
                        "Transmitter_Polarization": "<EAC_AN_POL4>", 
                        "Feeding_Loss_dB": "<EAC_FEEDING_LOSS4>", 
                        "Connector_Loss_dB": "<EAC_CONNECTOR_LOSS4>", 
                        "Branch_Loss_dB": "<EAC_BRANCH_LOSS4>", 
                        "Transmitter_Freq_MHz": "<EFL_FREQ3>a", 
                        "Bandwidth_Hz": "<EFL_RF_BWIDTH3>b", 
                        "Antenna_Type - Copy": "<RECEIVER   EAN_ANT_IDENT2=", 
                        "Equipment_type - Copy": "<EQP_EQUIP_IDENT4>", 
                        "Antenna_Height_AGL_m - Copy": "<EAC_AN_H5>", 
                        "Transmitter_Polarization - Copy": "<EAC_AN_POL5>", 
                        "Feeding_Loss_dB - Copy": "<EAC_FEEDING_LOSS5>", 
                        "Connector_Loss_dB - Copy": "<EAC_CONNECTOR_LOSS5>", 
                        "Branch_Loss_dB - Copy": "<EAC_BRANCH_LOSS5>", 
                        "Receiver_Freq_MHz": "<EFL_FREQ4>", 
                        "Bandwidth_Hz.1": "<EFL_RF_BWIDTH4>"}, 
              inplace=True)
    
    # Add custom
    df_odd["<EFL_FREQ>"] = df_odd["<EFL_FREQ>x"]
    df_odd['<EFL_RF_BWIDTH>'] = df_odd['<EFL_RF_BWIDTH>x']

    df_even["<EFL_FREQ3>"] = df_even["<EFL_FREQ3>a"]
    df_even['<EFL_RF_BWIDTH3>'] = df_even['<EFL_RF_BWIDTH3>b']
    
    # Reorder column
    df_odd = df_odd[["<SV_SV_ID>", "<SS_SS_ID>", "<AP_NAME>", "<AP_NAME_REFERENCE>", "<AP_REF_NUMBER>", 
                 "<AP_ACTION_TYPE>", "<AD_MAN_NUMBER>", "<TCS_NAME>", "<SID_H_NN>", "<SID_LONG_DEG>", 
                 "<SID_LONG_E_W>", "<SID_LONG_MIN>", "<SID_LONG_SEC>", "<SID_LAT_DEG>", "<SID_LAT_N_S>", 
                 "<SID_LAT_MIN>", "<SID_LAT_SEC>", "<AD_STREET>", "<AD_CITY>", "<AD_DISTRICT>", "<AD_COUNTY>", 
                 "<TRANSMITTER   EAN_ANT_IDENT=", "<EQP_EQUIP_IDENT>", "<ETX_EQ_OUTPUT>", "<EAC_AN_H>", 
                 "<EAC_AN_POL>", "<EAC_FEEDING_LOSS>", "<EAC_CONNECTOR_LOSS>", "<EAC_BRANCH_LOSS>", 
                 "<EFL_FREQ>", "<EFL_RF_BWIDTH>", "<EFL_FREQ>x", "<EFL_RF_BWIDTH>x", "<RECEIVER   EAN_ANT_IDENT=", 
                 "<EQP_EQUIP_IDENT2>", "<EAC_AN_H2>", "<EAC_AN_POL2>", "<EAC_FEEDING_LOSS2>", 
                 "<EAC_CONNECTOR_LOSS2>", "<EAC_BRANCH_LOSS2>", "<EFL_FREQ2>", "<EFL_RF_BWIDTH2>"]]

    # Remove/drop column
    df_odd.drop(columns=["<EFL_FREQ>x", "<EFL_RF_BWIDTH>x"], inplace=True) # ganjil
    df_even.drop(columns=["Service_ID", "SubService_ID", "Reference name", "App ID", "Action Type", "Client_ID"], 
             inplace=True)
    
    # Change type
    df_odd["<AP_REF_NUMBER>"] = df_odd["<AP_REF_NUMBER>"].astype(str)
    
    # -----------------------------------------------------------------------------
    
    # Merge data - outer left join
    temp_df_even = df_even[["RR", "<TCS_NAME2>", "<SID_H_NN2>", "<SID_LONG_DEG2>", "<SID_LONG_E_W2>", "<SID_LONG_MIN2>", "<SID_LONG_SEC2>", 
                        "<SID_LAT_DEG2>", "<SID_LAT_N_S2>", "<SID_LAT_MIN2>", "<SID_LAT_SEC2>", "<AD_STREET2>", "<AD_CITY2>", 
                        "<AD_DISTRICT2>", "<AD_COUNTY2>", "<TRANSMITTER   EAN_ANT_IDENT2=", "<EQP_EQUIP_IDENT3>", "<ETX_EQ_OUTPUT4>", 
                        "<EAC_AN_H4>", "<EAC_AN_POL4>", "<EAC_FEEDING_LOSS4>", "<EAC_CONNECTOR_LOSS4>", "<EAC_BRANCH_LOSS4>", 
                        "<EFL_FREQ3>", "<EFL_RF_BWIDTH3>", "<RECEIVER   EAN_ANT_IDENT2=", "<EQP_EQUIP_IDENT4>", "<EAC_AN_H5>", 
                        "<EAC_AN_POL5>", "<EAC_FEEDING_LOSS5>", "<EAC_CONNECTOR_LOSS5>", "<EAC_BRANCH_LOSS5>", "<EFL_FREQ4>", 
                        "<EFL_RF_BWIDTH4>"]]
    merged = pd.merge(df_odd, temp_df_even, 
                    left_on='<AP_NAME>', right_on='RR', 
                    how = 'left')
    merged.drop(columns="RR", inplace=True)
    
    # Rename
    merged.rename(columns={"<AD_DISTRICT>": "<AD_CITY>", 
                       "<AD_CITY>": "<AD_DISTRICT>", 
                       "<AD_CITY2>": "<AD_DISTRICT2>", 
                       "<AD_DISTRICT2>": "<AD_CITY2>"},
              inplace=True)
    
    # Reorder
    merged = merged[["<SV_SV_ID>", "<SS_SS_ID>", "<AP_NAME>", "<AP_NAME_REFERENCE>", "<AP_REF_NUMBER>", 
                 "<AP_ACTION_TYPE>", "<AD_MAN_NUMBER>", "<TCS_NAME>", "<SID_H_NN>", "<SID_LONG_DEG>", 
                 "<SID_LONG_E_W>", "<SID_LONG_MIN>", "<SID_LONG_SEC>", "<SID_LAT_DEG>", "<SID_LAT_N_S>", 
                 "<SID_LAT_MIN>", "<SID_LAT_SEC>", "<AD_STREET>", "<AD_CITY>", "<AD_DISTRICT>", "<AD_COUNTY>", 
                 "<TRANSMITTER   EAN_ANT_IDENT=", "<EQP_EQUIP_IDENT>", "<ETX_EQ_OUTPUT>", "<EAC_AN_H>", 
                 "<EAC_AN_POL>", "<EAC_FEEDING_LOSS>", "<EAC_CONNECTOR_LOSS>", "<EAC_BRANCH_LOSS>", "<EFL_FREQ>", 
                 "<EFL_RF_BWIDTH>", "<RECEIVER   EAN_ANT_IDENT=", "<EQP_EQUIP_IDENT2>", "<EAC_AN_H2>", 
                 "<EAC_AN_POL2>", "<EAC_FEEDING_LOSS2>", "<EAC_CONNECTOR_LOSS2>", "<EAC_BRANCH_LOSS2>", 
                 "<EFL_FREQ2>", "<EFL_RF_BWIDTH2>", "<TCS_NAME2>", "<SID_H_NN2>", "<SID_LONG_DEG2>", 
                 "<SID_LONG_E_W2>", "<SID_LONG_MIN2>", "<SID_LONG_SEC2>", "<SID_LAT_DEG2>", "<SID_LAT_N_S2>", 
                 "<SID_LAT_MIN2>", "<SID_LAT_SEC2>", "<AD_STREET2>", "<AD_CITY2>", "<AD_DISTRICT2>", 
                 "<AD_COUNTY2>", "<TRANSMITTER   EAN_ANT_IDENT2=", "<EQP_EQUIP_IDENT3>", "<ETX_EQ_OUTPUT4>", 
                 "<EAC_AN_H4>", "<EAC_AN_POL4>", "<EAC_FEEDING_LOSS4>", "<EAC_CONNECTOR_LOSS4>", "<EAC_BRANCH_LOSS4>", 
                 "<EFL_FREQ3>", "<EFL_RF_BWIDTH3>", "<RECEIVER   EAN_ANT_IDENT2=", "<EQP_EQUIP_IDENT4>", "<EAC_AN_H5>", 
                 "<EAC_AN_POL5>", "<EAC_FEEDING_LOSS5>", "<EAC_CONNECTOR_LOSS5>", "<EAC_BRANCH_LOSS5>", "<EFL_FREQ4>", 
                 "<EFL_RF_BWIDTH4>"]]
    
    # Add index column
    merged['Index'] = merged.index + 1

    # Add 'Program' column
    merged['Program'] = merged.apply(lambda row: 1 
                                 if 1 <= row.Index <= 1000 else 2, 
                                 axis=1)
    
    # Drop index column
    merged.drop(columns=['Index'], inplace=True)
    
    # Change column type
    merged["<AP_NAME>"] = merged["<AP_NAME>"].astype(str)
    merged['UniqueID'] = merged['<AP_NAME>'].astype(str) + '#' + merged['<AP_REF_NUMBER>'].astype(str)
    
    # Result
    fileResult = merged.drop(columns='UniqueID')
    
    # Return
    return fileResult
    
def duplicate_column(df, list, add=' - Copy'):
    for x in list:
        df[x+add] = df[x]
    
# def generate_sql_query(df, row_exists=False, uniqueID=None):
#     if uniqueID is None:
#         raise ValueError("uniqueID must not be null")
    
#     # Construct SQL query
#     columns = ",".join(df.columns)
#     values = ",".join(["(" + ",".join(["'" + str(val) + "'" for val in row]) + ")" for _, row in df.iterrows()])
#     if row_exists:
#         sql_query = f"UPDATE isr_modify_m2m SET --- WHERE unique_id = {uniqueID}"
#     else:
#         sql_query = f"INSERT IGNORE INTO isr_modify_m2m ({columns}, unique_id) VALUES {values}, {uniqueID}"
#     return sql_query

def main(fileName, output_dir):
    try:
        # Import file
        df = pd.DataFrame(pd.read_excel(fileName,
                                        converters={'RR':str,'App ID':str},
                                        skiprows=1)) 

        # result as data frame
        data_to_export = process_file(df)

        current_dateTime = datetime.now().strftime("%Y%m%d_%H%M_%S%f")[:-3]

        resultFileName = 'importXML_Data_'+current_dateTime+'.csv'
        filePath = output_dir+'/'+resultFileName
        data_to_export.to_csv(filePath, index=False)

        # print(f"CSV file successfully exported to: {filePath}")
        return f"{filePath}"
    except Exception as e:
        # print(f"Error exporting CSV file: {e}")
        return f"Error exporting CSV file: {e}"
    
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python your_script.py input_file output_dir")
        sys.exit(1)

    # Parse command-line arguments
    input_file = sys.argv[1]
    output_dir = sys.argv[2]

    # Run main function and print result
    print(main(input_file, output_dir))