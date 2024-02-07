import pandas as pd
import os
import pyodbc
import csv
from printUtilities import printLabel


def add_to_bancale(label):

    df_bancale = pd.read_excel('OpenBancale.xlsx', dtype='string')

    if len(df_bancale['Index']) == 0:
        prev_index = 0

    else:
        prev_index = int(df_bancale['Index'][len(df_bancale) - 1])

    try:
        my_dict = {'Index': str(prev_index + 1), 'SerialNumber': str(label['SerialNumber'][0]),
                   'Year of production': "N.D.", 'Pmpp': str(label['Pmpp'][0]), 'IrradiatedEnergy': str(label['E'][0]),
                   'Temp': str(label['Temp'][0]), 'Uoc': str(label['Uoc'][0]), 'Isc': str(label['Isc'][0]),
                   'Umpp': str(label['Umpp'][0]), 'Impp': str(label['Impp'][0]), 'Rs': str(label['Rs'][0]),
                   'Rsh': str(label['Rsh'][0]), 'FlashDate': label['FlashDate'][0]}

    except Exception as err:
        print(err)
        my_dict = {'Index': str(prev_index + 1), 'Serial number': label['SerialNumber'],
                   'Year of production': str(label['ProdYear']), 'Measured power [W]': str(label['Pmpp']),
                   'Irradiated power [W/m^2]': str(label['E']), 'Lab. temperature [°C]': str(label['Temp']),
                   'Open circuit voltage [V]': str(label['Uoc']), 'Short circuit current [A]': str(label['Isc']),
                   'Maximum power voltage [V]': str(label['Umpp']), 'Maximum power current [A]': str(label['Impp']),
                   'Series resistance [omh]': str(label['Rs']), 'Shunt resistance [omh]': str(label['Rsh']),
                   'Test date': label['FlashDate']}

    df_in = pd.DataFrame(my_dict, index=[0])

    df_bancale = pd.concat([df_bancale, df_in])

    # dfBancale.append(dfIn)
    df_bancale.to_excel('OpenBancale.xlsx', index=False)


def last_label_read_ps_load_db():

    conn = pyodbc.connect(r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                          r"DBQ=C:\PSLoad\DB\PSLoad.mdb;")
    cur = conn.cursor()

    sql = 'SELECT * FROM Results;'  # your query goes here
    rows = cur.execute(sql)

    # salvo in Results.csv
    with open('Results.csv', 'w') as fou:
        csv_writer = csv.writer(fou)
        csv_writer.writerows(rows)

    cur.close()
    conn.close()

    df = pd.read_csv('Results.csv', encoding='cp1252', header=None)
    df.columns = [
        'SerialNumber', 'ModuleName', 'FrameTestName', 'MonitorCellName', 'Uref', 'FlashDate', 'Operator', 'Comment',
        'Batch', 'Class', 'Temp', 'Uoc', 'Isc', 'Pmpp', 'Umpp', 'Impp', 'Uop', 'Iop', 'FF', 'E', 'NMod', 'NCell', 'Rs',
        'Rsh', 'PE1_Resistance', 'PE1_Current', 'PE1_Time', 'PE1_InspectionOK', ' PE2_Resistance', 'PE2_Current',
        'PE2_Time', 'PE2_InspectionOK', 'PE3_Resistance', 'PE3_Current', 'PE3_Time', 'PE3_Inspect', 'HV_Current',
        'HV_Voltage', 'HV_Time', 'HV_InspectionOK', 'ISO_Resistance', 'ISO_Voltage', 'ISO_Time', 'ISO_Inspect',
        'FLASHER_InspectionOK', 'Remeasure', 'CalibID', 'CalibDate', 'FlasherID', 'Reduction'
                  ]

    df = df.sort_values(by='FlashDate')
    last_val = df.iloc[len(df) - 1]

    label_data = {
        "SerialNumber": last_val['SerialNumber'], 'Pmpp': last_val['Pmpp'], 'E': last_val['E'],
        'Temp': last_val['Temp'], 'Uoc': last_val['Uoc'], 'Isc': last_val['Isc'], 'Umpp': last_val['Umpp'],
        'Impp': last_val['Impp'], 'FillFactor': last_val['FF'], 'Rs': last_val['Rs'], 'Rsh': last_val['Rsh'],
        'FlashDate': last_val['FlashDate']
    }

    label_data = pd.DataFrame(label_data, index=[0])

    return label_data


def check_modification_date():

    print("Controllo data ultima modifica del database")

    # leggo l'ultima volta che ho visto una modifica del database tLastModify
    df_last_seen_mod = pd.read_excel("lastLabelTimeStamp.xlsx")
    last_seen_mod_timestamp = df_last_seen_mod['last t'][0]

    # leggo l'ultima volta che è stato modificqato il DB tLastDBUpdate
    file_path = r'C:\PSLoad\DB\PSLoad.mdb'
    last_db_update = os.path.getmtime(file_path)

    # se tLastDBIUpdate > tLastModify
    if round(last_seen_mod_timestamp) < round(last_db_update):  # se il DB è stato modificato
        print("Modifica rilevata")

        # leggo  i dati dell'ultima misura da stampare
        last_label_data = last_label_read_ps_load_db()

        # la stampo alla stampa etichette
        printLabel(last_label_data)

        # la aggiungo al bancale
        add_to_bancale(last_label_data)

        # salvo il valore dell'ultima data di stampa
        df_last_seen_mod['last t'] = last_db_update
        df_last_seen_mod.to_excel("lastLabelTimeStamp.xlsx")
