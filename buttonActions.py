import pandas as pd
from django.shortcuts import redirect


def elimina_modulo(request):

    to_delete = request.POST['pitcher']
    df_bancale = pd.read_excel('OpenBancale.xlsx', dtype='string')

    df_bancale_new = []

    for i in range(len(df_bancale['SerialNumber'])):

        if str(df_bancale['SerialNumber'][i]) != to_delete:
            df_bancale_new.append(df_bancale.iloc[i, :])

    df_bancale_new = pd.DataFrame(df_bancale_new)

    if len(df_bancale_new) == 0:
        df_bancale_new = pd.read_excel('EmptyOpenBancale.xlsx')

    else:

        for i in range(len(df_bancale_new['SerialNumber'])):
            df_bancale_new["Index"] = str(i + 1)

    df_bancale_new.to_excel("OpenBancale.xlsx", index=False)

    return redirect('')


def carica_bancale(request):

    file_name = request.POST['Scegli bancale']

    df_bancale = pd.read_excel("Database bancali/" + file_name, dtype='string')
    df_bancale.to_html('Bancale.html', index=False)
    df_bancale.to_excel('OpenBancale.xlsx', index=False)

    return redirect('')
