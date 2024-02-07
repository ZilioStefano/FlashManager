from fpdf import FPDF
from pdf2image import convert_from_path
import os
import pandas as pd
from datetime import datetime
from django.shortcuts import redirect


def printBancale(request):

    dfBancale = pd.read_excel('OpenBancale.xlsx')
    dfBancaleRounded = []
    for i in range(len(dfBancale)):
        dfBancaleRounded.append(arrotondaValori(dfBancale.iloc[i]))
    dfBancaleRounded = pd.DataFrame.from_dict(dfBancaleRounded)
    dfBancaleHead = ["", "Serial number", "YP", "MP [W]", "IP [W/m\u00b2]", "LT [°C]", "VOC [V]", "ISC [A]", "VMP [V]", "IMP [A]", "RS [omh]", "RSh [omh]", "Test date"]
    pdf = FPDF(format="A4", orientation='P')
    pdf.set_margins(6, 5)
    # pdf = FPDF(format=(297, 225.56), unit='mm')
    pdf.add_page()

    # pdf.set_font("Times", size=5)

    colWidth = 18
    rowHeight = 5
    # x0 = pdf.get_x()
    # y0 = pdf.get_y()
    # y = y0
    y = pdf.get_y()
    x = pdf.get_x()
    pdf.set_font("Times", 'b', size=12)

    pdf.multi_cell(10*colWidth, 0.5 * rowHeight, "Rapporto di prova redatto ai sensi dell'allegato VI al D.lgs 49/2014 per il trasporto di AEE usate", border=0, align='C')
    pdf.ln(5)
    pdf.set_font("Times", size=7)
    y = pdf.get_y()
    pdf.multi_cell(10*colWidth, 0.2 * rowHeight, "CATEGORIA (ALLEGATO I): 4-apparecchiature di consumo e pannelli fotovoltaici", border=0, align='L')

    pdf.ln()

    pdf.ln()
    pdf.multi_cell(10*colWidth, 0.2 * rowHeight, "NOME (ALLEGATO II): 4.9-pannelli fotovoltaici", border=0, align='L')
    pdf.ln(5)

    # creo la colonna degli indici
    Indexes = []
    # Indexes.append("")

    for i in range(len(dfBancale)):
        Indexes.append(str(i+1))

    dfBancaleRounded = pd.concat([pd.Series(Indexes), dfBancaleRounded], axis=1, ignore_index=True)


    for row in range(len(dfBancaleRounded[:])+1):

        if row == 0:
            hFactor = 1
            pdf.set_font("Times", 'b', size=7)
        else:
            hFactor = 1
            pdf.set_font("Times", size=7)

        y = pdf.get_y()
        colWidth = 0
        currPos = 0
        for col in range(12):
            PrevCol = colWidth
            colWidth = 18

            if col == 0:
                colWidth = colWidth/4
            #     # pdf.set_x(x + col * colWidth)
            # # elif col == 1:
            #     # colWidth = colWidth
            #     # pdf.set_x(x + col * colWidth/2)
            else:
                colWidth = colWidth

            currPos = currPos + PrevCol

            pdf.set_y(y)
            pdf.set_x(x + currPos)

            if col == 1:
                pdf.set_font("Times", size=6)
            else:
                pdf.set_font("Times", size=7)

            if row == 0:

                pdf.set_font("Times", 'b', size=7)
                try:
                    pdf.multi_cell(colWidth, hFactor*rowHeight, dfBancaleHead[col], border=1, align='C')
                except Exception as err:
                    print(err)

            else:

                try:
                    pdf.multi_cell(colWidth, hFactor*rowHeight, dfBancaleRounded[col][row-1], border=1, align='C')
                except Exception as err:
                    print(err)

            # print(str(row)+", "+str(col))

            pdf.set_y(y+rowHeight)

        if row == 32:
            pdf.add_page()

    # colWidth = 18

    IL = 0.3
    rowHeight = 3
    pdf.ln(5)
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(10*colWidth, rowHeight, "Legend:", border=0, align='L')
    pdf.ln(IL)

    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(0.5*colWidth, rowHeight, "YP:", border=0, align='L')
    pdf.set_y(y)
    pdf.set_x(x + 0.5*colWidth)
    pdf.set_font("Times", size=8)
    pdf.multi_cell(2*colWidth, rowHeight, "Year of production", border=0, align='L')
    pdf.ln(IL)

    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(0.5*colWidth, rowHeight, "MP:", border=0, align='L')
    pdf.set_y(y)
    pdf.set_x(x + 0.5*colWidth)
    pdf.set_font("Times", size=8)
    pdf.multi_cell(2*colWidth, rowHeight, "Measured power", border=0, align='L')
    pdf.ln(IL)

    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(0.5*colWidth, rowHeight, "IP:", border=0, align='L')
    pdf.set_y(y)
    pdf.set_x(x + 0.5*colWidth)
    pdf.set_font("Times", size=8)
    pdf.multi_cell(2*colWidth, rowHeight, "Irradiated power", border=0, align='L')
    pdf.ln(IL)

    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(0.5*colWidth, rowHeight, "LT:", border=0, align='L')
    pdf.set_y(y)
    pdf.set_x(x + 0.5*colWidth)
    pdf.set_font("Times", size=8)
    pdf.multi_cell(2*colWidth, rowHeight, "Laboratory temperature", border=0, align='L')
    pdf.ln(IL)

    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(0.5*colWidth, rowHeight, "VOC:", border=0, align='L')
    pdf.set_y(y)
    pdf.set_x(x + 0.5*colWidth)
    pdf.set_font("Times", size=8)
    pdf.multi_cell(2*colWidth, rowHeight, "Open circuit voltage", border=0, align='L')
    pdf.ln(IL)

    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(0.5*colWidth, rowHeight, "ISC:", border=0, align='L')
    pdf.set_y(y)
    pdf.set_x(x + 0.5*colWidth)
    pdf.set_font("Times", size=8)
    pdf.multi_cell(2*colWidth, rowHeight, "Short circuit current", border=0, align='L')
    pdf.ln(IL)

    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(0.5*colWidth, rowHeight, "VMP:", border=0, align='L')
    pdf.set_y(y)
    pdf.set_x(x + 0.5*colWidth)
    pdf.set_font("Times", size=8)
    pdf.multi_cell(2*colWidth, rowHeight, "Maximum power voltage", border=0, align='L')
    pdf.ln(IL)

    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(0.5*colWidth, rowHeight, "IMP:", border=0, align='L')
    pdf.set_y(y)
    pdf.set_x(x + 0.5*colWidth)
    pdf.set_font("Times", size=8)
    pdf.multi_cell(2*colWidth, rowHeight, "Maximum power current", border=0, align='L')
    pdf.ln(IL)

    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(0.5*colWidth, rowHeight, "RS:", border=0, align='L')
    pdf.set_y(y)
    pdf.set_x(x + 0.5*colWidth)
    pdf.set_font("Times", size=8)
    pdf.multi_cell(2*colWidth, rowHeight, "Series resistance", border=0, align='L')
    pdf.ln(IL)

    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=8)
    pdf.multi_cell(0.5*colWidth, rowHeight, "RSh:", border=0, align='L')
    pdf.set_y(y)
    pdf.set_x(x + 0.5*colWidth)
    pdf.set_font("Times", size=8)
    pdf.multi_cell(2*colWidth, rowHeight, "Shunt resistance", border=0, align='L')
    pdf.ln(IL)

    pdf.ln(5)
    pdf.multi_cell(10*colWidth, 0.5*rowHeight, "Measures performed by Regenerasolar S.r.l, Via San Francesco 20, 20360, Pove del Grappa, Italy", border=0, align='L')
    pdf.ln(2)

    pdf.multi_cell(10*colWidth, 0.5*rowHeight, "Measurements set-up: Flash test with pulsed light at 1000 W/m\u00b2 circa", border=0, align='L')
    pdf.ln(2)

    Now = datetime.now()
    NowStr = Now.strftime('%d/%m/%Y')
    pdf.multi_cell(10*colWidth, 0.5*rowHeight, "Document creation date: "+NowStr, border=0, align='L')

    pdf.output("Bancale.pdf")
    # os.startfile("Bancale.pdf")

    # win32print.SetDefaultPrinter('Canon IP C910 - PRISMAsync PS')
    os.startfile("Bancale.pdf")

    Now = datetime.now()
    Name =Now.strftime('%Y%m%d%H%M')
    pdf.output("Stampe/Bancali/"+Name+".pdf")
    dfBancale.to_excel("Database bancali/"+Name + ".xlsx", index=False)

    # leggo la tabella del bancale
    Bancale = pd.read_excel("OpenBancale.xlsx")
    Bancale.to_html("OpenBancale.html")

    with open('OpenBancale.html', 'r') as f:
        BancaleHTML = f.read()

    # BancalatiHTML, BancaliList = showBancaleInCorso("EmptyOpenBancale.xlsx")

    dfEmptyBancale = pd.read_excel("EmptyOpenBancale.xlsx")
    dfEmptyBancale.to_excel("OpenBancale.xlsx", index=False)
    dfEmptyBancale.to_html("OpenBancale.html")
    # SNList = ShowLastReads(request)
    #
    # Logo = Image.open('im innovation logo abbreviato BASSA RISOLUZIONE.png')
    # Logo64 = image_to_base64(Logo)
    #
    # # apro il file BancaleAperto
    #
    # DatiHTML = {'SNList': SNList, 'Bancali': BancalatiHTML, 'BancaliList': BancaliList, "Logo": Logo64}
    #
    # # flaggo e muto la matricola stampata dall'alatra lista
    # Template = "HP_AND_POPUP.html"

    return redirect('')


def createLabel(dataLabel):

    print('Printing label...')

    pdf = FPDF(format=(104, 50), unit='mm')

    # pdf = FPDF(format=(208, 296), unit='mm')
    pdf.set_margins(2,0)
    # pdf = FPDF(format=(297, 225.56), unit='mm')
    pdf.add_page()
    FontSize = 11
    pdf.set_font("Times", size=FontSize)
    colWidth = 32
    lineHeight = 4
    chSize = 7
    SpaceBetweenColumns = 1.6
    SpaceBetweenSecondColumn = 2.5


    x0 = pdf.get_x()
    x = pdf.get_x()
    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Item:", align="LEFT")
    pdf.set_y(y)
    pdf.set_x(x + 1*colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(1*colWidth, lineHeight, "Used PV module", align="LEFT")

    # x0 = pdf.get_x()
    # y = pdf.get_y()
    pdf.set_y(y)
    pdf.set_x(x0 + SpaceBetweenColumns*colWidth)
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Serial number", align="LEFT")
    pdf.set_y(y)
    pdf.set_x(x0 + SpaceBetweenSecondColumn*colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, dataLabel['SerialNumber'], align="LEFT")

    x = x0
    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Measured Power", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, str(dataLabel['Power'])+" W", align="RIGHT")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenColumns*colWidth)
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Year of production", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenSecondColumn*colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "N.D.", align="RIGHT")

    #
    x = x0
    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Irradiated power", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, str(dataLabel['IrradiatedEnergy'])+" W/m\u00b2", align="RIGHT")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenColumns*colWidth)
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Test temperature", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenSecondColumn*colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, str(dataLabel["Temperature"])+" °C", align="RIGHT")

    x = x0
    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Open circuit voltage", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, str(dataLabel['VOC'])+" V", align="RIGHT")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenColumns*colWidth)
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Short circuit current", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenSecondColumn*colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, str(dataLabel['ISC'])+" A", align="RIGHT")

    x = x0
    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Maximum power voltage", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, str(dataLabel['VMP'])+" V", align="RIGHT")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenColumns*colWidth)
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Maximum power current", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenSecondColumn*colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, str(dataLabel['IMP'])+" A", align="RIGHT")

    x = x0
    y = pdf.get_y()

    x = x0
    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Series resistance", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, str(dataLabel['Rs'])+" Omh", align="RIGHT")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenColumns*colWidth)

    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Shunt resistance", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenSecondColumn*colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, str(dataLabel['RSh'])+" Omh", align="RIGHT")

    x = x0
    y = pdf.get_y()
    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Type of measurement", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Flash test", align="RIGHT")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenColumns*colWidth)

    pdf.set_font("Times", 'b', size=chSize)
    pdf.multi_cell(colWidth, lineHeight, "Performed at", align="CENTER")
    pdf.set_y(y)
    pdf.set_x(x + SpaceBetweenSecondColumn*colWidth)
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(colWidth, 1*lineHeight, dataLabel['FlashDate'], align="RIGHT")
    # x = x0
    # y = pdf.get_y()
    chSize = 4
    pdf.set_font("Times", size=chSize)
    pdf.multi_cell(4*colWidth, 0.4*lineHeight, "Performed by Regenerasolar S.r.l, Via Vallina Orticella, 31030, Borso del Grappa, Italy", align="CENTER")
    # pdf.set_y(y)
    # pdf.set_x(x + 2*colWidth)
    # pdf.multi_cell(2*colWidth, lineHeight, "Regenerasolar S.r.l, Via San Francesco 20, 20360, Pove del Grappa, Italy", align="RIGHT")

    pdf.output(dataLabel["SerialNumber"]+".pdf")
    os.startfile(dataLabel["SerialNumber"]+".pdf")#, 'print')

    images = convert_from_path(dataLabel["SerialNumber"] + ".pdf", poppler_path=r"C:\Users\Utente\Downloads\poppler-23.11.0\Library\bin")

    for i in range(len(images)):
        # Save pages as images in the pdf
        images[i].save('page' + str(i) + '.png', 'PNG')

    return images


def arrotondaValori(Label):

    try:

        NewLabel = {"Power": str(round(Label['Pmpp'][0]))}
        # NewLabel['Power'] = str(round(Label['Pmpp']))
        NewLabel['IrradiatedEnergy'] = str(round(Label['E'][0]))
        NewLabel['Temperature'] = str(round(Label['Temp'][0], 1))
        NewLabel['VOC'] = str(round(Label['Uoc'][0],1))
        NewLabel['ISC'] = str(round(Label['Isc'][0],2))
        NewLabel['VMP'] = str(round(Label['Umpp'][0],2))
        NewLabel['IMP'] = str(round(Label['Impp'][0],2))
        NewLabel['Rs'] = str(round(Label['Rs'][0],2))
        NewLabel['RSh'] = str(round(Label['Rsh'][0]))
        NewLabel['FlashDate'] = str(Label['FlashDate'][0])

    except Exception as err:

        print(err)
        NewLabel = {"SerialNumber": Label['SerialNumber']}
        NewLabel["YearOfProduction"] = "N.D."
        NewLabel["Power"] = str(round(Label['Pmpp']))
        # NewLabel['Power'] = str(round(Label['Pmpp']))
        NewLabel['IrradiatedEnergy'] = str(round(Label['IrradiatedEnergy']))
        NewLabel['Temperature'] = str(round(Label['Temp'], 1))
        NewLabel['VOC'] = str(round(Label['Uoc'],1))
        NewLabel['ISC'] = str(round(Label['Isc'],2))
        NewLabel['VMP'] = str(round(Label['Umpp'],2))
        NewLabel['IMP'] = str(round(Label['Impp'],2))
        NewLabel['Rs'] = str(round(Label['Rs'],2))
        NewLabel['RSh'] = str(round(Label['Rsh']))
        NewLabel['FlashDate'] = str(Label['FlashDate'])

    return NewLabel


def printLabel(Label):

    LabelDiStampa = arrotondaValori(Label)
    LabelDiStampa["SerialNumber"] = Label["SerialNumber"][0]
    LabelDiStampa["FlashDate"] = Label["FlashDate"][0]

    createLabel(LabelDiStampa)
