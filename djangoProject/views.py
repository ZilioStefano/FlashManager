from threading import Timer
import time
from checkModificationDate import check_modification_date
import pandas as pd
from PIL import Image
from django.shortcuts import render
from graphicalUtilities import image_to_base64
import webbrowser


class RepeatTimer(Timer):
    def run(self):
        while not self.finished.wait(self.interval):
            self.function(*self.args, **self.kwargs)
            print(' ')


timer = RepeatTimer(15, check_modification_date, "")
timer.start()  # recalling run
print('Threading started')
time.sleep(1)
url = 'http://localhost:8000/'
webbrowser.open(url)

def main(request):


    template = "HP.html"

    bancali_file = pd.read_excel("OpenBancale.xlsx")
    bancali_file.to_html("OpenBancale.html")

    logo = Image.open('im innovation logo abbreviato BASSA RISOLUZIONE.png')
    logo_im = image_to_base64(logo)

    logo = Image.open('logo regenerasolar BASSA RISOLUZIONE PER EMAIL(1).png')
    logo_rs = image_to_base64(logo)

    with open('OpenBancale.html', 'r') as f:
        bancale_html = f.read()

    bancali_sn = bancali_file["SerialNumber"]

    # apro il file BancaleAperto
    dati_html = {'ListaBancali': bancale_html, "LogoIM": logo_im, "LogoRS": logo_rs, "BancaliSN": bancali_sn}


    return render(request, template, context=dati_html)
