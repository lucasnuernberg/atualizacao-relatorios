from json import load, loads
from requests import get
from gspread import authorize
from datetime import datetime
import oauth2client.service_account
class Raiz:
    
    def __init__(self):
        
        with open('./traducaoSkus.json') as fileData:
            self.dadosErrados = load(fileData)

    def conectarPlanilha(self, nomePlanilha):

        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = oauth2client.service_account.ServiceAccountCredentials.from_json_keyfile_name('./clientJson.json', scope)
        self.client = authorize(creds)

        return self.client.open(nomePlanilha)

    def buscarData(self, mesesVoltar):

        dataHoje = datetime.today()

        dia = dataHoje.day
        dia3meses = dia
        mes = dataHoje.month
        mesesSem31 = [2, 4, 6, 9, 11]

        if (dia == 31) and ((mes - mesesVoltar) in mesesSem31):
            if (mes - 1) == 2:
                dia -= 2
            else:
                dia -= 1

        if (dia == 31) and ((mes - 3) in mesesSem31):
            if (mes - 1) == 2:
                dia3meses = dia - 2
            else:
                dia3meses = dia - 1

        dataCompletaAtual = (dataHoje.strftime('%d/%m/%Y'))
        data1MesAntes = (dataHoje.strftime(f'{dia}/{mes - mesesVoltar}/%Y'))
        data2MesesAntes = (dataHoje.strftime(f'{dia3meses}/{mes - 2}/%Y'))

        return [data1MesAntes, dataCompletaAtual, data2MesesAntes]

    def buscar_dados(self, modulo, pagina, opcional=''):
        headers = { 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36' }

        apiResponse = get(f'https://bling.com.br/Api/v2/{modulo}/page={pagina}/json/&apikey={self.chaveApi}{opcional}', headers=headers)
        
        return loads(apiResponse.text)
