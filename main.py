from models.Plataformas import AnalisePlataformas
from models.Atualizacao import Atualizacao
from models import Raiz
from dotenv import load_dotenv
from json import load
from os import getenv

load_dotenv()
#Global Variables
hakonParts = getenv('APIKEYHAKON')

with open('plataformas.json') as fileData:
    plataformasHakonParts = load(fileData)

funcoesGerais = Raiz()
dataInicial = funcoesGerais.buscarData(1)[0]
dataFinal = funcoesGerais.buscarData(1)[1]

def atualizarPlanilhaCompras():

    plataformas = AnalisePlataformas(chaveApi=hakonParts, plataformas=plataformasHakonParts, planilha='Mercado Livre', dataInicial=dataInicial, dataFinal=dataFinal).runAllProcesses()
    #Inicializa a classe
    atualiza = Atualizacao(hakonParts, dataInicial=dataInicial, dataFinal=dataFinal)
    #compras
    atualiza.buscarPedidosCompra()
    atualiza.preencherAbaPedidosCompra()
    
    atualiza.calcularQuantidadeVendida(dataAnterior=dataInicial, dataAtual=dataFinal)
    atualiza.filtrarProdutos()
    
    if atualiza.abaDados.cell(2, 4).value != atualiza.mesAtual:
        atualiza.preencherPlanihaDados()
    
    if atualiza.abaRelatorio.cell(2, 20).value != atualiza.mesAtual:
        atualiza.preencherAbaRelatorio()
    
    atualiza.preenchePlanilhaEstoque()
    atualiza.preencherAbaVendas()

atualizarPlanilhaCompras()