from models import Raiz, Atualizacao, AnalisePlataformas
from dotenv import load_dotenv
from json import load
from os import getenv
import logging

logging.basicConfig(level=logging.INFO, filename="logger.log", format="%(asctime)s - %(levelname)s - %(message)s")

load_dotenv()

#Global Variables
hakonParts = getenv('API')
with open('plataformas.json') as fileData:
    plataformasHakonParts = load(fileData)
    
funcoesGerais = Raiz()
dataInicial = funcoesGerais.buscarData(1)[0]
dataFinal = funcoesGerais.buscarData(1)[1]
nomePlanilhaAnalise = getenv('NOMEPLANILHA')

def atualizarPlataformas():
        
    plataformas = AnalisePlataformas(chaveApi=hakonParts, plataformas=plataformasHakonParts, planilha=nomePlanilhaAnalise, dataInicial=dataInicial, dataFinal=dataFinal).runAllProcesses()

def atualizarPlanilhaCompras():
    
    
    #Inicializa a classe
    atualiza = Atualizacao(hakonParts, dataInicial=dataInicial, dataFinal=dataFinal)
    #compras
    atualiza.buscarPedidosCompra()
    atualiza.preencherAbaPedidosCompra()
    atualiza.calcularQuantidadeVendida(dataAnterior=dataInicial, dataAtual=dataFinal)
    atualiza.filtrarProdutos()  
    atualiza.preenchePlanilhaEstoque()
    atualiza.preencherAbaVendas()  
    
    if atualiza.abaDados.cell(2, 4).value != atualiza.mesAtual:
        atualiza.preencherPlanihaDados()
        logging.info('Preencheu aba Dados')
    
    if atualiza.abaRelatorio.cell(2, 20).value != atualiza.mesAtual:
        atualiza.preencherAbaRelatorio()
        logging.info('Preencheu aba Relatorio')
    

try:     
    atualizarPlanilhaCompras()
    atualizarPlataformas()
except Exception as e:
    logging.critical(f'Error: {e}')