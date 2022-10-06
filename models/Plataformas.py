from requests import get
from xmltodict import parse
from time import sleep
from .Raiz import Raiz
import logging

logging.basicConfig(level=logging.DEBUG, filename="logger.log", format="%(asctime)s - %(levelname)s - %(message)s")

class AnalisePlataformas(Raiz):

    def __init__(self, chaveApi, plataformas, planilha, dataInicial='', dataFinal=''):
        Raiz.__init__(self)
        
        #Setting Global Variables
        self.planiha = self.conectarPlanilha(nomePlanilha=planilha)
        self.abaAnalise = self.planiha.worksheet('Analise Plataformas')
        self.chaveApi = chaveApi
        self.plataformas = plataformas
        self.dataAnterior = dataInicial
        self.dataFinal = dataFinal
        
    def runAllProcesses(self):

        print('Rodando analise de plataformas')
        self.criarDadosPlataformas()
        print('Buscando Produtos')
        self.buscarProdutos()
        print('Buscando pedidos')
        self.calcularFaturamento()
        print('Buscando Notas')
        self.buscarNotasFiscais()
        print('Lendo Notas')
        self.lerNotasFiscais()
        print("Preenchendo Aba Analise Plataformas")
        self.preeencherPlanilha()
        
    
    def apurarLucro(self):

        print('Rodando analise de plataformas')
        self.criarDadosPlataformas()
        print('Buscando Produtos')
        self.buscarProdutos()
        print('Buscando pedidos')
        self.calcularFaturamento()
        
        for codigo, dadosPlataforma in self.objetosFaturamento.items():
            print(f'------------------------------')
            print(f'LOJA: {dadosPlataforma["loja"]}')
            print(f'FATURAMENTO: {dadosPlataforma["faturamento"]}')
            print(f'CUSTO PRODUTOS: {dadosPlataforma["custoProdutos"]}')
            print(f'QUANTIDADE VENDIDA: {dadosPlataforma["quantidadeVendida"]}')
        

    def buscarProdutos(self):

        pagina = 1
        self.todosProdutos = {}

        while True:
            #Time required to not overload API
            sleep(.5)
            responseApi = self.buscar_dados("produtos", pagina)

            listaProdutos = dict(responseApi["retorno"]).get("produtos")            

            if listaProdutos == None:
                print(f'Motivo da pausa na busca de produtos: {listaProdutos}')
                break
            
            for produto in listaProdutos:
                sku = produto["produto"]["codigo"].upper().strip()
                if sku != '':
                    self.todosProdutos[sku] = produto["produto"]["precoCusto"]
                    
            pagina += 1
            
    def buscarPedidos(self):
        
        pagina = 1
        dicionarioListaPedidos = []
        while True:
            sleep(.5)
            jsonPedidos = self.buscar_dados("pedidos", pagina, f'&filters=dataEmissao[{self.dataAnterior} TO {self.dataFinal}];idSituacao[9]')

            listaPedidos = dict(jsonPedidos["retorno"]).get("pedidos")

            if listaPedidos == None:
                print(f'Motivo de para buscar pedidos: {jsonPedidos}')
                break

            for pedido in listaPedidos:
                dicionarioListaPedidos.append(pedido)

            pagina += 1
        return dicionarioListaPedidos

    def buscarNotasFiscais(self):
        
        pagina = 1
        listaNotasFiscais = []

        while True:
            
            sleep(.5)
            responseApi = self.buscar_dados("notasfiscais", pagina, f'&filters=dataEmissao[{self.dataAnterior} TO {self.dataFinal}]; situacao[6]')
            
            listaNotas = dict(responseApi["retorno"]).get("notasfiscais")            

            if listaNotas == None:
                print(f'Motivo da pausa na busca de notas: {responseApi}')
                break

            for nota in listaNotas:
                listaNotasFiscais.append(nota) if nota["notafiscal"]["serie"] == "2" and nota["notafiscal"]["cliente"]["nome"].upper() not in "MERCADO ENVIOS SERVICOS DE LOGISTICA LTDA.EBAZAR.COM.BR LTDA" else ""                    

            pagina += 1

        self.notasFull = listaNotasFiscais

    def criarDadosPlataformas(self):
        
        self.objetosFaturamento = dict()

        for nomeLoja, codigoLoja in self.plataformas.items():
            
            self.objetosFaturamento[codigoLoja] = {             
                
                "codigoLoja" : codigoLoja,
                "loja" : nomeLoja,
                "faturamento" : 0,
                "quantidadeVendida" : 0,
                "faturamentoFull" : 0,
                "quantidadeVendasFull" : 0,
                "custoProdutos" : 0                
            }

    def calcularFaturamento(self):

        for pedido in self.buscarPedidos():
            try:
                lojaPedido = pedido["pedido"]["loja"]

                loja = self.objetosFaturamento.get(lojaPedido)
                
                if loja != None:
                    for item in pedido["pedido"]["itens"]:
                        if item["item"]["codigo"] != None:                            
                            
                            quantidadeVendida = float(item["item"]["quantidade"])
                            
                            fator = (float(pedido["pedido"]["totalprodutos"]) - float(pedido["pedido"]["desconto"].replace(',', '.'))) / float(pedido["pedido"]["totalprodutos"])
                            loja["faturamento"] += (float(item["item"]["valorunidade"]) * quantidadeVendida * fator)
                            loja["quantidadeVendida"] += float(item["item"]["quantidade"])
                            loja["custoProdutos"] += quantidadeVendida * float(item["item"]["precocusto"])
                            
                else:
                    logging.warning(f'Loja não encontrada{lojaPedido}')

            except Exception as e:
                logging.warning(f'{e} - Pedido: {pedido["pedido"]["numero"]}')
                

    def lerNotasFiscais(self):

        for nota in self.notasFull:

            lojaNota = nota["notafiscal"]["loja"]

            loja = self.objetosFaturamento.get(lojaNota)

            if loja != None:
                try:
                    dados = get(url=nota["notafiscal"]["xml"])
                    dadosx = parse(dados.content)
                    listaItens = dadosx["nfeProc"]["NFe"]["infNFe"]["det"]

                    if type(listaItens) is list:

                        for item in listaItens:
                            
                            quantidadeComprada = item["prod"]["qCom"]
                            skuProduto = item["prod"]["cProd"].upper().strip()
                            
                            loja["quantidadeVendasFull"] += float(quantidadeComprada)
                            
                            #CONFERENCIA DE SKU ERRADO                            
                            produtoErrado = self.dadosErrados.get(skuProduto)
                            
                            if produtoErrado != None:
                                valor = self.todosProdutos.get(produtoErrado.upper().strip())
                                loja["custoProdutos"] += (float(valor) + float(quantidadeComprada))

                            else:                                
                                valor = self.todosProdutos.get(skuProduto)                                
                                loja["custoProdutos"] += (float(valor) + float(quantidadeComprada))              
                            
                    else:
                        
                        quantidadeComprada = listaItens["prod"]["qCom"]
                        skuProduto = listaItens["prod"]["cProd"]
                        
                        loja["quantidadeVendasFull"] += float(quantidadeComprada)
                        
                        #CONFERENCIA DE SKU ERRADO
                        
                        produtoErrado = self.dadosErrados.get(skuProduto.upper().strip())
                            
                        if produtoErrado != None:

                            valor = self.todosProdutos.get(produtoErrado.upper().strip())
                            loja["custoProdutos"] += (float(valor) + float(quantidadeComprada))                          
                        
                        else:
                            
                            valor = self.todosProdutos.get(skuProduto.upper().strip())
                            loja["custoProdutos"] += (float(valor) + float(quantidadeComprada))

                    loja["faturamentoFull"] += float(nota["notafiscal"]["valorNota"])

                except Exception as e:
                    logging.warning(f'erro: {e} - nota {nota["notafiscal"]["numero"]}')                   

            else:
                logging.info(f'Loja não encontrada: {lojaNota}')
                
                
    def preeencherPlanilha(self):

        i = 2
        self.abaAnalise.update_cell(1, 1, "LOJA")
        self.abaAnalise.update_cell(2, 1, "CODIGO")
        self.abaAnalise.update_cell(3, 1, "FATURAMENTO 30 DIAS")
        self.abaAnalise.update_cell(4, 1, "QUANTIDADE VENDIDA")
        self.abaAnalise.update_cell(5, 1, "FATURAMENTO FULL")
        self.abaAnalise.update_cell(6, 1, "QUANTIDADE VENDIDA FULL")

        for codigoLoja, lojaMarket in self.objetosFaturamento.items():

            self.abaAnalise.update_cell(1, i, lojaMarket["loja"])
            sleep(2)
            self.abaAnalise.update_cell(2, i, lojaMarket["codigoLoja"])
            sleep(2)
            self.abaAnalise.update_cell(3, i, lojaMarket["faturamento"])
            sleep(2)
            self.abaAnalise.update_cell(4, i, lojaMarket["quantidadeVendida"])
            sleep(2)
            self.abaAnalise.update_cell(5, i, lojaMarket["faturamentoFull"])
            sleep(2)
            self.abaAnalise.update_cell(6, i, lojaMarket["quantidadeVendasFull"])
            sleep(2)
            self.abaAnalise.update_cell(8, i, lojaMarket["custoProdutos"])
            sleep(2)
            i += 1

        self.abaAnalise.update_cell(1, 13, "ATUALIZAÇÃO")
        self.abaAnalise.update_cell(3, 13, self.dataFinal)