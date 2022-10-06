from .PedidoCompra import PedidoCompra
from .Produto import Produto
from .Raiz import Raiz
from numpy import array, append
from requests import get, put
from pandas import DataFrame
from xmltodict import parse
from time import sleep
import logging

logging.basicConfig(level=logging.DEBUG, filename="logger.log", format="%(asctime)s - %(levelname)s - %(message)s")

class Atualizacao(Raiz):

    def __init__(self, chaveApi, dataInicial='', dataFinal=''):
        Raiz.__init__(self)

        self.planilhaCompras = self.conectarPlanilha('planilhaCompras')
        self.setarAbasPlanilha()
        
        self.chaveApi = chaveApi
        self.todosProdutos = dict()
        self.pedidosCompra = array([])
        #datas
        self.mesAnterior = dataInicial
        self.mesAtual = dataFinal

    def atualizarTudo(self):
        
        self.calcularQuantidadeVendida(dataAnterior=self.mesAnterior, dataAtual=self.mesAtual)
        self.filtrarProdutos()
        self.preencherPlanihaDados()
        self.preenchePlanilhaEstoque()
        self.preencherAbaVendas()
        self.buscarPedidosCompra()
        self.preencherAbaPedidosCompra()
        self.preencherAbaRelatorio()

    #Requisições API Bling
    def atenderPedidoCompra(self, numeroPedido):

        xml = f'apikey={self.chaveApi}&xml=<?xml version="1.0" encoding="UTF-8"?><pedidocompra><situacao>1</situacao></pedidocompra>'
        atualizacaoStatusPed = put(f'https://bling.com.br/Api/v2/pedidocompra/{numeroPedido}/json/&apikey={self.chaveApi}&xml={xml}', data=xml)

    def buscarProdutos(self):
        pagina = 1
        self.todosProdutos = dict()
        while True:
            sleep(.5)

            jsonProdutos = self.buscar_dados("produtos", pagina, opcional="&estoque=S")

            listaProdutos = dict(jsonProdutos["retorno"]).get("produtos")

            if listaProdutos == None:
                break

            for produto in listaProdutos:
                sku = produto["produto"]["codigo"].upper()
                if sku != '':
                    try:
                        estrutura = produto["produto"]["estrutura"]
                        componentes = []
                        for componente in estrutura:
                            componentes.append(componente["componente"])

                        estoqueAtual = float(produto["produto"]["estoqueAtual"])
                        
                        #Na linha abaixo eu adciono os produtos ao dicionário
                        self.todosProdutos[sku] = Produto(sku, estoqueAtual, composicao=componentes)
                    except:
                        estoqueAtual = float(produto["produto"]["estoqueAtual"])
                        
                        self.todosProdutos[sku] = Produto(sku, estoqueAtual)

            pagina += 1

    def buscarNotasFiscais(self, dataAnterior, dataAtual):
        pagina = 1
        
        while True:
            sleep(.5)
            
            jsonNotas = self.buscar_dados("notasfiscais", pagina, f'&filters=dataEmissao[{dataAnterior} TO {dataAtual}]; situacao[6]')
            
            listaNotas = dict(jsonNotas["retorno"]).get("notasfiscais")
            
            if listaNotas == None:
                break
            else:       
                for nota in listaNotas:
                    if nota["notafiscal"]["serie"] == "2" and nota["notafiscal"]["cliente"]["nome"].upper() not in "MERCADO ENVIOS SERVICOS DE LOGISTICA LTDA.EBAZAR.COM.BR LTDA":
                        try:
                            dados = get(url=nota["notafiscal"]["xml"])
                            dadosx = parse(dados.content)

                            # lista ou objeto de item/itens vendidos
                            listaItens = dadosx["nfeProc"]["NFe"]["infNFe"]["det"]

                            # verifica se existe mais de 1 item na venda
                            if type(listaItens) is list:
                                for item in listaItens:
                                    
                                    skuNota = item["prod"]["cProd"].upper().strip()
                                    quantidadeComprada = item["prod"]["qCom"]
                                    valorVenda = item["prod"]["vUnCom"]

                                    traducaoSku = self.dadosErrados.get(skuNota)
                                    produto = self.todosProdutos.get(skuNota)                                    
                                    
                                    #Caso não for achado em nenhum dos 2
                                    if produto == None and traducaoSku == None:
                                        logging.warning(f'O seguinte sku deu erro: {skuNota}')
                                    
                                    elif traducaoSku == None:
                                        #caso o sku não esteja dentro dos dados errados
                                        produto = self.todosProdutos.get(skuNota)
                                        
                                        if produto == None:
                                            logging.warning(f'SKU não encontrado: {skuNota}')
                                        
                                        dataExiste = False                                    
                                        produto.vendasFull += float(quantidadeComprada)
                                        produto.faturado += float(valorVenda) * float(quantidadeComprada)
                                        # Caso o produto seja composto ele distribui para a sua composição
                                        if len(produto.composicao) > 0:
                                            for componente in produto.composicao:
                                                dataExiste = False
                                                #acha o produto

                                                produtoComposto = self.todosProdutos.get(componente["codigo"].upper().strip())
                                                
                                                if produtoComposto == None:
                                                    logging.warning(f'SKU não encontrado: {componente["codigo"]}')                      

                                                # Adciona vendas do produto composto a variavel
                                                produtoComposto.vendasFull += float(quantidadeComprada) * float(componente["quantidade"])
                                                produtoComposto.faturado += float(valorVenda) * float(quantidadeComprada)
                                                for vendasData in produtoComposto.vendasPorDia:
                                                    if vendasData["data"] == nota["notafiscal"]["dataEmissao"].split(' ')[0]:
                                                        vendasData["vendas"] += float(quantidadeComprada) * float(componente["quantidade"])
                                                        dataExiste = True
                                                        break

                                                if dataExiste == False:
                                                    objetoData = {
                                                        "data": nota["notafiscal"]["dataEmissao"].split(' ')[0],
                                                        "vendas": float(quantidadeComprada) * float(componente["quantidade"])
                                                    }
                                                    produtoComposto.vendasPorDia.append(objetoData)

                                                #
                                        for vendasData in produto.vendasPorDia:
                                            if vendasData["data"] == nota["notafiscal"]["dataEmissao"].split(' ')[0]:
                                                vendasData["vendas"] += float(quantidadeComprada)
                                                dataExiste = True

                                        if dataExiste == False:
                                            objetoData = {
                                                "data": nota["notafiscal"]["dataEmissao"].split(' ')[0],
                                                "vendas": float(quantidadeComprada)
                                            }
                                            produto.vendasPorDia.append(objetoData)

                                    #Caso o sku esteja errado
                                    else:
                                        # Esse bloco de for verifica as vendas nos SKUS errados e direciona para os corretos
                        
                                        dataExiste = False
                                        produto = self.todosProdutos.get(traducaoSku.upper().strip())
                                        
                                        if produto == None:
                                            logging.warning(f'SKU não encontrado: {traducaoSku}')                                            
                                        
                                        produto.vendasFull += float(quantidadeComprada)
                                        produto.faturado += float(valorVenda) * float(quantidadeComprada)
                                        for vendasData in produto.vendasPorDia:
                                            if vendasData["data"] == nota["notafiscal"]["dataEmissao"].split(' ')[0]:
                                                vendasData["vendas"] += float(quantidadeComprada)
                                                dataExiste = True
                                                break

                                        if dataExiste == False:
                                            objetoData = {
                                                "data": nota["notafiscal"]["dataEmissao"].split(' ')[0],
                                                "vendas": float(quantidadeComprada)
                                            }
                                            produto.vendasPorDia.append(objetoData)
                                    
                            else:
                            # Se a venda for de somente 1 item
                            
                                skuNota = listaItens["prod"]["cProd"].upper().strip()
                                quantidadeComprada = listaItens["prod"]["qCom"]
                                valorVenda = listaItens["prod"]["vUnCom"]
                                
                                traducaoSku = self.dadosErrados.get(skuNota)
                                produto = self.todosProdutos.get(skuNota)                                
                                
                                if produto == None and traducaoSku == None:
                                    logging.warning(f'SKU não encontrado: {skuNota}')                                    
                                
                                elif traducaoSku == None:
                                    
                                    dataExiste = False
                                    produto.vendasFull += float(quantidadeComprada)
                                    produto.faturado += float(valorVenda) * float(quantidadeComprada)
                                    # Componente
                                    if len(produto.composicao) > 0:
                                        for componente in produto.composicao:
                                            dataExiste = False
                                            produtoComposto = self.todosProdutos.get(componente["codigo"].upper().strip())
                                            
                                            if produtoComposto == None:
                                                logging.warning(f'SKU não encontrado: {componente["codigo"]}')
                                            
                                            # adciona vendas Full a variavel
                                            produtoComposto.vendasFull += float(quantidadeComprada) * float(componente["quantidade"])
                                            produtoComposto.faturado += float(valorVenda) * float(quantidadeComprada)
                                            for vendasData in produtoComposto.vendasPorDia:
                                                if vendasData["data"] == nota["notafiscal"]["dataEmissao"].split(' ')[0]:
                                                    vendasData["vendas"] += float(quantidadeComprada) * float(componente["quantidade"])
                                                    dataExiste = True
                                                    break

                                            if dataExiste == False:
                                                objetoData = {
                                                    "data": nota["notafiscal"]["dataEmissao"].split(' ')[0],
                                                    "vendas": float(quantidadeComprada) * float(
                                                        componente["quantidade"])
                                                }

                                                produtoComposto.vendasPorDia.append(objetoData)

                                    for vendasData in produto.vendasPorDia:
                                        if vendasData["data"] == nota["notafiscal"]["dataEmissao"].split(' ')[0]:
                                            vendasData["vendas"] += float(quantidadeComprada)
                                            dataExiste = True

                                    if dataExiste == False:
                                        objetoData = {
                                            "data": nota["notafiscal"]["dataEmissao"].split(' ')[0],
                                            "vendas": float(quantidadeComprada)
                                        }
                                        produto.vendasPorDia.append(objetoData)
                                    
                                else:
                                    #Caso o codigo esteja errado
                                        
                                    produto = self.todosProdutos.get(traducaoSku.upper().strip())
                                    
                                    if produto == None:
                                        logging.warning(f'SKU não encontrado: {traducaoSku}')                                  

                                    dataExiste = False                            
                                    produto.vendasFull += float(quantidadeComprada)
                                    produto.faturado += float(valorVenda) * float(quantidadeComprada)

                                    for vendasData in produto.vendasPorDia:
                                        if vendasData["data"] == nota["notafiscal"]["dataEmissao"].split(' ')[0]:
                                            vendasData["vendas"] += float(quantidadeComprada)
                                            dataExiste = True
                                            break

                                    if dataExiste == False:
                                        objetoData = {
                                            "data": nota["notafiscal"]["dataEmissao"].split(' ')[0],
                                            "vendas": float(quantidadeComprada)
                                        }
                                        produto.vendasPorDia.append(objetoData)

                        except Exception as e:
                            logging.critical(f'{e} nota: {nota["notafiscal"]["numero"]}')

            pagina += 1

    def calcularQuantidadeVendida(self, dataAnterior, dataAtual):
        
        page = 1
        print('Buscando os Produtos')
        self.buscarProdutos()
        print('Buscando Vendas')
        while True:
            sleep(.5)
            
            jsonPedidos = self.buscar_dados("pedidos", page, f'&filters=dataEmissao[{dataAnterior} TO {dataAtual}];idSituacao[9]')
            
            listaPedidos = dict(jsonPedidos["retorno"]).get("pedidos")
            if listaPedidos == None:
                break
            else:                    
                for pedido in listaPedidos:
                    for item in pedido["pedido"]["itens"]:
                        if item["item"]["codigo"] != None:
                            dataExiste = False
                            if item["item"]["codigo"].replace('', ' ') != '':

                                produto = self.todosProdutos.get(item["item"]["codigo"].upper().strip())
                                try:
                                    # Componente
                                    if len(produto.composicao) > 0:
                                        for componente in produto.composicao:

                                            produtoComposto = self.todosProdutos.get(componente["codigo"].upper().strip())
                                            
                                            for vendasData in produtoComposto.vendasPorDia:
                                                if vendasData["data"] == pedido["pedido"]["data"]:
                                                    vendasData["vendas"] += float(item["item"]["quantidade"]) * float(componente["quantidade"])
                                                    dataExiste = True

                                            if dataExiste == False:
                                                objetoData = {
                                                    "data": pedido["pedido"]["data"],
                                                    "vendas": float(item["item"]["quantidade"]) * float(componente["quantidade"])
                                                }
                                                produtoComposto.vendasPorDia.append(objetoData)

                                            produtoComposto.quantidadeVendida += float(item["item"]["quantidade"]) * float(componente["quantidade"])

                                            fator = (float(pedido["pedido"]["totalprodutos"]) - float(pedido["pedido"]["desconto"].replace(',', '.'))) / float(pedido["pedido"]["totalprodutos"])

                                            produtoComposto.faturado += (float(item["item"]["valorunidade"]) * fator * float(item["item"]["quantidade"]))

                                    for vendasData in produto.vendasPorDia:
                                        if vendasData["data"] == pedido["pedido"]["data"]:
                                            vendasData["vendas"] += float(item["item"]["quantidade"])
                                            dataExiste = True

                                    if dataExiste == False:
                                        objetoData = {
                                            "data": pedido["pedido"]["data"],
                                            "vendas": float(item["item"]["quantidade"])
                                        }
                                        produto.vendasPorDia.append(objetoData)

                                    produto.quantidadeVendida += float(item["item"]["quantidade"])

                                    fator = (float(pedido["pedido"]["totalprodutos"]) - float(pedido["pedido"]["desconto"].replace(',', '.'))) / float(pedido["pedido"]["totalprodutos"])
                                    produto.faturado += (float(item["item"]["valorunidade"]) * fator * float(item["item"]["quantidade"]))

                                except Exception as e:
                                    logging.critical(f'SKU not found: {item["item"]["codigo"]} - Error: {e}')

            page += 1

        print('Buscando Vendas Full')
        self.buscarNotasFiscais(dataAnterior, dataAtual)
        
        return self.todosProdutos

    def buscarPedidosCompra(self):
        print('Buscando Pedidos Compra')
        pagina = 1
        while True:
            sleep(.5)
            
            jsonPedidosCompra = self.buscar_dados("pedidoscompra", pagina, opcional="&filters=situacao[0]")

            listaPedidosCompra = dict(jsonPedidosCompra["retorno"]).get("pedidoscompra")

            if listaPedidosCompra == None:
                break

            for pedidoFeito in listaPedidosCompra:
                for pedCompFeito in pedidoFeito:

                    if ',' in pedCompFeito["pedidocompra"]["observacaointerna"]:
                        rastreiosLista = pedCompFeito["pedidocompra"]["observacaointerna"].split(',')
                    else:
                        rastreiosLista = pedCompFeito["pedidocompra"]["observacaointerna"].split(';')
                        
                    novoPedidoCompra = PedidoCompra(pedCompFeito["pedidocompra"]["numeropedido"],
                                                    pedCompFeito["pedidocompra"]["ordemcompra"],
                                                    pedCompFeito["pedidocompra"]["itens"], rastreiosLista,
                                                    pedCompFeito["pedidocompra"]["datacompra"])

                    self.pedidosCompra = append(self.pedidosCompra, novoPedidoCompra)
            pagina += 1

    def excluirComprasFinalizadas(self):
        linhasDeletadas = 0
        for ped in self.pedidosCompra:
            if ped.itemsChegaram == len(ped.listaItens):
                for linhaDeletar in ped.linhasDeletar:
                    
                    linhaDel = int(linhaDeletar) - linhasDeletadas
                    self.abaPedidosChegar.delete_rows(linhaDel)

                    # Diminuir 1 linha de todas as linhas pois a cada vez que deleta a planilha muda
                    linhasDeletadas += 1
                    
                self.atenderPedidoCompra(ped.numSistema)
                logging.info(f'Pedido atendido: {ped.numeroPedido}')

    # EXCEL FUNCTIONS

    def setarAbasPlanilha(self):
        self.abaDados = self.planilhaCompras.worksheet("dados")
        self.abaEstoque = self.planilhaCompras.worksheet("estoque")
        self.abaVendas = self.planilhaCompras.worksheet('vendas')
        self.abaRelatorio = self.planilhaCompras.worksheet('relatorio')
        self.abaRelatorioAux = self.planilhaCompras.worksheet('relatorioaux')
        self.abaPedidosChegar = self.conectarPlanilha('HAKON PARTS').worksheet('Pedidos')

    def preencherPlanihaDados(self):
            
        print('Preenchendo Aba Dados')
        linha = 2
        self.abaDados.update_cell(1, 1, "ESCREVENDO AQUI")
        self.abaDados.update_cell(1, 2, "ESCREVENDO AQUI")
        self.abaDados.update_cell(1, 3, "ESCREVENDO AQUI")
        self.abaDados.update_cell(1, 4, "ESCREVENDO AQUI")

        for sku, produto in self.todosProdutos.items():
            for vendaDia in produto.vendasPorDia:
                self.abaDados.update_cell(linha, 1, vendaDia["data"])
                sleep(2)
                self.abaDados.update_cell(linha, 2, vendaDia["vendas"])
                sleep(2)
                self.abaDados.update_cell(linha, 3, produto.sku)
                sleep(2)
                linha += 1

        self.abaDados.update_cell(linha, 1, "Acabou aqui")
        self.abaDados.update_cell(2, 4, f'{self.mesAtual}')

        self.abaDados.update_cell(1, 1, "DATA")
        self.abaDados.update_cell(1, 2, "VENDAS")
        self.abaDados.update_cell(1, 3, "SKU")
        self.abaDados.update_cell(1, 4, "ULTIMA ATUALIZAÇÃO")

    def preencherAbaPedidosCompra(self):
        
        print('Preenchendo Aba Pedido a Chegar')
        linha = 1
        listaPedidosPlanilha = []

        celulaEscrita = True
        # confere as linhas escritas

        df_compras = DataFrame(self.abaPedidosChegar.get_all_values())
        #Lista de códigos de pedidos
        listaCodigoPedido = df_compras[0]
        #Usa para ver se acabaram os pedidos
        listaValorProd = df_compras[5]
        #Identificar os pedidos completos
        listaItensChegaram = df_compras[4]

        cont = 0
        while celulaEscrita:
            codigoPedido = listaCodigoPedido[cont].replace(' ', '')
            verificarLinha = listaValorProd[cont].replace(' ', '')
            if verificarLinha == '' or verificarLinha == None:
                celulaEscrita = False
                break
            else:
                naLista = False        

                cedulaPedidoChegou = listaItensChegaram[cont]
                
                for pedidoCompra in self.pedidosCompra:
                    if pedidoCompra.numeroPedido == codigoPedido:
                        # valida se o item do pedido ja chegou completamente
                        if cedulaPedidoChegou == '0' or cedulaPedidoChegou == 0:
                            pedidoCompra.itemsChegaram += 1
                            pedidoCompra.linhasDeletar.append(linha)

                for codigoPed  in listaPedidosPlanilha:
                    if codigoPed == codigoPedido:
                        naLista = True
                
                if naLista == False:
                    listaPedidosPlanilha.append(codigoPedido)
            linha += 1
            cont += 1

        formatoPintar = {
                "backgroundColor": {
                "red": 99.0,
                "green": 99.0,
                "blue": 99.0
                },
                "horizontalAlignment": "RIGHT",
                "textFormat": {
                "foregroundColor": {
                    "red": 1.0,
                    "green": 1.0,
                    "blue": 1.0
                },
                "fontSize": 10,
                "bold": False
                }

            }

        for pedComp in self.pedidosCompra:

            pedidoNaPlanilha = False
            for peidoPlanilha in listaPedidosPlanilha:
                if peidoPlanilha == pedComp.numeroPedido:
                    pedidoNaPlanilha = True

            if pedidoNaPlanilha == False:
                logging.info(f'Pedido escrito: {pedComp.numeroPedido}')
                self.abaPedidosChegar.format(f'{linha}:{linha}', formatoPintar)

                self.abaPedidosChegar.update_cell(linha, 7, pedComp.dataCompra)
                sleep(2)
                coluna = 8
                for rastreio in pedComp.rastreios:
                    self.abaPedidosChegar.update_cell(linha, coluna, rastreio)
                    sleep(2)
                    coluna += 1

                for skuCompra in pedComp.listaItens:
                    self.abaPedidosChegar.update_cell(linha, 1, pedComp.numeroPedido)
                    sleep(3)
                    self.abaPedidosChegar.update_cell(linha, 2, skuCompra['item']["codigo"])
                    sleep(3)
                    self.abaPedidosChegar.update_cell(linha, 3, skuCompra['item']["qtde"])
                    sleep(3)
                    self.abaPedidosChegar.update_cell(linha, 6, skuCompra['item']["valor"])
                    sleep(3)

                    linha += 1

        #exclui os pedidos ja finalizados
        self.excluirComprasFinalizadas()

    def filtrarProdutos(self):
        self.produtosFiltrados = []
        for sku, produto in self.todosProdutos.items():
            if len(produto.composicao) == 0:
                self.produtosFiltrados.append(produto)

    def preenchePlanilhaEstoque(self):

        # Começa inserir os dados na planilha
        
        ('Preenchendo aba Estoque')
        linha = 2
        self.abaEstoque.update_cell(1, 1, "ESCREVENDO AQUI")
        self.abaEstoque.update_cell(1, 2, "ESCREVENDO AQUI")

        for produto in self.produtosFiltrados:
            self.abaEstoque.update_cell(linha, 1, produto.sku)
            sleep(2)
            self.abaEstoque.update_cell(linha, 2, produto.estoqueAtual)
            sleep(2)
            linha += 1
        
        self.abaEstoque.update_cell(1, 1, "SKU")
        self.abaEstoque.update_cell(1, 2, "FÍSICO")

    def preencherAbaVendas(self):
        print('Preenchendo aba Vendas')
        coluna = 2
        for produto in self.produtosFiltrados:
            self.abaVendas.update_cell(1, coluna, produto.sku)
            sleep(2)
            coluna += 1

    def preencherAbaRelatorio(self):
        linha = 3
        print("Preenchendo Aba Relatorio")        
        
        self.abaRelatorio.update_cell(1, 1, "ESCREVENDO AQUI")
        sleep(1)
        self.abaRelatorio.update_cell(1, 2, "ESCREVENDO AQUI")
        sleep(1)
        self.abaRelatorio.update_cell(1, 3, "ESCREVENDO AQUI")


        for sku, produto in self.todosProdutos.items():

            self.abaRelatorio.update_cell(linha, 1, produto.sku)
            sleep(2)
            self.abaRelatorio.update_cell(linha, 2, produto.estoqueAtual)
            sleep(1)
            self.abaRelatorio.update_cell(linha, 3, produto.quantidadeVendida)
            sleep(2)
            self.abaRelatorio.update_cell(linha, 4, produto.vendasFull)
            sleep(1)
            self.abaRelatorio.update_cell(linha, 5, float(produto.quantidadeVendida) + float(produto.vendasFull))
            sleep(2)
            self.abaRelatorio.update_cell(linha, 6, produto.faturado)
            sleep(1)
            linha += 1
        
        #Preenchendo headers
        self.abaRelatorio.update_cell(2, 1, "SKU")
        sleep(1)
        self.abaRelatorio.update_cell(2, 2, "ESTOQUE")
        sleep(1)
        self.abaRelatorio.update_cell(2, 3, "VENDAS FORA FULL 30D")
        sleep(1)
        self.abaRelatorio.update_cell(2, 4, "VENDAS FULL 30D")
        sleep(1)
        self.abaRelatorio.update_cell(2, 5, "TOTAL 30D")
        sleep(1)
        self.abaRelatorio.update_cell(2, 6, "TOTAL FATURADO 30D")
        sleep(1)
        self.abaRelatorio.update_cell(2, 7, "TICKET MÉDIO")

        self.abaRelatorio.update_cell(2, 20, self.mesAtual)
        linha = 2

    def preenchendoOutroMes(self, colunaComecar):
        
        linha = 2
        print("Preenchendo Aba Relatorio")

        for sku, produto in self.todosProdutos.items():

            coluna = colunaComecar
            
            self.abaRelatorioAux.update_cell(linha, coluna, produto.sku)
            sleep(2)
            coluna += 1
            self.abaRelatorioAux.update_cell(linha, coluna, float(produto.quantidadeVendida) + float(produto.vendasFull))
            sleep(2)
            coluna += 1
            self.abaRelatorioAux.update_cell(linha, coluna, produto.faturado)
            sleep(1)

            linha += 1
