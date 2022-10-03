from time import sleep

class PlanilhaSheets:
    
    def __init__(self, planilha):
        self.planilhaCompras = planilha

    def setarAbasPlanilha(self):
        
        self.abaDados = self.planilhaCompras.worksheet("dados")
        self.abaEstoque = self.planilhaCompras.worksheet("estoque")
        self.abaVendas = self.planilhaCompras.worksheet('vendas')
        self.abaRelatorio = self.planilhaCompras.worksheet('relatorio')
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

        df_compras = pd.DataFrame(self.abaPedidosChegar.get_all_values())
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
        linha = 2
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
        self.abaRelatorio.update_cell(1, 1, "SKU")
        sleep(1)
        self.abaRelatorio.update_cell(1, 2, "ESTOQUE")
        sleep(1)
        self.abaRelatorio.update_cell(1, 3, "VENDAS FORA FULL 30D")
        sleep(1)
        self.abaRelatorio.update_cell(1, 4, "VENDAS FULL 30D")
        sleep(1)
        self.abaRelatorio.update_cell(1, 5, "TOTAL 30D")
        sleep(1)
        self.abaRelatorio.update_cell(1, 6, "TOTAL FATURADO 30D")
        sleep(1)
        self.abaRelatorio.update_cell(1, 7, "TICKET MÉDIO")

        self.abaRelatorio.update_cell(2, 20, self.mesAtual)
        linha = 2

    def preenchendoOutroMes(self, colunaComecar):
        
        linha = 2
        print("Preenchendo Aba Relatorio")        


        for sku, produto in self.todosProdutos.items():

            coluna = colunaComecar
            
            self.abaRelatorio.update_cell(linha, coluna, produto.sku)
            sleep(2)
            coluna += 1
            self.abaRelatorio.update_cell(linha, coluna, float(produto.quantidadeVendida) + float(produto.vendasFull))
            sleep(2)
            coluna += 1
            self.abaRelatorio.update_cell(linha, coluna, produto.faturado)
            sleep(1)

            linha += 1
