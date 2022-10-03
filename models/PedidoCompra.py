class PedidoCompra:
    def __init__(self, numSistema, numPedido, listaItens, rastreios, data):
        self.numSistema = numSistema
        self.numeroPedido = numPedido
        self.listaItens = listaItens
        self.rastreios = rastreios
        self.dataCompra = data
        self.itemsChegaram = 0
        self.jaChegou = False
        self.linhasDeletar = []
