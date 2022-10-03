class Produto:
    def __init__(self, sku, estoqueAtual, composicao=[], faturado=0):
        self.sku = sku
        self.estoqueAtual = estoqueAtual
        self.quantidadeVendida = 0
        self.vendasFull = 0
        self.vendasPorDia = []
        self.composicao = composicao
        self.faturado = faturado
