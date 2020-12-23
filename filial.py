class Filial:
    def __init__(self):
        """

        :var _nome: Nome da Filial
        :var _sigla: Sigla da Filial
        :var _producoes: Relação de Serial com Produção, de cada impressora da filial
        """

        self._nome = ''
        self._sigla = ''
        self._producoes = {}
        #self._producoesAnteriores = []

    @property
    def nome(self):
        return self._nome

    @nome.setter
    def nome(self, nome):
        self._nome = nome

    @property
    def sigla(self):
        return self._sigla

    @sigla.setter
    def sigla(self, sigla):
        self._sigla = sigla

    @property
    def serial(self):
        return self._serial

    @serial.setter
    def _serial(self, serial):
        self._serial = serial

    @property
    def producoes(self):
        return self._producoes

    @producoes.setter
    def producoes(self, producoes):
        self._producoes = producoes

    # @property
    # def _producoesAnteriores(self):
    #     return self._producoesAnteriores
    #
    # @_producoesAnteriores.setter
    # def _producoesAnteriores(self, producoes):
    #     self._producoesAnteriores = producoes
