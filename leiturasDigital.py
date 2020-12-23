from filial import Filial


class LeiturasDigital:
    def __init__(self, configs, planilha):
        self._configs = configs
        self.linha = self._configs.LINHA_INICIAL_DIGITAL
        self._planilha = planilha
        self._filiais = []

        filial = None
        celulaBranco = 0  # Quantidade de celulas em branco consecutivas

        # Repita até que aconteça 20 celulas em branco consecutivas
        fp = open("imp.txt", "w")
        while celulaBranco < 20:
            cell = 'A' + str(self.linha)
            valueCell = self._planilha[cell].value

            # Verifica se existe e remove espaço no inicio da celula
            if valueCell is not None and valueCell[0] == ' ':
                valueCell = valueCell[1:]

            if valueCell is None:
                # Celula Vazia
                celulaBranco += 1

            elif valueCell[:4] in self._configs.MODELO_IMPRESSORAS:
                # Linha das produções

                celulaBranco = 0

                if '-' not in valueCell[7:]:
                    valueCell = valueCell[:7] + valueCell[7:].replace(' ', ' - ')

                serial = valueCell.split('-')[1].strip()
                if serial.isalnum() is False:
                    serial = serial[:serial.find('(')].strip()

                producao = self._planilha[(self._configs.COLUNA_PRODUCOES + str(self.linha))].value

                filial.producoes[serial] = producao
            elif self._planilha[cell].value[:3] == "COD":
                # Linha CNPJ -> Ignorar
                celulaBranco = 0

            elif valueCell[:9].rstrip() == "POLIPECAS":
                # Linha Nome Filial -> Criar nova filial na lista
                celulaBranco = 0

                filial = Filial()
                filial.nome = valueCell.split(' - ')[1]
                filial.sigla = self._configs.SIGLAS[filial.nome]
                self._filiais.append(filial)
            self.linha += 1
    @property
    def filiais(self):
        return self._filiais
