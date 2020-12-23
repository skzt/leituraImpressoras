from string import ascii_uppercase


class ControlePoli:
    def __init__(self, configs, workbook):
        self._configs = configs
        self._workbook = workbook
        self._mapaSerial = {}  # [SERIAL] = (FILIAL, CELL)
        self._sheets = self._workbook.worksheets
        self._impSheet = self._sheets[self._configs.TOTAIS_IMP_SHEET]
        self._coluna_escrita = ascii_uppercase[(ascii_uppercase.index('E') - 1)]
        self._lista_nao_mexer = [f'<Worksheet "GYN-Totais">', f'<Worksheet "Imp">']

        # Varre todas as planilhas, mapeando a celula e a planilha de cada numero de serie.
        for sheet in self._sheets:

            if str(sheet) in self._lista_nao_mexer:
                continue
            for linha in range(self._configs.LINHA_INICIAL_CONTROLE, self._configs.LINHA_FINAL_CONTROLE):
                cell = self._configs.COLUNA_SERIAL + str(linha)
                valor = sheet[cell].value

                if valor is None or valor == "Serial":
                    continue

                escrita = self._coluna_escrita + str(linha)
                serial = valor
                filial = self._workbook.index(sheet)

                self._mapaSerial[serial] = (filial, escrita)

    def prepararPlanilha(self):
        # Mover os dados da coluna do meio, para esquerda
        # E da coluna da direita para a do meio.
        coluna_esquerda = ascii_uppercase[(ascii_uppercase.index(self._coluna_escrita) - 2)]
        coluna_meio = ascii_uppercase[(ascii_uppercase.index(self._coluna_escrita) - 1)]
        coluna_direita = self._coluna_escrita

        for sheet in self._sheets:
            if str(sheet) in self._lista_nao_mexer:
                continue
            for linha in range(self._configs.LINHA_INICIAL_CONTROLE, self._configs.LINHA_FINAL_CONTROLE):

                cellEsquerda = coluna_esquerda + str(linha)
                cellMeio = coluna_meio + str(linha)
                cellDireita = coluna_direita + str(linha)

                if isinstance(sheet[cellEsquerda].value, str) \
                        and sheet[cellEsquerda].value[0] == '=':
                    continue

                sheet[cellEsquerda].value = sheet[cellMeio].value
                sheet[cellMeio].value = sheet[cellDireita].value
                sheet[cellDireita].value = None

        # Move para esquerda o total de impressoras

        total_esquerda = self._impSheet[self._configs.COLUNA_ESQUERDA_TOTAIS_IMP['TOTAL']
                                        + str(self._configs.LINHA_INICIAL_TOTAIS_IMP)].value
        # Valida se a celula não está vazia
        total_esquerda = 0 if total_esquerda is None else total_esquerda

        total_meio = self._impSheet[self._configs.COLUNA_MEIO_TOTAIS_IMP['TOTAL']
                                    + str(self._configs.LINHA_INICIAL_TOTAIS_IMP)].value
        # Valida se a celula não está vazia
        total_meio = 0 if total_meio is None else total_meio

        total_direita = self._impSheet[self._configs.COLUNA_DIREITA_TOTAIS_IMP['TOTAL']
                                       + str(self._configs.LINHA_INICIAL_TOTAIS_IMP)].value
        # Valida se a celula não está vazia
        total_direita = 0 if total_direita is None else total_direita

        self.limpa_coluna(total_esquerda + self._configs.LINHA_INICIAL_TOTAIS_IMP,
                          list(self._configs.COLUNA_ESQUERDA_TOTAIS_IMP.values()))

        for linha in range(self._configs.LINHA_INICIAL_TOTAIS_IMP,
                           total_meio + self._configs.LINHA_INICIAL_TOTAIS_IMP):
            celula_esquerda_id = self._configs.COLUNA_ESQUERDA_TOTAIS_IMP['ID'] + str(linha)
            celula_esquerda_serial = self._configs.COLUNA_ESQUERDA_TOTAIS_IMP['SERIAL'] + str(linha)
            celula_esquerda_total = self._configs.COLUNA_ESQUERDA_TOTAIS_IMP['TOTAL'] + str(linha)

            celula_meio_id = self._configs.COLUNA_MEIO_TOTAIS_IMP['ID'] + str(linha)
            celula_meio_serial = self._configs.COLUNA_MEIO_TOTAIS_IMP['SERIAL'] + str(linha)
            celula_meio_total = self._configs.COLUNA_MEIO_TOTAIS_IMP['TOTAL'] + str(linha)

            self._impSheet[celula_esquerda_id].value = self._impSheet[celula_meio_id].value
            self._impSheet[celula_esquerda_serial].value = self._impSheet[celula_meio_serial].value
            self._impSheet[celula_esquerda_total].value = self._impSheet[celula_meio_total].value

            self._impSheet[celula_meio_id].value = None
            self._impSheet[celula_meio_serial].value = None
            self._impSheet[celula_meio_total].value = None

        for linha in range(self._configs.LINHA_INICIAL_TOTAIS_IMP,
                           total_direita + self._configs.LINHA_INICIAL_TOTAIS_IMP):
            celula_meio_id = self._configs.COLUNA_MEIO_TOTAIS_IMP['ID'] + str(linha)
            celula_meio_serial = self._configs.COLUNA_MEIO_TOTAIS_IMP['SERIAL'] + str(linha)
            celula_meio_total = self._configs.COLUNA_MEIO_TOTAIS_IMP['TOTAL'] + str(linha)

            celula_direita_id = self._configs.COLUNA_DIREITA_TOTAIS_IMP['ID'] + str(linha)
            celula_direita_serial = self._configs.COLUNA_DIREITA_TOTAIS_IMP['SERIAL'] + str(linha)
            celula_direita_total = self._configs.COLUNA_DIREITA_TOTAIS_IMP['TOTAL'] + str(linha)

            self._impSheet[celula_meio_id].value = self._impSheet[celula_direita_id].value
            self._impSheet[celula_meio_serial].value = self._impSheet[celula_direita_serial].value
            self._impSheet[celula_meio_total].value = self._impSheet[celula_direita_total].value

            self._impSheet[celula_direita_id].value = None
            self._impSheet[celula_direita_serial].value = None
            self._impSheet[celula_direita_total].value = None

        return True

    def limpa_coluna(self, linha_final, colunas):
        """
        Esvazia as colunas da self._configs.LINHA_INICIAL_TOTAIS_IMP a linha_final
        :param linha_final:
        :param colunas:
        :return:
        """
        for coluna in colunas:
            for linha in range(self._configs.LINHA_INICIAL_TOTAIS_IMP, linha_final):
                self._impSheet[coluna + str(linha)].value = None

    def escreverLeituras(self, leiturasDigital):
        # Itera sobrea uma lista de Filiais
        contador_impressoras = 0

        for filial in leiturasDigital.filiais:
            # Para cada Filial, itera sobre o serial em Filial.producoes
            for serial in filial.producoes:
                linha = contador_impressoras + self._configs.LINHA_INICIAL_TOTAIS_IMP
                celula_direita_id = self._configs.COLUNA_DIREITA_TOTAIS_IMP['ID'] + str(linha)
                celula_direita_serial = self._configs.COLUNA_DIREITA_TOTAIS_IMP['SERIAL'] + str(linha)
                try:

                    # Identifica o local na planilha onde deve gravar os dados
                    local = self._mapaSerial[serial]
                except KeyError:
                    return [filial.nome, filial.sigla, serial]

                # Coloca os dados na planilha
                self._sheets[local[0]][local[1]].value = filial.producoes[serial]
                # Anota o serial e incrementa o contador de impressoras
                self._impSheet[celula_direita_id].value = 1 + contador_impressoras  # soma 1 para ID iniciar do 1.
                self._impSheet[celula_direita_serial].value = serial
                contador_impressoras += 1

        linha = self._configs.LINHA_INICIAL_TOTAIS_IMP
        celula_direita_total = self._configs.COLUNA_DIREITA_TOTAIS_IMP['TOTAL'] + str(linha)
        self._impSheet[celula_direita_total].value = contador_impressoras  # já foi somado 1 no ultimo loop do for.
        return True
