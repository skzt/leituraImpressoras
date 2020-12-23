import json
from os import stat


class Configs:
    def __init__(self):
        try:
            if stat("config.json").st_size == 0:
                self.createDefaultFile()

        except FileNotFoundError:
            self.createDefaultFile()

        finally:

            file = open("config.json", 'r', encoding='utf-8')
            try:
                configFile = json.load(file)
            except json.decoder.JSONDecodeError:
                file.close()
                self.createDefaultFile()
                file = open("config.json", 'r', encoding='utf-8')
                configFile = json.load(file)

            self._siglas = configFile["SIGLAS"]
            self._totais_imp_sheet = configFile["TOTAIS_IMP_SHEET"]

            self._linha_inicial_digital = configFile["LINHA_INICIAL_DIGITAL"]
            self._linha_inicial_controle = configFile["LINHA_INICIAL_CONTROLE"]
            self._linha_final_controle = configFile["LINHA_FINAL_CONTROLE"]
            self._linha_inicial_totais_imp = configFile["LINHA_INICIAL_TOTAIS_IMP"]

            self._coluna_producoes = configFile["COLUNA_PRODUCOES"]
            self._coluna_serial = configFile["COLUNA_SERIAL"]
            self._coluna_dpto = configFile["COLUNA_DPTO"]
            self._coluna_esquereda_totais_imp = configFile["COLUNA_ESQUERDA_TOTAIS_IMP"]
            self._coluna_meio_totais_imp = configFile["COLUNA_MEIO_TOTAIS_IMP"]
            self._coluna_direita_totais_imp = configFile["COLUNA_DIREITA_TOTAIS_IMP"]

            self._modelo_impressoras = configFile["MODELO_IMPRESSORAS"]

            file.close()

    @staticmethod
    def createDefaultFile():
        configFile = open("config.json", 'w', encoding='utf-8')
        defaultFile = {
            "SIGLAS": {
                "ARAGUAINA": "ARG",
                "BAURU": "BAU",
                "BELEM": "BEL",
                "BELO HORIZONTE": "BHZ",
                "BRASÍLIA": "BSB",
                "CUIABÁ": "CBA",
                "CAMPO GRANDE": "CGR",
                "GOIÂNIA": "GYN",
                "LONDRINA": "LON",
                "PORTO ALEGRE": "POA",
                "RIBEIRÃO PRETO": "RPT",
                "SÃO LUIZ": "SLZ",
                "SÃO PAULO": "SPO",
                "UBERABA": "URA",
                "VITÓRIA": "VIT"
            },
            "LINHA_INICIAL_DIGITAL": 12,
            "LINHA_INICIAL_CONTROLE": 21,
            "LINHA_FINAL_CONTROLE": 30,
            "COLUNA_PRODUCOES": "E",
            "COLUNA_SERIAL": "E",
            "COLUNA_DPTO": "A",
            "MODELO_IMPRESSORAS": [
                "8952",
                "8912",
                "6182",
                "8157",
                "5200",
                "7460",
                "5452",
                "C711",
                "5652",
                "6202"
            ]
        }

        json.dump(defaultFile, configFile, ensure_ascii=False, indent=4)
        configFile.close()

    @property
    def SIGLAS(self):
        return self._siglas

    @property
    def TOTAIS_IMP_SHEET(self):
        return self._totais_imp_sheet

    @property
    def LINHA_INICIAL_DIGITAL(self):
        return self._linha_inicial_digital

    @property
    def LINHA_INICIAL_CONTROLE(self):
        return self._linha_inicial_controle

    @property
    def LINHA_INICIAL_TOTAIS_IMP(self):
        return self._linha_inicial_totais_imp

    @property
    def LINHA_FINAL_CONTROLE(self):
        return self._linha_final_controle

    @property
    def COLUNA_PRODUCOES(self):
        return self._coluna_producoes

    @property
    def COLUNA_SERIAL(self):
        return self._coluna_serial

    @property
    def COLUNA_DPTO(self):
        return self._coluna_dpto

    @property
    def COLUNA_ESQUERDA_TOTAIS_IMP(self):
        return self._coluna_esquereda_totais_imp

    @property
    def COLUNA_MEIO_TOTAIS_IMP(self):
        return self._coluna_meio_totais_imp

    @property
    def COLUNA_DIREITA_TOTAIS_IMP(self):
        return self._coluna_direita_totais_imp

    @property
    def MODELO_IMPRESSORAS(self):
        return self._modelo_impressoras
