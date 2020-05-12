import logging
import pandas as pd
import re

from wordfile import WordDoc

#Tests
from wordfile3cols import Word3Cols

LOGGER = logging.getLogger(__name__)
LOGGER.setLevel(logging.DEBUG)

MONTH = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro"
}


class ExcelInfo:

    def __init__(self, filepath):
        self.excluded_sheets = ['INICIO', 'Indice', 'Contas,Carlos', 'TESTE']
        self.saved = None

        try:
            self.df = pd.ExcelFile(filepath)
        except Exception as exp:
            LOGGER.error("Error reading excel file: ", exp)

    def get(self):
        clients_list = []

        for x in self.df.sheet_names:
            if x not in self.excluded_sheets:
                client = Cliente(self.df.parse(x).to_dict())
                clients_list.append(client)
            else:
                LOGGER.info(f'Sheet {x} passed')

        return clients_list


class Cliente:

    def __init__(self, data):
        self._data = data

    @property
    def nome(self):
        try:
            return list(self._data.keys())[0].strip().title()
        except KeyError:
            LOGGER.error("Key not find")

    @property
    def raw_dias(self):
        try:
            return self._data.get('Nº DIAS')
        except KeyError:
            LOGGER.error("Key not find")

    @property
    def raw_valor_diario(self):
        try:
            return self._data.get('VALOR DIÁRIO')
        except KeyError:
            LOGGER.error("Key not find")

    @property
    def raw_valor_extra(self):
        try:
            return self._data.get(' VALOR EX')
        except KeyError:
            LOGGER.error("Key not find")

    @property
    def raw_valor_total(self):
        try:
            return self._data.get('TOTAL')
        except KeyError:
            LOGGER.error("Key not find")

    @property
    def raw_data_pagamento(self):
        try:
            return self._data.get('D/A')
        except KeyError:
            LOGGER.error("Key not find")

    @property
    def raw_observacoes(self):
        try:
            return self._data.get('OBSERVAÇÕES')
        except KeyError:
            LOGGER.error("Key not find")

    def month_verify(self, month):
        months_dict = {}
        for it in range(month):
            real_month = it + 1
            if pd.isna(self.raw_data_pagamento.get(real_month)):
                if self.raw_valor_total.get(real_month) != 0:
                    months_dict[MONTH[real_month]] = round(self.raw_valor_total.get(real_month), 2)
        return months_dict

    def month_total(self, month):
        """Total month to pay"""
        debt = []
        for it in range(month):
            real_month = it + 1
            if pd.isna(self.raw_data_pagamento.get(real_month)):
                debt.append(self.raw_valor_total.get(real_month))
        if not debt or sum(debt) == 0:
            return f"Todos os meses pagos até {MONTH[month]}."
        return "Total - " + str(round(sum(debt), 2)) + "€"

    def __repr__(self):
        return self.nome


if __name__ == '__main__':
    info = ExcelInfo('/home/joaocps/git-reps/experiments/excel_pd/updated.xls')
    clients = info.get()

    """Get month total"""
    for client in clients:
        client.month_total(3)

    # tr = WordDoc(clients,4)
    # tr.create()

    tr = Word3Cols(clients, 4)
    tr.create()
