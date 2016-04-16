import os
from datetime import datetime

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from conf import Config

CONFIG_PATH = 'reporter.cfg'


class CSVBackend(object):
    extension = 'csv'

    @staticmethod
    def read_table(fileobj):
        from csv import reader
        csv_reader = reader(fileobj)
        headers = csv_reader.next()
        return headers, csv_reader

BACKENDS = (CSVBackend,)


class Reporter(object):
    def __init__(self, config_path):
        self.config = Config(config_path)
        self.specs = self.config.get_spreadsheets()
        self.auth = self.config.get_auth()
        self.credentials = None
        self.client = None

    def initialize(self):
        scope, key_file = self.auth
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(key_file, [scope])
        self.client = gspread.authorize(self.credentials)

    def get_sheet_name(self, spec):
        return datetime.today().strftime("%m-%d-%y")

    def write_headers(self, sheet, headers):
        start_cell = sheet.get_addr_int(1, 1)
        end_cell = sheet.get_addr_int(1, len(headers))
        header_cells = sheet.range('{}:{}'.format(start_cell, end_cell))
        for cell, header in zip(header_cells, headers):
            cell.value = header
        sheet.update_cells(header_cells)

    def write_rows(self, sheet, rows, header_count):
        batch_size = 100
        batch = []
        row_count = 0
        start_row = 2

        def _commit():
            sheet.add_rows(row_count)
            print(
                sheet.get_addr_int(start_row, 1),
                sheet.get_addr_int(row_count + start_row - 1, header_count),
            )
            alphanum = '{}:{}'.format(
                sheet.get_addr_int(start_row, 1),
                sheet.get_addr_int(row_count + start_row - 1, header_count),
            )
            cells = sheet.range(alphanum)
            for cell, value in zip(cells, batch):
                cell.value = value
            sheet.update_cells(cells)

        for row in rows:
            row_count += 1
            batch.extend(row)
            if row_count != batch_size:
                continue

            # commit
            _commit()
            start_row += row_count
            row_count = 0
            batch = []

        if row_count:
            _commit()


    def _import_single(self, spec, fl):
        file_path = spec('file')
        ext = os.path.splitext(file_path)[1][1:]
        backend = next((backend for backend in BACKENDS if backend.extension == ext), None)
        if not backend:
            raise ValueError('Input source of format "{}" is not supported'.format(ext.upper()))

        headers, rows = backend.read_table(fl)
        gdoc = self.client.open_by_key(spec('key'))
        sheet = gdoc.add_worksheet(self.get_sheet_name(spec), 1, len(headers))
        self.write_headers(sheet, headers)
        self.write_rows(sheet, rows, len(headers))

    def import_all(self):
        for sheet_spec in self.specs:
            with open(sheet_spec('file', 'r')) as fl:
                self._import_single(sheet_spec, fl)


def main():
    config_path = CONFIG_PATH
    reporter = Reporter(config_path)
    reporter.initialize()
    reporter.import_all()

if __name__ == '__main__':
    main()
