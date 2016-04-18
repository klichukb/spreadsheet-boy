#! -*- encoding: utf-8 -*-

from __future__ import unicode_literals

import os
from datetime import datetime
from logging import getLogger

import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

from conf import Config

BACKENDS = {}

logger = getLogger('spreadsheets.reporter')


def backend(extension):
    def wrapper(func):
        BACKENDS[extension] = func
        return func
    return wrapper

@backend('csv')
def read_table(fileobj):
    from csv import reader
    csv_reader = reader(fileobj)
    headers = csv_reader.next()
    return headers, csv_reader

    
class Reporter(object):
    def __init__(self, config):
        self.config = config
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
        batch_size = 1000
        batch = []
        row_count = 0
        start_row = 2
        available = sheet.row_count - 1

        def _commit(available):
            if available < row_count:
                sheet.add_rows(row_count - available)
                available = 0
            else:
                available -= row_count
            alphanum = '{}:{}'.format(
                sheet.get_addr_int(start_row, 1),
                sheet.get_addr_int(row_count + start_row - 1, header_count),
            )
            cells = sheet.range(alphanum)
            for cell, value in zip(cells, batch):
                cell.value = value
            sheet.update_cells(cells)
            return available

        for row in rows:
            row_count += 1
            batch.extend(row)
            if row_count != batch_size:
                continue

            # commit
            available = _commit(available)
            start_row += row_count
            row_count = 0
            batch = []

        if row_count:
            available = _commit(available) 
            
    def upload(self, spec_name, update=False):
        create = not update
        spec = self.specs.get(spec_name)
        if not spec:
            raise KeyError('Spreadsheet "{}" is not configured'.format(spec_name))
        logger.info('Uploading "{}"...'.format(spec_name))
            
        file_path = spec('file')
        sheet_name = self.get_sheet_name(spec)
        ext = os.path.splitext(file_path)[1][1:]
        backend = BACKENDS.get(ext)
        
        if not backend:
            raise ValueError('Input source of format "{}" is not supported'.format(ext.upper()))
            
        gdoc = self.client.open_by_key(spec('key'))
 
        with open(file_path, 'r') as fileobj:
            headers, rows = backend(fileobj)
            sheet = None
            if not create:
                try:
                    sheet = gdoc.worksheet(sheet_name)
                except WorksheetNotFound:
                    sheet = None
            create = not sheet
            if create:
                sheet = gdoc.add_worksheet(sheet_name, 1, len(headers))
            self.write_headers(sheet, headers)
            self.write_rows(sheet, rows, len(headers))
           
        logger.info('Successfully uploaded "{}": {} sheet "{}".'.format(
            spec_name, 'created' if create else 'updated', sheet_name))