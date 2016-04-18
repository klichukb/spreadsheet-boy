#! /bin/env python

import sys
from argparse import ArgumentParser
import logging

from spreadsheet_boy.reporter import Reporter
from spreadsheet_boy.conf import Config

logger = logging.getLogger('spreadsheets')

CONFIG_PATH = 'reporter.cfg'

def get_parser():
    parser = ArgumentParser(prog='PROD')
    parser.add_argument('--doc', nargs='*', help='spreadsheets to upload')
    parser.add_argument('--config', default=CONFIG_PATH, help='config file path')
    return parser

def main():
    # read configuration
    parser = get_parser()
    args = parser.parse_args(sys.argv[1:])
    config = Config(args.config)
    
    # configure logging
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s: %(message)s'))  
    logger.addHandler(handler)
    log_level = config.get_key('app', 'log_level')
    if log_level:
        logger.setLevel(getattr(logging, log_level))
    
    # get things done
    reporter = Reporter(config)
    reporter.initialize()
    for doc in args.doc or reporter.specs.keys():
        reporter.upload(doc, update=True)

if __name__ == '__main__':
    main()

