from functools import partial

# Py2/3
try:
    from ConfigParser import ConfigParser, NoOptionError
except ImportError:
    from configparser import ConfigParser, NoOptionError


DEFAULT_CONFIG = {
    'auth': {
        'scope': 'https://spreadsheets.google.com/feeds',
    }
}


class Config(object):
    def __init__(self, path):
        self.parser = ConfigParser()
        with open(path, 'r') as fp:
            self.parser.readfp(fp)

    def get_key(self, section, key):
        if self.parser.has_option(section, key):
            return self.parser.get(section, key)
        # NoSectionError should be propageted, we only do safe look up for options.
        return DEFAULT_CONFIG[section].get(key)

    def get_auth(self):
        scope = self.get_key('auth', 'scope')
        key_file = self.get_key('auth', 'key_file')
        return scope, key_file

    def get_spreadsheets(self):
        specs = (self.get_key('core', 'spreadsheets') or '').strip().splitlines()
        spreadsheets = []
        for doc in specs:
            section = 'doc:{}'.format(doc)
            if not self.parser.has_section(section):
                raise ValueError('Spreadsheet section "{}" has not been described.'.format(section))
            spreadsheets.append(partial(self.parser.get, section))
        return spreadsheets
