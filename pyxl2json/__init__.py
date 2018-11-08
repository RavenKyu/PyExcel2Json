# Version
NUM_VERSION = (1, 0, 0)
VERSION = ".".join(str(nv) for nv in NUM_VERSION)
__version__ = VERSION

import sys
import re
from argparse import ArgumentParser


################################################################################
class ArgsExit(Exception):
    pass


################################################################################
class ArgsError(Exception):
    pass


################################################################################
def re_match(s, p):
    m = re.match(p, s)
    return m is not None


################################################################################
class ExtArgParser(ArgumentParser):
    EXTENDED_ATTRS = {'re_match': re_match}

    def __init__(self, *args, **kwargs):
        self._ext_argspec = {}
        super(ExtArgParser, self).__init__(*args, **kwargs)


    # ==========================================================================
    def add_argument(self, *args, **kwargs):
        ext_kwargs = {}
        for ea in self.EXTENDED_ATTRS.keys():
            if ea in kwargs:
                ext_kwargs[ea] = kwargs[ea]
                del kwargs[ea]
        action = ArgumentParser.add_argument(self, *args, **kwargs)
        self._ext_argspec[action.dest] = ext_kwargs
        return action

    # ==========================================================================
    def parse_args(self, args=None, namespace=None):
        if args is not None and not args:
            args = None
        elif isinstance(args, (list, tuple)):
            args = list(map(str, args))
            # 에러 없이 돌리고
        self._args_error = None
        try:
            self._args = ArgumentParser.parse_args(self, args, namespace)
        except Exception as _:
            raise ArgsError('%s' % self._args_error)
        self._args = ArgumentParser.parse_args(self, args, namespace)
        # EXTENDED_ATTRS 에 대한 검증
        for dest, ext_args in self._ext_argspec.items():
            av = getattr(self._args, dest, None)
            if av is None:
                continue
            for ext_att, cv in ext_args.items():
                # None은 비교대상이 아님
                if cv is None:
                    continue
                if not isinstance(av, list):
                    av = [av]
                for iav in av:
                    if not isinstance(iav, type(cv)):
                        raise ValueError('%s must be type(%s) for %s'
                                         % (cv, type(iav).__name__, dest))
                    if not self.EXTENDED_ATTRS[ext_att](iav, cv):
                        raise ArgsError('For Argument "%s", "%s" validatation '
                                        'error: user input is "%s" but rule is '
                                        '"%s"' % (dest, ext_att, iav, cv))
        return self._args
    # ==========================================================================
    def __enter__(self):
        return self

    # ==========================================================================
    def __exit__(self, type, value, traceback):
        return True


################################################################################
def parse_args(*argv):
    parser = ExtArgParser(
            prog=sys.argv[0],
            description='PyExcel2Json is easy to convert Excel to Json')
    parser.add_argument('excel_filename', help='')
    parser.add_argument('--head', '-t',
                        re_match='^\w+\d+:\w+\d+$', help='A1:F1')
    parser.add_argument('--data', '-d',
                        re_match='^\w+\d+:\w+\d+$', help='A2:F10')
    parser.add_argument('--sheet', '-s', default='Sheet1',
                        type=str, help='SHEET NAME', )
    r = parser.parse_args(*argv)
    return r

