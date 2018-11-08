import sys
from pyxl2json import parse_args
from pyxl2json.pyxl2json import main


argspec = parse_args(sys.argv[1:])
excel2json = main(argspec)

