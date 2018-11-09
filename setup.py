try: # for pip >= 10
    from pip._internal.req import parse_requirements
    from pip._internal.download import PipSession
except ImportError: # for pip <= 9.0.3
    from pip.req import parse_requirements
    from pip.download import PipSession

from pyxl2json import VERSION



################################################################################
# parse_requirements() returns generator of pip.req.InstallRequirement objects
install_reqs = parse_requirements("requirements.txt", session=PipSession())
from setuptools import setup, find_packages

# e.g. ['django==1.5.1', 'mezzanine==1.4.6']
reqs = [str(ir.req) for ir in install_reqs]


setup(
    name='pyxl2json',
    version=VERSION,
    description='Excel to Json',
    author='Duk-kyu Lim',
    author_email='hong18s@gmail.com',
    url='',
    download_url='',
    install_requires=reqs,
    packages=find_packages(include=['pyxl2json*', 'tests', ]),
    keywords=['excel', 'json', 'converter'],
    python_requires='>=3',
    package_data={},
    zip_safe=False,
    classifiers=[
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.2',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
    ],
    entry_points={
              'console_scripts': [
                  'pyxl2json=pyxl2json:__main__',
              ],
          },
    test_suite='discover_tests',
)