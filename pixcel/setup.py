
from setuptools import setup

setup(
    name                = 'pixcel',
    version             = '0.1.0',
    description         = 'Excel to bitmap converter',
    author              = 'Michael Stella',
    author_email        = 'michael@thismetalsky.org',
    scripts             = [
        'pixcel',
    ],
    install_requires    = [
        'pillow',
        'webcolors',
        'openpyxl',
    ],
    python_requires     = '>=3.6.0',
    classifiers = [
        'Development Status :: 4 - Beta ',
        'Environment :: Console',
        'License :: Free for non-commercial use',
        'Natural Language :: English',
        'Operating System :: POSIX',
        'Programming Language :: Python :: 3',
        'Topic :: Utilities',
    ]
)
