from setuptools import setup

setup(
    name='pyOutLook',
    version='0.1',
    packages=['pyOutLook', 'lib', 'src'],
    install_requires=['pywin32'],
    url='https://github.com/mani-src/pyOutLook.git',
    license='GNU General Public License Version 3',
    author='Manikanta Ambadipudi',
    author_email='ambadipudi.manikanta@gmail.com',
    description='Library and command line application for Outlook interface',
)
