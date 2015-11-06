from distutils.core import setup
import py2exe
import os

absolute_filename = os.path.dirname(os.path.abspath(__file__)) + "\magazi.py"

setup(windows=[absolute_filename])
#setup(console=['C:/Users/George/Desktop/magazi1.py'])