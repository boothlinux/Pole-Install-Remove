from __future__ import unicode_literals
from textwrap import dedent

from plyer import filechooser

from kivy.app import App
from kivy.lang import Builder
from kivy.properties import ListProperty
from kivy.uix.button import Button
import filefile
import os
import openpyxl
import streetlightremoval
import streetlightinstall

def load(self, filename):
    stream = open(os.path.join(filename[0]), 'rb')
    img = openpyxl.drawing.image.Image(stream)
    img.anchor = 'A1'
    streetlightremoval.wb.create_sheet('Photo').add_image(img)
    stream.flush()

def toad(self, filename):
    stream = open(os.path.join(filename[0]), 'rb')
    img = openpyxl.drawing.image.Image(stream)
    img.anchor = 'A1'
    streetlightinstall.wb.create_sheet('Photo').add_image(img)
    stream.flush()
