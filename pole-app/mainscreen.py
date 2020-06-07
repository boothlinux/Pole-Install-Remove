import kivy
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.anchorlayout import AnchorLayout
from kivy.core.window import Window
import kivy  # Required to run Kivy such as the next line of code
kivy.require('1.9.1')  # used to alert user if this code is run on an earlier version of Kivy.
from kivy.app import App  # Imports the base App class required for Kivy Apps
from kivy.lang import Builder  # Imports the KV language builder that provides the layout of kivy screens
from kivy.uix.screenmanager import ScreenManager, Screen # Imports the Kivy Screen manager and Kivys Screen class
import os
import dropbox
from openpyxl import load_workbook, drawing
import openpyxl
import poleremoval

class WelcomeScreen(Screen):
    def polerempressed(self, instance):
        return PoleRemoval()
