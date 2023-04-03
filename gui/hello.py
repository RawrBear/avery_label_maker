from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.widget import Widget


class MainGrid(Widget):
    pass


class LabelMaker(App):
    def build(self):

        return MainGrid()


if __name__ == '__main__':
    LabelMaker().run()
