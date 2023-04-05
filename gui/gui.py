from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.widget import Widget
from kivy.properties import ObjectProperty
from kivy.uix.popup import Popup
from plyer import filechooser

# TODO: Figure out a way to send the path[0] to the "df" var in the main.py
# TODO: Link that to the main button in the gui so that it fires the loop


class MainGrid(Widget):
    dataList = ObjectProperty(None)
    outputDir = ObjectProperty(None)

    def btnOpenFile(self):

        path = filechooser.open_file(title="Where do you want to save the output?", filters=[
            ("All Files", "*.*"),])
        self.ids.dataList.text = path[0]
        print(path[0])

    def btnSaveFile(self):
        path = filechooser.choose_dir(title="Choose your list..", filters=[
            ("Comma-separated Values", "*.csv"), ("Excel", "*.xlsx")])
        self.ids.outputDir.text = path[0]
        print(path)

    def btnGo(self):
        self.ids.dataList.text = 'Input'
        self.ids.outputDir.text = 'Output'


class LabelMaker(App):
    def build(self):

        return MainGrid()


if __name__ == '__main__':
    LabelMaker().run()
