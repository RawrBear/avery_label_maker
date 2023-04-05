from kivy.app import App
from kivy.uix.widget import Widget
from kivy.properties import ObjectProperty
from plyer import filechooser
from main import runProcess

# TODO: Fire the popup and the runProcess functions from the PROCESS button


class MainGrid(Widget):
    dataList = ObjectProperty(None)
    outputDir = ObjectProperty(None)
    outputConsole = ObjectProperty(None)

    listPath = ""
    outputPath = ""

    def btnOpenFile(self):
        self.listPath = filechooser.open_file(title="Choose your list..", filters=[
            ("Excel", "*.xlsx")])
        self.ids.dataList.text = self.listPath[0]

    def btnSaveFile(self):
        self.outputPath = filechooser.choose_dir(title="Where do you want to save the output?", filters=[
            ("All Files", "*.*")])
        self.ids.outputDir.text = self.outputPath[0]

    def btnGo(self):

        print("OUTPUTPATH: ", self.outputPath[0])
        runProcess(self, self.listPath, self.outputPath)


class LabelMaker(App):
    def build(self):

        return MainGrid()


if __name__ == '__main__':
    LabelMaker().run()
