from kivy.app import App
# from kivy.uix.widget import Widget
# from kivy.properties import ObjectProperty
# from plyer import filechooser
# from main import runProcess
# import threading
# from kivy.clock import Clock
# import global_

# TODO: Send messaged to console output box


# class MainGrid(Widget):

# dataList = ObjectProperty(None)
# outputDir = ObjectProperty(None)
# outputConsole = ObjectProperty(None)

# listPath = ""
# outputPath = ""

# def updateConsole(self):
#     print("MAINGRID SAYS: ", global_.message)

# Clock.schedule_interval(updateConsole, 1)

# def btnOpenFile(self):
#     self.listPath = filechooser.open_file(title="Choose your list..", filters=[
#         ("Excel", "*.xlsx")])
#     self.ids.dataList.text = self.listPath[0]

# def btnSaveFile(self):
#     self.outputPath = filechooser.choose_dir(title="Where do you want to save the output?", filters=[
#         ("All Files", "*.*")])
#     self.ids.outputDir.text = self.outputPath[0]

# def btnGo(self):
#     runProcess(self, self.listPath, self.outputPath)

from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout


class LabelMaker(App):
    def build(self):

        label = Label(text="Hello World!", font_size=30)
        return BoxLayout()
        # return MainGrid()


app = LabelMaker()
app.run()

# processDataThread = threading.Thread(target=runProcess, daemon=True)
# updateConsoleThread = threading.Thread(
#     target=MainGrid.updateConsole, daemon=True)
# if __name__ == '__main__':
#     LabelMaker = LabelMaker()
#     LabelMaker.run()

# threading.Thread(target=LabelMaker.run())
# processDataThread.start()
# updateConsoleThread.start()
