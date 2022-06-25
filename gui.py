import sys
import os

from kivy.resources import resource_add_path, resource_find
from kivy.app import App
from kivy.uix.screenmanager import ScreenManager
from kivy.uix.screenmanager import Screen
from kivy.core.window import Window


class ScreenManagement(ScreenManager):
    def __init__(self, **kwargs):
        super(ScreenManagement, self).__init__(**kwargs)


class MainUi(Screen):
    def __init__(self, **kwargs):
        Window.bind(on_drop_file=self._on_file_drop)
        super(MainUi, self).__init__(**kwargs)

    def _on_file_drop(self, window, file_path, x, y):
        if self.parent.current == "Antismirnova":
            if self.ids.test.collide_point(*Window.mouse_pos):
                print(file_path.decode("utf-8"))


class AntismirnovaApp(App):
    def build(self):
        return ScreenManagement()


if __name__ == "__main__":
    if hasattr(sys, '_MEIPASS'):
        resource_add_path(os.path.join(sys._MEIPASS))
    AntismirnovaApp().run()
