import sys
import os

from kivy.resources import resource_add_path, resource_find
from kivy.app import App
from kivy.uix.screenmanager import ScreenManager
from kivy.uix.screenmanager import Screen
from kivy.uix.anchorlayout import AnchorLayout
from kivy.core.window import Window
from kivy.animation import Animation


class CustomButton(AnchorLayout):
    def on_press_animation(self):
        default_size = self.ids.button.size
        animation_size = (self.ids.button.size[0]*0.98, self.ids.button.size[1]*0.98)

        default_pos = self.ids.button.pos
        animation_pos = (self.ids.button.pos[0]+(default_size[0]-animation_size[0])/2,
                         self.ids.button.pos[1]+(default_size[1]-animation_size[1])/2)

        animation = Animation(size=animation_size, duration=0.1)
        animation += Animation(size=default_size, duration=0.1)

        animation1 = Animation(pos=animation_pos, duration=0.1)
        animation1 += Animation(pos=default_pos, duration=0.1)

        animation.start(self.canvas.before.get_group("button_background")[0])
        animation1.start(self.canvas.before.get_group("button_background")[0])

    def __init__(self, **kwargs):
        super(CustomButton, self).__init__(**kwargs)


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
