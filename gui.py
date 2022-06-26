import pathlib
import sys
import os
import time

from kivy.clock import Clock
from kivy.resources import resource_add_path, resource_find
from kivy.app import App
from kivy.uix.screenmanager import ScreenManager
from kivy.uix.screenmanager import Screen
from kivy.core.window import Window
from kivy.animation import Animation
from kivy.uix.textinput import TextInput
from kivy.uix.anchorlayout import AnchorLayout
from kivy.properties import VariableListProperty

import word


class CustomTextInput(AnchorLayout):
    bg_color = VariableListProperty([0, 0, 0, 0])
    radius = VariableListProperty([0, 0, 0, 0])

    @property
    def text(self):
        return self.__text_input.text

    @text.setter
    def text(self, value):
        self.__text_input.text = value

    def __init__(self, **kwargs):
        super(CustomTextInput, self).__init__(**kwargs)

        self.__text_input = TextInput(
            background_color=(0, 0, 0, 0),
            size_hint=(1, None),
            size=(self.size[0], 30),
            multiline=False,
            write_tab=False
        )

        self.add_widget(self.__text_input)


class ScreenManagement(ScreenManager):
    def __init__(self, **kwargs):
        super(ScreenManagement, self).__init__(**kwargs)


class MainUi(Screen):
    def __init__(self, **kwargs):
        Window.bind(on_drop_file=self._on_file_drop)
        super(MainUi, self).__init__(**kwargs)

        self.current_working_file = None

        self.default_drag_and_drop_label_color = None
        self.drag_and_drop_text_changing = None

    def change_drag_and_drop_label_text(self, period, new_text):
        if self.drag_and_drop_text_changing:
            return
        self.drag_and_drop_text_changing = True

        label = self.ids.drag_and_drop_label
        old_text, label.text = label.text, new_text

        def set_default(text):
            label.text = old_text
            self.drag_and_drop_text_changing = False

        Clock.schedule_once(set_default, period)

    def invalid_file_animation(self, *args):
        drag_and_drop_label = self.ids.drag_and_drop_label
        canvas_color = drag_and_drop_label.canvas.before.get_group("color")[0]
        canvas = drag_and_drop_label.canvas.before.get_group("rounded_rectangle")[0]

        if self.default_drag_and_drop_label_color is None:
            self.default_drag_and_drop_label_color = tuple(canvas_color.rgba)

        color_animation = Animation(rgba=(255, 0, 0, 0.6), duration=0.0)
        color_animation += Animation(rgba=(255, 0, 0, 0.6), duration=0.25)
        color_animation += Animation(rgba=self.default_drag_and_drop_label_color, duration=100)

        print(self.default_drag_and_drop_label_color)
        Animation.cancel_all(canvas, "pos")

        animation = Animation(pos=(drag_and_drop_label.pos[0] - 2, drag_and_drop_label.pos[1]), d=0.05)
        animation += Animation(pos=(drag_and_drop_label.pos[0] + 4, drag_and_drop_label.pos[1]), d=0.05)
        animation += Animation(pos=(drag_and_drop_label.pos[0] - 4, drag_and_drop_label.pos[1]), d=0.05)
        animation += Animation(pos=(drag_and_drop_label.pos[0] + 4, drag_and_drop_label.pos[1]), d=0.05)
        animation += Animation(pos=(drag_and_drop_label.pos[0], drag_and_drop_label.pos[1]), d=0.05)

        color_animation.start(canvas_color)
        animation.start(canvas)

        self.change_drag_and_drop_label_text(period=1, new_text="File is not a word file")

    def __initialize_file(self, file: pathlib.Path):
        if word.is_word_file(file) is False:
            self.invalid_file_animation()
            return

        self.current_working_file = file
        metadata = word.Metadata(self.current_working_file)

        if (creator := metadata.creator) is not None:
            self.ids.creator_text_input.text = creator
        else:
            self.ids.creator_text_input.text = ""

        if (last_modified_by := metadata.last_modified_by) is not None:
            self.ids.last_modified_by_text_input.text = last_modified_by
        else:
            self.ids.last_modified_by_text_input.text = ""

        if (revision := metadata.revision) is not None:
            self.ids.revision_text_input.text = str(revision)
        else:
            self.ids.revision_text_input.text = ""

        if (application_name := metadata.application_name) is not None:
            self.ids.application_text_input.text = application_name
        else:
            self.ids.application_text_input.text = ""

        if (editing_time := metadata.editing_time) is not None:
            self.ids.editing_time_text_input.text = str(editing_time)
        else:
            self.ids.editing_time_text_input.text = ""

    def _on_file_drop(self, window, file_path, x, y):
        if self.parent.current == "Antismirnova":
            if self.ids.drag_and_drop_label.collide_point(*Window.mouse_pos):
                self.__initialize_file(pathlib.Path(file_path.decode("utf-8")))


class AntismirnovaApp(App):
    def build(self):
        return ScreenManagement()


if __name__ == "__main__":
    if hasattr(sys, '_MEIPASS'):
        resource_add_path(os.path.join(sys._MEIPASS))
    AntismirnovaApp().run()
