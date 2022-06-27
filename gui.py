import pathlib
import sys
import os
import time
from typing import NamedTuple

from kivy import utils, Config
from kivy.clock import Clock
from kivy.graphics import RoundedRectangle, Color
from kivy.resources import resource_add_path, resource_find
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.image import Image
from kivy.uix.label import Label
from kivy.uix.screenmanager import ScreenManager
from kivy.uix.screenmanager import Screen
from kivy.core.window import Window
from kivy.animation import Animation
from kivy.uix.textinput import TextInput
from kivy.uix.anchorlayout import AnchorLayout
from kivy.properties import VariableListProperty, StringProperty, AliasProperty

from textwrap import wrap

import word


def label_text_wrap(text: str, label_width, label) -> str:
    if label.texture_size[0] == 0:
        return text

    if label.texture_size[0] != label_width:
        max_line_length = len((lines := label.text.split("\n"))[0])
        for line in lines:
            if (new_len := len(line)) > max_line_length:
                max_line_length = new_len

        char_width = label.texture_size[0] / max_line_length
        wrapped_text = '\n'.join(wrap(text, int(label_width / char_width)))
        return wrapped_text
    else:
        return text


class FileDragAndDropperStateLabel(BoxLayout):
    bg_color = VariableListProperty([0, 0, 0, 0])
    drag_and_drop_label_text = StringProperty("Drag & drop your file here")

    def __init__(self, **kwargs):
        super(FileDragAndDropperStateLabel, self).__init__(**kwargs)
        self.default_drag_and_drop_label_color = None


class FileDragAndDropperStateInfo(BoxLayout):
    bg_color = VariableListProperty([0, 0, 0, 0])
    drag_and_drop_label_text = StringProperty("Drag & drop your file here")
    image_source = StringProperty("")
    info_label_text = StringProperty("")
    drag_and_drop_label_bg_color = VariableListProperty([0, 0, 0, 0])

    def __init__(self, **kwargs):
        super(FileDragAndDropperStateInfo, self).__init__(**kwargs)
        self.default_drag_and_drop_label_color = None


class WidgetChangeAnimation:
    def new_widget_add(self):
        self.parent_widget.remove_widget(self.old_widget)
        self.parent_widget.add_widget(self.new_widget)
        self.new_widget_animation()

    def old_widget_animation(self):
        animation = Animation(
            opacity=0,
            duration=0.3
        )
        animation.bind(on_complete=lambda *_: self.new_widget_add())  # )
        animation.start(self.old_widget)

    def new_widget_animation(self):
        animation = Animation(
            opacity=1,
            duration=0.3
        )
        animation.start(self.new_widget)

    def start(self):
        self.old_widget_animation()

    def __init__(self, parent_widget, old_widget, new_widget):
        self.parent_widget = parent_widget
        self.old_widget = old_widget
        self.new_widget = new_widget

        new_widget.opacity = 0


class WidgetChangePropertiesAnimation:
    def opacity_up_animation(self):
        animation = Animation(
            opacity=1,
            duration=0.3
        )
        animation.start(self.widget)

    def set_properties(self):
        self.setter()
        self.opacity_up_animation()

    def opacity_down_animation(self):
        animation = Animation(
            opacity=0,
            duration=0.3
        )
        animation.bind(on_complete=lambda *_: self.set_properties())
        animation.start(self.widget)

    def start(self):
        self.opacity_down_animation()

    def __init__(self, widget, setter: callable):
        self.widget = widget
        self.setter = setter


class FileDragAndDropper(BoxLayout):
    def update_info_widget(self):
        self.__state_info_box_layout.info_label_text = f'Word file "{self.current_working_file.name}"'
        self.__state_info_box_layout.image_source = "images/word_icon.png"
        self.__state_info_box_layout.drag_and_drop_label_bg_color = utils.get_color_from_hex("5ec6ff")

    def set_state(self, state):
        match state, self.current_state:
            case str("label"), str("info") | None:
                self.add_widget(self.__state_label_box_layout)
                self.current_state = "label"
            case str("info"), str("label") | None:
                self.__state_info_box_layout.info_label_text = f'Word file "{self.current_working_file.name}"'
                self.__state_info_box_layout.image_source = "images/word_icon.png"
                self.__state_info_box_layout.drag_and_drop_label_bg_color = utils.get_color_from_hex("5ec6ff")
                self.__state_info_box_layout.opacity = 0

                change_widget_animation = WidgetChangeAnimation(
                    parent_widget=self,
                    old_widget=self.__state_label_box_layout,
                    new_widget=self.__state_info_box_layout,
                )
                change_widget_animation.start()

                self.current_state = "info"
            case str("info"), str("info"):
                properties_animation = WidgetChangePropertiesAnimation(
                    self.__state_info_box_layout,
                    self.update_info_widget
                )
                properties_animation.start()

    def __init__(self, **kwargs):
        Window.bind(on_drop_file=self._on_file_drop)

        self.creator_text_input = None
        self.last_modified_by_text_input = None
        self.revision_text_input = None
        self.application_text_input = None
        self.editing_time_text_input = None

        super(FileDragAndDropper, self).__init__(**kwargs)

        self.current_state = None
        self.current_working_file = None
        self.default_drag_and_drop_label_color = None
        self.drag_and_drop_text_changing = None

        self.__state_label_box_layout = FileDragAndDropperStateLabel(
            bg_color=(255, 255, 0, .5)
        )

        self.__state_info_box_layout = FileDragAndDropperStateInfo(
            bg_color=(255, 255, 0, .5)
        )

        self.set_state("label")

    def change_drag_and_drop_label_text(self, period, new_text):
        if self.drag_and_drop_text_changing:
            return
        self.drag_and_drop_text_changing = True

        if self.current_state == "label":
            label = self.__state_label_box_layout.ids.drag_and_drop_label
        else:
            label = self.__state_info_box_layout.ids.drag_and_drop_label

        old_text, label.text = label.text, new_text

        def set_default(text):
            label.text = old_text
            self.drag_and_drop_text_changing = False

        if period is not None:
            Clock.schedule_once(set_default, period)

    def invalid_file_animation(self, *args):
        if self.current_state == "label":
            canvas_color = self.__state_label_box_layout.canvas.before.get_group("bg_color")[0]
            canvas = self.__state_label_box_layout.canvas.before.get_group("background")[0]
            if self.__state_label_box_layout.default_drag_and_drop_label_color is None:
                self.__state_label_box_layout.default_drag_and_drop_label_color = tuple(canvas_color.rgba)
            default_drag_and_drop_label_color = self.__state_label_box_layout.default_drag_and_drop_label_color
        else:
            canvas_color = self.__state_info_box_layout.ids.drag_and_drop_label.canvas.before.get_group("bg_color")[0]
            canvas = self.__state_info_box_layout.ids.drag_and_drop_label.canvas.before.get_group("background")[0]
            if self.__state_info_box_layout.default_drag_and_drop_label_color is None:
                self.__state_info_box_layout.default_drag_and_drop_label_color = tuple(canvas_color.rgba)
            default_drag_and_drop_label_color = self.__state_info_box_layout.default_drag_and_drop_label_color

        pos_animation = Animation(
            pos=(canvas.pos[0] - 2, canvas.pos[1]),
            d=.05
        ) + Animation(
            pos=(canvas.pos[0] + 4, canvas.pos[1]),
            d=.05
        ) + Animation(
            pos=(canvas.pos[0] - 4, canvas.pos[1]),
            d=0.05
        ) + Animation(
            pos=(canvas.pos[0] + 4, canvas.pos[1]),
            d=0.05
        ) + Animation(
            pos=(canvas.pos[0], canvas.pos[1]),
            d=0.05
        )

        color_animation = Animation(
            rgba=(255, 0, 0, 0.6),
            d=0.0
        ) + Animation(
            rgba=(255, 0, 0, 0.6),
            d=0.25
        ) + Animation(
            rgba=default_drag_and_drop_label_color,
            d=0.75 if self.current_state == "label" else 0.25,
            t="in_expo"
        )

        Animation.cancel_all(canvas, "pos")

        color_animation.start(canvas_color)
        pos_animation.start(canvas)

        self.change_drag_and_drop_label_text(period=1, new_text="File is not a word file")

    def initialize_word_file_animation(self):
        self.set_state("info")

    def __initialize_file(self, file: pathlib.Path):
        if word.is_word_file(file) is False:
            self.invalid_file_animation()
            return

        self.current_working_file = file
        metadata = word.Metadata(self.current_working_file)

        if self.creator_text_input is not None:
            if (creator := metadata.creator) is not None:
                self.creator_text_input.text = creator
            else:
                self.creator_text_input.text = ""

        if self.last_modified_by_text_input is not None:
            if (last_modified_by := metadata.last_modified_by) is not None:
                self.last_modified_by_text_input.text = last_modified_by
            else:
                self.last_modified_by_text_input.text = ""

        if self.revision_text_input is not None:
            if (revision := metadata.revision) is not None:
                self.revision_text_input.text = str(revision)
            else:
                self.revision_text_input.text = ""

        if self.application_text_input is not None:
            if (application_name := metadata.application_name) is not None:
                self.application_text_input.text = application_name
            else:
                self.application_text_input.text = ""

        if self.editing_time_text_input is not None:
            if (editing_time := metadata.editing_time) is not None:
                self.editing_time_text_input.text = str(editing_time)
            else:
                self.editing_time_text_input.text = ""

        self.initialize_word_file_animation()

    def _on_file_drop(self, window, file_path, x, y):
        if window.children[0].current != "Antismirnova":
            return

        if self.current_state == "label":
            dropped_correctly = self.__state_label_box_layout.collide_point(*Window.mouse_pos)
        else:
            dropped_correctly = self.__state_info_box_layout.ids.drag_and_drop_label.collide_point(*Window.mouse_pos)

        if dropped_correctly:
            file_path = pathlib.Path(file_path.decode("utf-8"))
            self.__initialize_file(file_path)


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
        super(MainUi, self).__init__(**kwargs)


class AntismirnovaApp(App):
    def build(self):
        return ScreenManagement()


if __name__ == "__main__":
    if hasattr(sys, '_MEIPASS'):
        resource_add_path(os.path.join(sys._MEIPASS))
    Config.set('graphics', 'maxfps', 0)
    AntismirnovaApp().run()
