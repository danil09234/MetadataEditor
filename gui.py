import os
import pathlib
import sys
from threading import Thread
from typing import Iterable

from kivy import utils, Config
from kivy.clock import Clock, mainthread
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.screenmanager import ScreenManager
from kivy.uix.screenmanager import Screen
from kivy.core.window import Window
from kivy.animation import Animation
from kivy.uix.textinput import TextInput
from kivy.uix.anchorlayout import AnchorLayout
from kivy.properties import VariableListProperty, StringProperty, ObjectProperty

from kivymd.app import MDApp

import preferences
import word


Window.minimum_width = preferences.GUI_MINIMUM_WIDTH
Window.minimum_height = preferences.GUI_MINIMUM_HEIGHT


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
        animation.bind(on_complete=lambda *_: self.new_widget_add())
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


class TextInputDefaultValue:
    @property
    def can_set_none(self):
        return self.__can_set_none

    def __can_set_none_setter(self, value):
        match value:
            case bool():
                self.__can_set_none = value
            case _:
                raise TypeError

    @property
    def input_name(self) -> str:
        return self.__input_name

    def __input_name_setter(self, value: str) -> None:
        match value:
            case str():
                self.__input_name = value
            case _:
                raise TypeError

    @property
    def input_object(self):
        return self.__input_object

    @input_object.setter
    def input_object(self, value):
        self.__input_object = value

    @property
    def input_value(self) -> str | None:
        return self.__input_value

    @input_value.setter
    def input_value(self, value: str | None):
        match value:
            case str() | None:
                self.__input_value = value
            case _:
                raise TypeError

    @property
    def changed(self) -> bool:
        if self.input_value is None and self.input_object.text == "":
            return False
        elif self.input_value != self.input_object.text:
            if self.can_set_none:
                return True
            elif self.input_object.text != "":
                return True
            else:
                return False
        else:
            return False

    @mainthread
    def change_text_input_text(self, text):
        self.input_object.text = str(text)

    def apply_changes(self):
        if self.input_object.text == "":
            self.input_value = None
            if not self.can_set_none:
                self.change_text_input_text("0")
                self.input_value = "0"
        else:
            self.input_value = self.input_object.text

    def __init__(self, input_name: str, input_object, input_value: str | None = None, can_set_none: bool = True):
        self.__can_set_none_setter(can_set_none)
        self.__input_name_setter(input_name)
        self.input_object = input_object
        self.input_value = input_value


class TextInputsDefaultValues:
    @property
    def changed(self) -> bool:
        for value in self.__values:
            if value.changed:
                return True

    def apply_all_changes(self):
        if self.changed:
            for value in self.__values:
                value.apply_changes()

    def __iter__(self):
        return iter(self.__values)

    def __init__(self, values: Iterable[TextInputDefaultValue]):
        self.__values = []
        for value in values:
            self.__values.append(value)


class FileDragAndDropper(BoxLayout):
    def update_info_widget(self):
        self.__state_info_box_layout.info_label_text = f'Word file "{self.current_working_file.name}"'
        self.__state_info_box_layout.image_source = "images/word_icon.png"
        self.__state_info_box_layout.drag_and_drop_label_bg_color = utils.get_color_from_hex("5ec6ff")

    @mainthread
    def set_state(self, state):
        match state, self.current_state:
            case str("label"), None:
                self.add_widget(self.__state_label_box_layout)
                self.current_state = "label"
            case str("label"), str("info"):
                change_widget_animation = WidgetChangeAnimation(
                    parent_widget=self,
                    old_widget=self.__state_info_box_layout,
                    new_widget=self.__state_label_box_layout,
                )
                change_widget_animation.start()

                self.default_text_input_values = None

                self.reset_button.disabled = True
                self.send_hello_button.disabled = True
                self.save_button.disabled = True

                self.creator_text_input.disabled = True
                self.last_modified_by_text_input.disabled = True
                self.revision_text_input.disabled = True
                self.application_text_input.disabled = True
                self.editing_time_text_input.disabled = True

                self.creator_text_input.text = ""
                self.last_modified_by_text_input.text = ""
                self.revision_text_input.text = ""
                self.application_text_input.text = ""
                self.editing_time_text_input.text = ""

                self.current_working_file = None
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
        self.default_text_input_values = None

        self.reset_button = None
        self.send_hello_button = None
        self.save_button = None

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

        def set_default():
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

        if self.current_working_file == file:
            return

        self.current_working_file = file
        metadata = word.Metadata(self.current_working_file)

        if (creator := metadata.creator) is None:
            creator = ""

        if (last_modified_by := metadata.last_modified_by) is None:
            last_modified_by = ""

        if (revision := metadata.revision) is None:
            revision = ""
        else:
            revision = str(revision)

        if (application := metadata.application_name) is None:
            application = ""

        if (editing_time := metadata.editing_time) is None:
            editing_time = "0"
        else:
            editing_time = str(editing_time)

        self.creator_text_input.text = creator
        self.last_modified_by_text_input.text = last_modified_by
        self.revision_text_input.text = revision
        self.application_text_input.text = application
        self.editing_time_text_input.text = editing_time

        self.default_text_input_values = TextInputsDefaultValues(
            [
                TextInputDefaultValue(
                    input_name="creator",
                    input_object=self.creator_text_input,
                    input_value=self.creator_text_input.text
                ),
                TextInputDefaultValue(
                    input_name="lastModifiedBy",
                    input_object=self.last_modified_by_text_input,
                    input_value=self.last_modified_by_text_input.text
                ),
                TextInputDefaultValue(
                    input_name="revision",
                    input_object=self.revision_text_input,
                    input_value=self.revision_text_input.text
                ),
                TextInputDefaultValue(
                    input_name="Application",
                    input_object=self.application_text_input,
                    input_value=self.application_text_input.text
                ),
                TextInputDefaultValue(
                    input_name="TotalTime",
                    input_object=self.editing_time_text_input,
                    input_value=self.editing_time_text_input.text,
                    can_set_none=False
                )
            ]
        )

        self.initialize_word_file_animation()

        try:
            preferences_valid = preferences.Preferences(preferences.PREFERENCES_FILEPATH).valid
            if preferences_valid:
                self.reset_button.disabled = False
                self.send_hello_button.disabled = False
            else:
                self.reset_button.disabled = True
                self.send_hello_button.disabled = True
        except FileNotFoundError:
            self.reset_button.disabled = True
            self.send_hello_button.disabled = True

        self.creator_text_input.disabled = False
        self.last_modified_by_text_input.disabled = False
        self.revision_text_input.disabled = False
        self.application_text_input.disabled = False
        self.editing_time_text_input.disabled = False

        self.save_button.disabled = True

    def _on_file_drop(self, window, file_path, x, y):
        for children in window.children:
            match children:
                case ScreenManagement():
                    if children.current != "Antismirnova":
                        return
                    else:
                        break
        else:
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
    input_filter = ObjectProperty(None)
    text_changed_function = ObjectProperty(lambda obj, text: None)

    @property
    def text(self):
        return self.__text_input.text

    @text.setter
    def text(self, value):
        self.__text_input.text = value

    def text_updated(self, *args):
        self.text_changed_function(*args)

    @staticmethod
    def change_input_filter(instance, value):
        instance.__text_input.input_filter = value

    def __init__(self, **kwargs):
        super(CustomTextInput, self).__init__(**kwargs)

        self.__text_input = TextInput(
            background_color=(0, 0, 0, 0),
            size_hint=(1, None),
            size=(self.size[0], 30),
            multiline=False,
            write_tab=False,
            input_filter=self.input_filter
        )
        self.__text_input.bind(text=self.text_updated)
        self.bind(input_filter=self.change_input_filter)

        self.add_widget(self.__text_input)


class ScreenManagement(ScreenManager):
    def __init__(self, **kwargs):
        super(ScreenManagement, self).__init__(**kwargs)


class MainUi(Screen):
    def update_reset_metadata_button(self, preferences_valid: bool | None = None):
        if preferences_valid is None:
            preferences_valid = preferences.NewCommandPreferences(preferences.PREFERENCES_FILEPATH).valid

        if preferences_valid:
            if self.default_values is not None:
                self.ids.reset_button.disabled = False
            else:
                self.ids.reset_button.disabled = True
        else:
            self.ids.reset_button.disabled = True

    def update_send_hello_button(self, preferences_valid: bool | None = None):
        if preferences_valid is None:
            preferences_valid = preferences.PrivetSmirnovoyPreference(preferences.PREFERENCES_FILEPATH).valid

        if preferences_valid:
            if self.default_values is not None:
                self.ids.send_hello_button.disabled = False
            else:
                self.ids.send_hello_button.disabled = True
        else:
            self.ids.send_hello_button.disabled = True

    def update_save_button(self):
        if self.default_values is None:
            self.ids.save_button.disabled = True
            return

        if self.default_values.changed:
            self.ids.save_button.disabled = False
        else:
            self.ids.save_button.disabled = True

    def check_preferences(self) -> None:
        new_command = preferences.NewCommandPreferences(preferences.PREFERENCES_FILEPATH)
        privet_smirnovoy_command = preferences.PrivetSmirnovoyPreference(preferences.PREFERENCES_FILEPATH)
        try:
            if not (preferences_valid := new_command.valid):
                self.update_reset_metadata_button(preferences_valid)
                self.show_reset_button_warning(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.')
            else:
                self.update_reset_metadata_button(preferences_valid)
                self.hide_reset_button_warning()
        except FileNotFoundError:
            self.update_reset_metadata_button(False)
            self.show_reset_button_warning(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.')

        try:
            if not (preferences_valid := privet_smirnovoy_command.valid):
                self.update_send_hello_button(preferences_valid)
                self.show_send_hello_button_warning(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.')
            else:
                self.update_send_hello_button(preferences_valid)
                self.hide_send_hello_button_warning()
        except FileNotFoundError:
            self.update_send_hello_button(False)
            self.show_send_hello_button_warning(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.')

    @mainthread
    def show_reset_button_warning(self, text: str | None = None):
        if text is not None:
            self.ids.reset_button_warning_icon.tooltip_text = text
        animation = Animation(
            opacity=1,
            duration=0.3
        )
        animation.start(self.ids.reset_button_warning_icon)

    @mainthread
    def hide_reset_button_warning(self):
        animation = Animation(
            opacity=0,
            duration=0.3
        )
        animation.start(self.ids.reset_button_warning_icon)

    @mainthread
    def show_send_hello_button_warning(self, text: str | None = None):
        if text is not None:
            self.ids.send_hello_button_warning_icon.tooltip_text = text
        animation = Animation(
            opacity=1,
            duration=0.3
        )
        animation.start(self.ids.send_hello_button_warning_icon)

    @mainthread
    def hide_send_hello_button_warning(self):
        animation = Animation(
            opacity=0,
            duration=0.3
        )
        animation.start(self.ids.send_hello_button_warning_icon)

    @mainthread
    def update_text_inputs(self,
                           editing_time_text_input: str | None = None,
                           revision_text_input: str | None = None,
                           creator_text_input: str | None = None,
                           last_modified_by_text_input: str | None = None,
                           application_text_input: str | None = None):
        if editing_time_text_input is not None:
            self.ids.editing_time_text_input.text = editing_time_text_input
        if revision_text_input is not None:
            self.ids.revision_text_input.text = revision_text_input
        if creator_text_input is not None:
            self.ids.creator_text_input.text = creator_text_input
        if last_modified_by_text_input is not None:
            self.ids.last_modified_by_text_input.text = last_modified_by_text_input
        if application_text_input is not None:
            self.ids.application_text_input.text = application_text_input

    @property
    def default_values(self) -> TextInputsDefaultValues:
        return self.ids.file_drag_and_dropper.default_text_input_values

    def check_values_for_difference(self):
        if self.default_values is None:
            self.ids.save_button.disabled = True
            return

        if self.default_values.changed:
            self.ids.save_button.disabled = False
        else:
            self.ids.save_button.disabled = True

    def text_input_text_updated(self) -> None:
        self.update_save_button()

    def save_button_pressed(self):
        self.ids.save_button.disabled = True
        save_process = Thread(target=self.save_changes)
        save_process.start()

    def save_changes(self):
        if (current_file := self.ids.file_drag_and_dropper.current_working_file) is None:
            return

        metadata = word.Metadata(current_file)

        if not self.default_values.changed:
            return

        for value in self.default_values:
            value.apply_changes()
            if not current_file.exists():
                self.ids.file_drag_and_dropper.set_state("label")
                break
            match value.input_name:
                case str("revision" | "TotalTime"):
                    match value.input_value:
                        case str():
                            metadata[value.input_name] = int(value.input_value)
                        case None:
                            metadata[value.input_name] = value.input_value
                case _:
                    metadata[value.input_name] = value.input_value

        self.update_save_button()

    def reset_data_button_pressed(self):
        reset_process = Thread(target=self.reset_data)
        reset_process.start()

    def reset_data(self):
        if self.ids.file_drag_and_dropper.current_working_file is None:
            return

        new_command_preferences = preferences.NewCommandPreferences(preferences.PREFERENCES_FILEPATH)

        completed_with_errors = False

        editing_time = preferences.DEFAULT_WORD_EDITING_TIME
        try:
            editing_time = new_command_preferences.editing_time
        except preferences.PreferenceNotFoundError:
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            self.show_reset_button_warning(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.')
            return
        except FileNotFoundError:
            self.show_reset_button_warning(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.')
            return

        revision = preferences.DEFAULT_WORD_REVISION
        try:
            revision = new_command_preferences.revision
        except preferences.PreferenceNotFoundError:
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            self.show_reset_button_warning(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.')
            return
        except FileNotFoundError:
            self.show_reset_button_warning(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.')
            return

        creator = preferences.DEFAULT_WORD_CREATOR
        try:
            creator = new_command_preferences.creator
        except preferences.PreferenceNotFoundError:
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            self.show_reset_button_warning(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.')
            return
        except FileNotFoundError:
            self.show_reset_button_warning(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.')
            return

        last_modified_by = preferences.DEFAULT_WORD_LAST_MODIFIED_BY
        try:
            last_modified_by = new_command_preferences.last_modified_by
        except preferences.PreferenceNotFoundError:
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            self.show_reset_button_warning(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.')
            return
        except FileNotFoundError:
            self.show_reset_button_warning(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.')
            return

        application_name = preferences.DEFAULT_WORD_APPLICATION_NAME
        try:
            application_name = new_command_preferences.application
        except preferences.PreferenceNotFoundError:
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            self.show_reset_button_warning(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.')
            return
        except FileNotFoundError:
            self.show_reset_button_warning(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.')
            return

        self.update_text_inputs(
            editing_time_text_input=str(editing_time),
            revision_text_input=str(revision),
            creator_text_input=creator,
            last_modified_by_text_input=last_modified_by,
            application_text_input=application_name
        )

        if completed_with_errors:
            self.show_reset_button_warning(f"Some preferences was not found\n"
                                           f"Please, check {preferences.PREFERENCES_FILEPATH.name}")
        else:
            self.hide_reset_button_warning()

    def send_hello_button_pressed(self):
        send_hello_process = Thread(target=self.send_hello)
        send_hello_process.start()

    def send_hello(self):
        privet_smirnovoy_preferences = preferences.PrivetSmirnovoyPreference(preferences.PREFERENCES_FILEPATH)

        completed_with_errors = False

        random_creators_string = None
        try:
            random_creators_string = privet_smirnovoy_preferences.random_creators_string
        except preferences.PreferenceNotFoundError:
            completed_with_errors = True
        except preferences.InvalidPreferenceValueError:
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            self.show_send_hello_button_warning(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.')
            return
        except FileNotFoundError:
            self.show_send_hello_button_warning(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.')
            return

        random_modifiers_string = None
        try:
            random_modifiers_string = privet_smirnovoy_preferences.random_modifiers_string
        except preferences.PreferenceNotFoundError:
            completed_with_errors = True
        except preferences.InvalidPreferenceValueError:
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            self.show_send_hello_button_warning(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.')
            return
        except FileNotFoundError:
            self.show_send_hello_button_warning(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.')
            return

        random_application = None
        try:
            random_application = privet_smirnovoy_preferences.random_application
        except preferences.PreferenceNotFoundError:
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            self.show_send_hello_button_warning(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.')
            return
        except FileNotFoundError:
            self.show_send_hello_button_warning(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.')
            return

        self.update_text_inputs(
            editing_time_text_input=str(preferences.PRIVET_SMIRNOVOY_EDITING_TIME),
            revision_text_input=str(preferences.PRIVET_SMIRNOVOY_REVISION),
        )

        if random_creators_string is not None:
            self.update_text_inputs(creator_text_input=random_creators_string)
        if random_modifiers_string is not None:
            self.update_text_inputs(last_modified_by_text_input=random_modifiers_string)
        if random_application is not None:
            self.update_text_inputs(application_text_input=random_application)

        if completed_with_errors:
            self.show_send_hello_button_warning(f"Some preferences was not found or invalid\n"
                                                f"Please, check {preferences.PREFERENCES_FILEPATH.name}")
        else:
            self.hide_send_hello_button_warning()

        self.update_save_button()

    def on_leave(self, *args):
        self.check_preferences_event.stop_clock()

    def on_enter(self, *args):
        self.check_preferences_event = Clock.schedule_interval(lambda *_: self.check_preferences(), 3)
        self.check_preferences_event()

    def __init__(self, **kwargs):
        self.check_preferences_event = None
        super(MainUi, self).__init__(**kwargs)


class AntismirnovaApp(MDApp):
    def build(self):
        self.icon = "images/app_icon.png"
        return ScreenManagement()


if __name__ == "__main__":
    from kivy.resources import resource_add_path, resource_find

    if hasattr(sys, '_MEIPASS'):
        resource_add_path(os.path.join(sys._MEIPASS))

    Config.set('graphics', 'maxfps', 0)
    AntismirnovaApp().run()
