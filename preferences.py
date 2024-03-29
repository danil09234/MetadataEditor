import pathlib
import random
import yaml
from numpy import unique


# Constants
# GUI window preferences
GUI_MINIMUM_WIDTH = 830
GUI_MINIMUM_HEIGHT = 600
# Preferences file
PREFERENCES_FILEPATH = pathlib.Path("preferences.yaml")
# Command "new"
DEFAULT_WORD_EDITING_TIME = 0
DEFAULT_WORD_REVISION = 1
DEFAULT_WORD_CREATOR = "admin"
DEFAULT_WORD_LAST_MODIFIED_BY = "admin"
DEFAULT_WORD_APPLICATION_NAME = "Microsoft Office Word"
# Command "privet_smirnovoy"
PRIVET_SMIRNOVOY_EDITING_TIME = 599940
PRIVET_SMIRNOVOY_REVISION = 9999999


class InvalidPreferencesStructureError(Exception):
    pass


class NewCommandPreferences:
    def __dump(self):
        with open(self.__preferences_filepath, "w", encoding="utf-8") as yaml_file:
            yaml_file.write(yaml.safe_dump(self.__preferences))

    @property
    def __preferences(self) -> dict:
        with open(self.__preferences_filepath, "r", encoding="utf-8") as yaml_file:
            try:
                preferences_dict = yaml.safe_load(yaml_file)
            except yaml.YAMLError:
                raise InvalidPreferencesStructureError
        return preferences_dict

    @property
    def valid(self) -> bool:
        _ = self.application
        _ = self.creator
        _ = self.last_modified_by
        _ = self.editing_time
        _ = self.revision

        return True

    @property
    def application(self) -> str:
        try:
            return self.__preferences["new"]["application"]
        except KeyError:
            raise AttributeError("Preference application not found")
        except TypeError:
            raise AttributeError('Preferences section "new" is invalid')

    @property
    def creator(self) -> str:
        try:
            return self.__preferences["new"]["creator"]
        except KeyError:
            raise AttributeError("Preference creator not found")
        except TypeError:
            raise AttributeError('Preferences section "new" is invalid')

    @property
    def last_modified_by(self) -> str:
        try:
            return self.__preferences["new"]["last modified by"]
        except KeyError:
            raise AttributeError("Preference last modified by not found")
        except TypeError:
            raise AttributeError('Preferences section "new" is invalid')

    @property
    def editing_time(self) -> int:
        try:
            if (value := str(self.__preferences["new"]["editing time"])).isdigit():
                if len(str(int(value))) < 10:
                    return int(value)
                else:
                    raise ValueError("Editing time value should consist of less than 10 digits")
            else:
                raise ValueError("Editing time value must be integer")
        except KeyError:
            raise AttributeError("Preference editing time not found")
        except TypeError:
            raise AttributeError('Preferences section "new" is invalid')

    @property
    def revision(self) -> int:
        try:
            if (value := str(self.__preferences["new"]["revision"])).isdigit():
                return int(value)
            else:
                raise ValueError('Revision value must be integer')
        except KeyError:
            raise AttributeError("Preference revision not found")
        except TypeError:
            raise AttributeError('Preferences section "new" is invalid')

    def __init__(self, preferences_filepath: pathlib.Path = PREFERENCES_FILEPATH):
        self.__preferences_filepath = preferences_filepath


class PrivetSmirnovoyPreference:
    @property
    def __preferences(self) -> dict:
        with open(self.__preferences_filepath, "r", encoding="utf-8") as yaml_file:
            try:
                preferences_dict = yaml.safe_load(yaml_file)
            except yaml.YAMLError:
                raise InvalidPreferencesStructureError
        return preferences_dict

    @property
    def valid(self) -> bool:
        _ = self.applications
        _ = self.random_application
        _ = self.creators
        _ = self.random_creator
        _ = self.random_creators_list
        _ = self.random_creators_string
        _ = self.modifiers
        _ = self.random_modifier
        _ = self.random_modifiers_list
        _ = self.random_modifiers_string
        _ = self.creators_number
        _ = self.modifiers_number

        return True

    @property
    def applications(self) -> list:
        try:
            return self.__preferences["privet_smirnovoy"]["applications"]
        except KeyError:
            raise AttributeError("Preference applications not found")
        except TypeError:
            raise AttributeError('Preferences section "privet_smirnovoy" is invalid')

    @property
    def random_application(self) -> str:
        return random.choice(self.applications)

    @property
    def creators(self) -> list:
        try:
            return self.__preferences["privet_smirnovoy"]["creators"]
        except KeyError:
            raise AttributeError("Preference creators not found")
        except TypeError:
            raise AttributeError('Preferences section "privet_smirnovoy" is invalid')

    @property
    def random_creator(self) -> str:
        return random.choice(self.creators)

    @property
    def random_creators_list(self) -> list:
        unique_creators = list(unique(self.creators))
        if len(unique_creators) > self.creators_number:
            result = []
            for _ in range(self.creators_number):
                random_creator = random.choice(unique_creators)
                result.append(random_creator)
                unique_creators.remove(random_creator)
            return result
        elif len(unique_creators) == self.creators_number:
            return self.creators
        else:
            raise ValueError("Invalid preferences value")

    @property
    def random_creators_string(self):
        creators_string = ""
        random_creators_list = self.random_creators_list
        for index, creator in enumerate(random_creators_list):
            creators_string += creator
            if index != len(random_creators_list) - 1:
                creators_string += "; "
        return creators_string

    @property
    def modifiers(self) -> list:
        try:
            return self.__preferences["privet_smirnovoy"]["modifiers"]
        except KeyError:
            raise AttributeError("Preference modifiers not found")
        except TypeError:
            raise AttributeError('Preferences section "privet_smirnovoy" is invalid')

    @property
    def random_modifier(self) -> str:
        return random.choice(self.modifiers)

    @property
    def random_modifiers_list(self) -> list:
        unique_modifiers = list(unique(self.modifiers))
        if len(unique_modifiers) > self.modifiers_number:
            result = []
            for _ in range(self.modifiers_number):
                random_modifiers = random.choice(unique_modifiers)
                result.append(random_modifiers)
                unique_modifiers.remove(random_modifiers)
            return result
        elif len(unique_modifiers) == self.modifiers_number:
            return self.modifiers
        else:
            raise ValueError("Invalid preferences value")

    @property
    def random_modifiers_string(self) -> str:
        modifiers_string = ""
        random_modifiers_list = self.random_modifiers_list
        for index, modifier in enumerate(random_modifiers_list):
            modifiers_string += modifier
            if index != len(random_modifiers_list) - 1:
                modifiers_string += "; "
        return modifiers_string

    @property
    def creators_number(self) -> int:
        try:
            return self.__preferences["privet_smirnovoy"]["creators number"]
        except KeyError:
            raise AttributeError("Preference creators number not found")
        except TypeError:
            raise AttributeError('Preferences section "privet_smirnovoy" is invalid')

    @property
    def modifiers_number(self) -> int:
        try:
            return self.__preferences["privet_smirnovoy"]["modifiers number"]
        except KeyError:
            raise AttributeError("Preference modifiers number not found")
        except TypeError:
            raise AttributeError('Preferences section "privet_smirnovoy" is invalid')
    
    def __init__(self, preferences_filepath: pathlib.Path = PREFERENCES_FILEPATH):
        self.__preferences_filepath = preferences_filepath


class Preferences:
    @property
    def valid(self) -> bool:
        return self.new_command.valid and self.privet_smirnovoy.valid

    @property
    def new_command(self):
        return self.__new_command_preferences
    
    @property
    def privet_smirnovoy(self):
        return self.__privet_smirnovoy_preferences
    
    def __init__(self, preferences_filepath: pathlib.Path = PREFERENCES_FILEPATH):
        self.__new_command_preferences = NewCommandPreferences(preferences_filepath)
        self.__privet_smirnovoy_preferences = PrivetSmirnovoyPreference(preferences_filepath)
