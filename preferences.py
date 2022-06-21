import pathlib
import random
import yaml
from numpy import unique


class PreferenceNotFound(Exception):
    pass


class InvalidPreferenceValue(Exception):
    pass


class NewCommandPreferences:
    def __dump(self):
        with open(self.__preferences_filepath, "w") as yaml_file:
            yaml_file.write(yaml.safe_dump(self.__preferences))

    @property
    def __preferences(self) -> dict:
        with open(self.__preferences_filepath, "r") as yaml_file:
            preferences_dict = yaml.safe_load(yaml_file)
        return preferences_dict

    @property
    def application(self) -> str:
        try:
            return self.__preferences["new"]["application"]
        except KeyError:
            raise PreferenceNotFound("Preference application not found")

    @property
    def creator(self) -> str:
        try:
            return self.__preferences["new"]["creator"]
        except KeyError:
            raise PreferenceNotFound("Preference creator not found")

    @property
    def last_modified_by(self) -> str:
        try:
            return self.__preferences["new"]["last modified by"]
        except KeyError:
            raise PreferenceNotFound("Preference last modified by not found")

    @property
    def editing_time(self) -> int:
        try:
            return self.__preferences["new"]["editing time"]
        except KeyError:
            raise PreferenceNotFound("Preference editing time not found")

    @property
    def revision(self) -> int:
        try:
            return self.__preferences["new"]["revision"]
        except KeyError:
            raise PreferenceNotFound("Preference revision not found")

    def __init__(self, preferences_filepath: pathlib.Path):
        self.__preferences_filepath = preferences_filepath


class PrivetSmirnovoyPreference:
    @property
    def __preferences(self) -> dict:
        with open(self.__preferences_filepath, "r") as yaml_file:
            preferences_dict = yaml.safe_load(yaml_file)
        return preferences_dict

    @property
    def applications(self) -> list:
        try:
            return self.__preferences["privet_smirnovoy"]["applications"]
        except KeyError:
            raise PreferenceNotFound("Preference applications time not found")

    @property
    def random_application(self) -> str:
        return random.choice(self.applications)

    @property
    def creators(self) -> list:
        try:
            return self.__preferences["privet_smirnovoy"]["creators"]
        except KeyError:
            raise PreferenceNotFound("Preference creators not found")

    @property
    def random_creator(self) -> str:
        return random.choice(self.creators)

    @property
    def random_creators_list(self) -> list:
        unique_creators = list(unique(self.creators))
        if len(unique_creators) < self.creators_number:
            result = []
            for _ in range(self.creators_number):
                random_creator = random.choice(unique_creators)
                result.append(random_creator)
                unique_creators.remove(random_creator)
            return result
        elif len(unique_creators) == self.creators_number:
            return self.creators
        else:
            raise InvalidPreferenceValue

    @property
    def modifiers(self) -> list:
        try:
            return self.__preferences["privet_smirnovoy"]["modifiers"]
        except KeyError:
            raise PreferenceNotFound("Preference modifiers not found")

    @property
    def random_modifier(self) -> str:
        return random.choice(self.modifiers)

    @property
    def random_modifiers_list(self) -> list:
        unique_modifiers = list(unique(self.modifiers))
        if len(unique_modifiers) < self.modifiers_number:
            result = []
            for _ in range(self.modifiers_number):
                random_modifiers = random.choice(unique_modifiers)
                result.append(random_modifiers)
                unique_modifiers.remove(random_modifiers)
            return result
        elif len(unique_modifiers) == self.modifiers_number:
            return self.modifiers
        else:
            raise InvalidPreferenceValue

    @property
    def creators_number(self) -> int:
        try:
            return self.__preferences["privet_smirnovoy"]["creators number"]
        except KeyError:
            raise PreferenceNotFound("Preference creators number not found")

    @property
    def modifiers_number(self) -> int:
        try:
            return self.__preferences["privet_smirnovoy"]["modifiers number"]
        except KeyError:
            raise PreferenceNotFound("Preference modifiers number not found")
    
    def __init__(self, preferences_filepath: pathlib.Path):
        self.__preferences_filepath = preferences_filepath


class Preferences:
    @property
    def new_command(self):
        return self.__new_command_preferences
    
    @property
    def privet_smirnovoy(self):
        return self.__privet_smirnovoy_preferences
    
    def __init__(self, preferences_filepath: pathlib.Path):
        self.__new_command_preferences = NewCommandPreferences(preferences_filepath)
        self.__privet_smirnovoy_preferences = PrivetSmirnovoyPreference(preferences_filepath)
