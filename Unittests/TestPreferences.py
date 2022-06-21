import unittest
from preferences import Preferences
from pathlib import Path


if __name__ == "__main__":
    BETA_YAML_FILE_NAME = Path("Unittests/Beta preferences.yaml")
else:
    BETA_YAML_FILE_NAME = Path("Beta preferences.yaml")


class TestPreferences(unittest.TestCase):
    def setUp(self) -> None:
        with open(BETA_YAML_FILE_NAME, "rb") as file:
            self.source_file = file.read()

        self.preferences = Preferences(BETA_YAML_FILE_NAME)

    def tearDown(self) -> None:
        with open(BETA_YAML_FILE_NAME, "wb") as file:
            file.write(self.source_file)

    # Tests for Preferences.new_command.application
    def test_new_command_application_preference_read(self):
        self.assertEqual("Microsoft Office Word", self.preferences.new_command.application,
                         "application preference is not as expected")

    # Tests for Preference.new_command.creator
    def test_new_command_creator_preference_read(self):
        self.assertEqual("admin", self.preferences.new_command.creator,
                         "creator preference is not as expected")

    # Tests for Preference.new_command.last_modified_by
    def test_new_command_last_modified_by_preference_read(self):
        self.assertEqual("admin", self.preferences.new_command.last_modified_by,
                         "last_modified_by preference is not as expected")

    # Tests for Preference.new_command.editing_time
    def test_new_command_editing_time_preference_read(self):
        self.assertEqual(0, self.preferences.new_command.editing_time,
                         "editing_time preference is not as expected")

    # Tests for Preference.new_command.revision
    def test_new_command_revision_preference_read(self):
        self.assertEqual(1, self.preferences.new_command.revision,
                         "revision preference is not as expected")

    # Tests for Preference.privet_smirnovoy.applications
    def test_privet_smirnovoy_applications_preference_read(self):
        expected_result = ["Minecraft", "Calculator", "Excel"]
        self.assertEqual(expected_result, self.preferences.privet_smirnovoy.applications,
                         "applications preference is not as expected")

    # Tests for Preference.privet_smirnovoy.creators
    def test_privet_smirnovoy_creators_preference_read(self):
        expected_result = ["User 1", "User 2", "User 3"]
        self.assertEqual(expected_result, self.preferences.privet_smirnovoy.creators,
                         "creators preference is not as expected")

    # Tests for Preference.privet_smirnovoy.modifiers
    def test_privet_smirnovoy_modifiers_preference_read(self):
        expected_result = ["User 1", "User 2"]
        self.assertEqual(expected_result, self.preferences.privet_smirnovoy.modifiers,
                         "modifiers preference is not as expected")

    # Tests for Preference.privet_smirnovoy.creators_number
    def test_privet_smirnovoy_creators_number_preference_read(self):
        self.assertEqual(2, self.preferences.privet_smirnovoy.creators_number,
                         "creators_number preference is not as expected")

    # Tests for Preference.privet_smirnovoy.modifiers_number
    def test_privet_smirnovoy_modifiers_number_preference_read(self):
        self.assertEqual(1, self.preferences.privet_smirnovoy.modifiers_number,
                         "modifiers_number preference is not as expected")
