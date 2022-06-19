import unittest
from pathlib import Path
from word import Metadata

if __name__ == "__main__":
    BETA_FILE_NAME_WITH_PROPERTIES = Path("Unittests/Beta word file with properties.docx")
    BETA_FILE_NAME_WITHOUT_PROPERTIES = Path("Unittests/Beta word file without properties.docx")
else:
    BETA_FILE_NAME_WITH_PROPERTIES = Path("Beta word file with properties.docx")
    BETA_FILE_NAME_WITHOUT_PROPERTIES = Path("Beta word file without properties.docx")


class TestMetadataByFileWithProperties(unittest.TestCase):
    def setUp(self) -> None:
        with open(BETA_FILE_NAME_WITH_PROPERTIES, "rb") as file:
            self.source_file = file.read()

        self.metadata = Metadata(BETA_FILE_NAME_WITH_PROPERTIES)

    def tearDown(self) -> None:
        self.assertFalse(self.metadata._temp_folder_path.exists(),
                         f'Temp folder "{self.metadata._temp_folder_path.name}" wasn\'t removed')
        with open(BETA_FILE_NAME_WITH_PROPERTIES, "wb") as file:
            file.write(self.source_file)

    # Tests for Metadata.filepath
    def test_metadata_filepath(self):
        self.assertEqual(BETA_FILE_NAME_WITH_PROPERTIES, self.metadata.filepath)

    # Tests for Metadata.application_name
    def test_metadata_application_name_read_write(self):
        testing_value = "Beta application name"
        self.metadata.application_name = testing_value
        self.assertEqual(testing_value, self.metadata.application_name, "Application names are different")

    def test_metadata_application_name_with_invalid_type_int(self):
        with self.assertRaises(TypeError):
            self.metadata.application_name = 123

    # Tests for Metadata.editing_time
    def test_metadata_editing_time_read_write(self):
        testing_value = 60
        self.metadata.editing_time = testing_value
        self.assertEqual(testing_value, self.metadata.editing_time, "File editing time is different")

    def test_metadata_editing_time_with_invalid_type_str(self):
        with self.assertRaises(TypeError):
            self.metadata.editing_time = "60"

    # Tests for Metadata.creator
    def test_metadata_creator_read_write(self):
        testing_value = "Beta creator"
        self.metadata.creator = testing_value
        self.assertEqual(testing_value, self.metadata.creator, "File creators are different")

    def test_metadata_creator_with_invalid_type_int(self):
        with self.assertRaises(TypeError):
            self.metadata.creator = 123

    # Tests for Metadata.last_modified_by
    def test_metadata_last_modified_by_read_write(self):
        testing_value = "Beta user"
        self.metadata.last_modified_by = testing_value
        self.assertEqual(testing_value, self.metadata.last_modified_by, "File modifiers are different")

    def test_metadata_last_modified_with_invalid_type_int(self):
        with self.assertRaises(TypeError):
            self.metadata.last_modified_by = 123

    # Tests for Metadata.revision
    def test_metadata_revision_read_write(self):
        testing_value = 100
        self.metadata.revision = testing_value
        self.assertEqual(testing_value, self.metadata.revision, "Revisions are different")

    def test_metadata_revision_with_invalid_type_str(self):
        with self.assertRaises(TypeError):
            self.metadata.revision = "100"


class TestMetadataByFileWithoutProperties(unittest.TestCase):
    def setUp(self) -> None:
        with open(BETA_FILE_NAME_WITHOUT_PROPERTIES, "rb") as file:
            self.source_file = file.read()

        self.metadata = Metadata(BETA_FILE_NAME_WITHOUT_PROPERTIES)

    def tearDown(self) -> None:
        self.assertFalse(self.metadata._temp_folder_path.exists(),
                         f'Temp folder "{self.metadata._temp_folder_path.name}" wasn\'t removed')
        with open(BETA_FILE_NAME_WITHOUT_PROPERTIES, "wb") as file:
            file.write(self.source_file)

    # Test for Metadata.filepath
    def test_metadata_filepath(self):
        self.assertEqual(BETA_FILE_NAME_WITHOUT_PROPERTIES, self.metadata.filepath)

    # Tests for Metadata.application_name
    def test_metadata_application_name_read(self):
        self.assertEqual(None, self.metadata.application_name)

    def test_metadata_application_name_read_write(self):
        testing_value = "Beta application name"
        self.metadata.application_name = testing_value
        self.assertEqual(testing_value, self.metadata.application_name)

    def test_metadata_application_name_with_invalid_type_int(self):
        with self.assertRaises(TypeError):
            self.metadata.application_name = 123

    # Tests for Metadata.editing_time
    def test_metadata_editing_time_read(self):
        self.assertEqual(0, self.metadata.editing_time)

    def test_metadata_editing_time_read_write(self):
        testing_value = 60
        self.metadata.editing_time = testing_value
        self.assertEqual(testing_value, self.metadata.editing_time)

    def test_metadata_editing_time_with_invalid_type_str(self):
        with self.assertRaises(TypeError):
            self.metadata.editing_time = "60"

    # Tests for Metadata.creator
    def test_metadata_creator_read(self):
        self.assertEqual(None, self.metadata.creator)

    def test_metadata_creator_read_write(self):
        testing_value = "Beta creator"
        self.metadata.creator = testing_value
        self.assertEqual(testing_value, self.metadata.creator)

    def test_metadata_creator_with_invalid_type_int(self):
        with self.assertRaises(TypeError):
            self.metadata.creator = 123

    # Tests for Metadata.last_modified
    def test_metadata_last_modified_read(self):
        self.assertEqual(None, self.metadata.last_modified_by)

    def test_metadata_last_modified_read_write(self):
        testing_value = "Beta user"
        self.metadata.last_modified_by = testing_value
        self.assertEqual(testing_value, self.metadata.last_modified_by)

    def test_metadata_last_modified_with_invalid_type(self):
        with self.assertRaises(TypeError):
            self.metadata.last_modified_by = 123

    # Tests for Metadata.revision
    def test_metadata_revision_read(self):
        self.assertEqual(None, self.metadata.revision)

    def test_metadata_revision_read_write(self):
        testing_value = 100
        self.metadata.revision = testing_value
        self.assertEqual(testing_value, self.metadata.revision)

    def test_metadata_revision_with_invalid_type_str(self):
        with self.assertRaises(TypeError):
            self.metadata.revision = "100"


if __name__ == '__main__':
    unittest.main()
