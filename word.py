import zipfile
import pathlib
import xml.etree.ElementTree as ElementTree
import re
import os
import shutil


__word_file_suffixes = ["docx"]


def is_word_file(filepath: pathlib.Path) -> bool:
    if filepath.suffix in __word_file_suffixes:
        if zipfile.is_zipfile(filepath):
            return True
    return False


class PropertyNotFoundError(Exception):
    pass


class Metadata:
    def __remove_temp_folder(self):
        shutil.rmtree(pathlib.Path(f"temp_{self.__filepath.name}").absolute())

    def __extract_all(self):
        with zipfile.ZipFile(self.__filepath, "r") as zip_file:
            zip_file.extractall(f"temp_{self.__filepath.name}")

    def __pack_all(self):
        with zipfile.ZipFile(self.__filepath, 'w', compression=zipfile.ZIP_DEFLATED) as myzip:
            temp_folder = f"temp_{self.__filepath.name}"
            for root, dirs, files in os.walk(temp_folder):
                for f in files:
                    myzip.write(os.path.join(root, f), os.path.join(root.removeprefix(temp_folder), f))

    def __get_property_from_file(self, property_name: str, property_xml_file: pathlib.Path):
        match property_name, property_xml_file:
            case str(), pathlib.Path(suffix=".xml"):
                pass
            case _, _:
                raise TypeError(f"Metadata.__get_property_from_file(property_name, property_xml_file) should be"
                                f'str, pathlib.Path(suffix=".xml")'
                                f'(not {type(property_name)}, {type(property_xml_file)})')
        self.__extract_all()
        self.__register_all_namespaces(property_xml_file)

        tree = ElementTree.parse(property_xml_file)
        root = tree.getroot()

        for element in root:
            pattern = r"(\{.+\})(.+)"
            element_tag = re.search(pattern, element.tag).group(2)
            if element_tag == property_name:
                self.__remove_temp_folder()
                return element.text
        raise PropertyNotFoundError(f"property {property_name} not found")

    def __set_property_from_file(self, property_name: str, property_value: str, property_xml_file: pathlib.Path):
        match property_name, property_value, property_xml_file:
            case str(), str(), pathlib.Path(suffix=".xml"):
                pass
            case _, _, _:
                raise TypeError(f"Metadata.__set_property_from_file(property_name, property_value, property_xml_file)"
                                f'should be str, str, pathlib.Path(suffix=".xml")'
                                f"(not {type(property_name)}, {type(property_value)}, {type(property_xml_file)})")
        self.__extract_all()
        self.__register_all_namespaces(property_xml_file)

        tree = ElementTree.parse(property_xml_file)
        root = tree.getroot()

        for element in root:
            pattern = r"(\{.+\})(.+)"
            element_tag = re.search(pattern, element.tag).group(2)
            if element_tag == property_name:
                element.text = property_value
                tree.write(property_xml_file)
                self.__pack_all()
                self.__remove_temp_folder()
                return
        raise PropertyNotFoundError(f"property {property_name} not found")

    def __get_property_value(self, property_name: str) -> str:
        match property_name:
            case str():
                property_xml_file = pathlib.Path(f"temp_{self.__filepath.name}", "docProps", "app.xml")
                try:
                    return self.__get_property_from_file(property_name, property_xml_file)
                except PropertyNotFoundError:
                    pass

                property_xml_file = pathlib.Path(f"temp_{self.__filepath.name}", "docProps", "core.xml")
                return self.__get_property_from_file(property_name, property_xml_file)
            case _:
                raise TypeError(f"Metadata.__get_property_value(self, property_name) property_name should be str"
                                f"(not {type(property_name)})")

    def __set_property(self, property_name: str, property_value: str):
        match property_value, property_name:
            case str(), str():
                property_xml_file = pathlib.Path(f"temp_{self.__filepath.name}", "docProps", "app.xml")
                try:
                    self.__set_property_from_file(property_name, property_value, property_xml_file)
                    return
                except PropertyNotFoundError:
                    pass

                property_xml_file = pathlib.Path(f"temp_{self.__filepath.name}", "docProps", "core.xml")
                self.__set_property_from_file(property_name, property_value, property_xml_file)
            case str(), _:
                raise TypeError(f"Metadata.__set_property(property_name, property_value) "
                                f"property_value should be str (not {type(property_value)})")
            case _, str():
                raise TypeError(f"Metadata.__set_property(property_name, property_value) "
                                f"property_name should be str (not {type(property_name)})")
            case _, _:
                raise TypeError(f"Metadata.__set_property(property_name, property_value) "
                                f"property_name and property_value should be str, str "
                                f"(not {type(property_name)}, {type(property_value)})")

    @staticmethod
    def __register_all_namespaces(filename):
        namespaces = dict([node for _, node in ElementTree.iterparse(filename, events=['start-ns'])])
        for ns in namespaces:
            ElementTree.register_namespace(ns, namespaces[ns])

    @property
    def filepath(self) -> pathlib.Path:
        return self.__filepath

    @property
    def application_name(self) -> str:
        return self.__get_property_value("Application")

    @application_name.setter
    def application_name(self, value: str):
        match value:
            case str():
                self.__set_property("Application", value)
            case _:
                raise TypeError(f"Metadata.application_name should be str (not {type(value)})")

    @property
    def editing_time(self) -> int:
        return int(self.__get_property_value("TotalTime"))

    @editing_time.setter
    def editing_time(self, value: int):
        match value:
            case int():
                self.__set_property("TotalTime", str(value))
            case _:
                raise TypeError(f"Metadata.editing_time should be int (not {type(value)})")

    @property
    def creator(self) -> str:
        return self.__get_property_value("creator")

    @creator.setter
    def creator(self, value: str):
        match value:
            case str():
                self.__set_property("creator", value)
            case _:
                raise TypeError(f"Metadata.creator should be str (not {type(value)})")

    @property
    def last_modified_by(self) -> str:
        return self.__get_property_value("lastModifiedBy")

    @last_modified_by.setter
    def last_modified_by(self, value: str):
        match value:
            case str():
                self.__set_property("lastModifiedBy", value)
            case _:
                raise TypeError(f"Metadata.last_modified_by should be str (not {type(value)})")

    @property
    def revision(self):
        return self.__get_property_value("revision")

    @revision.setter
    def revision(self, value: int):
        match value:
            case int():
                self.__set_property("revision", str(value))
            case _:
                raise TypeError(f"Metadata.revision should be int (not {type(value)})")

    def __init__(self, filepath: pathlib.Path):
        self.__filepath = filepath
