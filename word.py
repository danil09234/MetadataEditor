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

    @staticmethod
    def __register_all_namespaces(filename):
        namespaces = dict([node for _, node in ElementTree.iterparse(filename, events=['start-ns'])])
        for ns in namespaces:
            ElementTree.register_namespace(ns, namespaces[ns])

    @property
    def filepath(self):
        return self.__filepath

    @property
    def application_name(self) -> str | None:
        self.__extract_all()

        temp_file_path = pathlib.Path(f"temp_{self.__filepath.name}", "docProps", "app.xml")
        self.__register_all_namespaces(temp_file_path)
        for _, element in ElementTree.iterparse(temp_file_path):
            pattern = r"(\{.+\})(.+)"
            element_tag = re.search(pattern, element.tag).group(2)
            if element_tag == "Application":
                return element.text
        self.__remove_temp_folder()

    @application_name.setter
    def application_name(self, value: str):
        self.__extract_all()

        temp_file_path = pathlib.Path(f"temp_{self.__filepath.name}", "docProps", "app.xml")

        self.__register_all_namespaces(temp_file_path)
        tree = ElementTree.parse(temp_file_path)
        root = tree.getroot()

        for element in root:
            pattern = r"(\{.+\})(.+)"
            element_tag = re.search(pattern, element.tag).group(2)
            if element_tag == "Application":
                element.text = value
        tree.write(temp_file_path)

        self.__pack_all()
        self.__remove_temp_folder()

    def __init__(self, filepath: pathlib.Path):
        self.__filepath = filepath

