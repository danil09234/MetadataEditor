import datetime
import re
import zipfile
import pathlib
import os
import shutil
from typing import NamedTuple, Literal
import xml.dom.minidom
import pytz


__word_file_suffixes = [".docx"]

# Constants
DEFAULT_WORD_CORE_XML_FILEPATH = pathlib.Path("default_word_core.xml")
DEFAULT_WORD_APP_XML_FILEPATH = pathlib.Path("default_word_app.xml")

W3CDTF_FORMAT = '%Y-%m-%dT%H:%M:%SZ'
RE_W3CDTF = "(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(.(\d{2}))?Z?"


def datetime_to_w3cdtf(dt):
    """Convert from a datetime to a timestamp string."""
    return datetime.datetime.strftime(dt, W3CDTF_FORMAT)


def w3cdtf_to_datetime(formatted_string):
    """Convert from a timestamp string to a datetime object."""
    match = re.match(RE_W3CDTF, formatted_string)
    digits = map(int, match.groups()[:6])
    return datetime.datetime(*digits)


class WordCoreProperty(NamedTuple):
    property_name: Literal[
        "title", "subject", "keywords", "description", "created", "modified", "lastModifiedBy", "revision", "creator"
    ]
    property_value: str | None = None


class WordAppProperty(NamedTuple):
    property_name: Literal[
        "TotalTime", "Application"
    ]
    property_value: str | None = None


class WordProperty(NamedTuple):
    property_name: Literal["creator", "lastModifiedBy", "revision", "TotalTime", "Application"]
    property_value: str | None = None


def is_word_file(filepath: pathlib.Path) -> bool:
    if filepath.suffix in __word_file_suffixes:
        if zipfile.is_zipfile(filepath):
            return True
    return False


class PropertyNotFoundError(Exception):
    pass


class WordRelsXml:
    """Class to work with .rels file"""
    def add_information_about_core(self):
        domtree = xml.dom.minidom.parse(str(self.xml_file_path.absolute()))
        core_file = domtree.documentElement

        new_property = domtree.createElement("Relationship")
        new_property.setAttribute("Id", "rId2")
        new_property.setAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
        new_property.setAttribute("Target", "docProps/core.xml")

        core_file.appendChild(new_property)

        with open(self.xml_file_path, "w", encoding='utf-8') as file:
            domtree.writexml(file, encoding='utf-8')

    def add_information_about_app(self):
        domtree = xml.dom.minidom.parse(str(self.xml_file_path.absolute()))
        core_file = domtree.documentElement

        new_property = domtree.createElement("Relationship")
        new_property.setAttribute("Id", "rId3")
        new_property.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
        new_property.setAttribute("Target", "docProps/app.xml")

        core_file.appendChild(new_property)

        with open(self.xml_file_path, "w", encoding='utf-8') as file:
            domtree.writexml(file, encoding='utf-8')

    def __init__(self, xml_file_path: pathlib.Path):
        match xml_file_path:
            case pathlib.Path():
                self.xml_file_path = xml_file_path
            case _:
                raise TypeError("WordRelsXml(xml_file_path) xml_file_path should be pathlib.Path"
                                f'(not {type(xml_file_path)})')


class WordContentTypesXml:
    """Class to work with [Content_Types].xml file"""
    def add_information_about_core(self):
        domtree = xml.dom.minidom.parse(str(self.xml_file_path.absolute()))
        core_file = domtree.documentElement

        new_property = domtree.createElement("Override")
        new_property.setAttribute("PartName", "/docProps/core.xml")
        new_property.setAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml")

        core_file.appendChild(new_property)

        with open(self.xml_file_path, "w", encoding='utf-8') as file:
            domtree.writexml(file, encoding='utf-8')

    def add_information_about_app(self):
        domtree = xml.dom.minidom.parse(str(self.xml_file_path.absolute()))
        core_file = domtree.documentElement

        new_property = domtree.createElement("Override")
        new_property.setAttribute("PartName", "/docProps/app.xml")
        new_property.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml")

        core_file.appendChild(new_property)

        with open(self.xml_file_path, "w", encoding='utf-8') as file:
            domtree.writexml(file, encoding='utf-8')

    def __init__(self, xml_file_path: pathlib.Path):
        match xml_file_path:
            case pathlib.Path():
                self.xml_file_path = xml_file_path
            case _:
                raise TypeError("WordContentTypesXml(xml_file_path) xml_file_path should be pathlib.Path"
                                f'(not {type(xml_file_path)})')


class WordCoreXml:
    """Class to work with core.xml file"""
    def __recover_namespaces(self) -> None:
        domtree = xml.dom.minidom.parse(str(self.xml_file_path.absolute()))
        core_file = domtree.documentElement
        namespaces = {
            "xmlns:cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "xmlns:dc": "http://purl.org/dc/elements/1.1/",
            "xmlns:dcterms": "http://purl.org/dc/terms/",
            "xmlns:dcmitype": "http://purl.org/dc/dcmitype/",
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"
        }

        for namespace, uri in namespaces.items():
            core_file.setAttribute(namespace, uri)

        with open(self.xml_file_path, "w", encoding='utf-8') as file:
            domtree.writexml(file, encoding='utf-8')

    @staticmethod
    def __get_tag_namespace(tag: str) -> str | None:
        core_namespaces = {
            "dc": ["title", "subject", "creator", "description"],
            "cp": ["keywords", "lastModifiedBy", "revision"],
            "dcterms": ["created", "modified"]
        }

        for namespace, namespace_tags in core_namespaces.items():
            for namespace_tag in namespace_tags:
                if namespace_tag == tag:
                    return namespace
        return None

    def __create_core_xml(self):
        domtree = xml.dom.minidom.parse(str(DEFAULT_WORD_CORE_XML_FILEPATH.absolute()))

        pathlib.Path(str(self.xml_file_path.absolute().parent)).mkdir()
        with open(self.xml_file_path.absolute(), mode="w") as file:
            domtree.writexml(file, encoding='utf-8')

        content_types = WordContentTypesXml(pathlib.Path(self.xml_file_path.parent.parent, "[Content_Types].xml"))
        content_types.add_information_about_core()

        rels = WordRelsXml(pathlib.Path(self.xml_file_path.parent.parent, "_rels", ".rels"))
        rels.add_information_about_core()
        self.__fill_created_date(datetime.datetime.now(pytz.utc))
        self.__fill_modified_data(datetime.datetime.now(pytz.utc))

    def __fill_created_date(self, date: datetime.datetime):
        self.__set_property(WordCoreProperty("created", datetime_to_w3cdtf(date)))

    def __fill_modified_data(self, date: datetime.datetime):
        self.__set_property(WordCoreProperty("modified", datetime_to_w3cdtf(date)))

    def __set_property(self, core_property: WordCoreProperty) -> None:
        if not self.xml_file_path.exists():
            self.__create_core_xml()
        self.__recover_namespaces()

        core_tags_subsequence = (
            "title", "subject", "creator", "keywords", "description",
            "lastModifiedBy", "revision", "created", "modified"
        )

        match core_property:
            case WordCoreProperty(
                    property_name=str("title" | "subject" | "creator" | "keywords" | "description" |
                                      "lastModifiedBy" | "revision" | "created" | "modified" as property_name),
                    property_value=str(property_value)):
                domtree = xml.dom.minidom.parse(str(self.xml_file_path.absolute()))
                core_file = domtree.documentElement

                core_property_name = f"{self.__get_tag_namespace(property_name)}:{property_name}"
                try:
                    if core_file.getElementsByTagName(core_property_name)[0].childNodes.length == 0:
                        text_node = domtree.createTextNode(property_value)
                        core_file.getElementsByTagName(core_property_name)[0].childNodes.append(text_node)
                    else:
                        core_file.getElementsByTagName(core_property_name)[0].childNodes[0].data = property_value
                except IndexError:
                    new_property = domtree.createElement(core_property_name)
                    new_property.appendChild(domtree.createTextNode(property_value))

                    after_property = None
                    for core_tag in core_tags_subsequence[core_tags_subsequence.index(property_name)+1:]:
                        after_core_property_tag = f"{self.__get_tag_namespace(core_tag)}:{core_tag}"
                        try:
                            after_property = core_file.getElementsByTagName(after_core_property_tag)[0]
                            break
                        except IndexError:
                            continue
                    core_file.insertBefore(new_property, after_property)

                with open(self.xml_file_path, "w", encoding='utf-8') as file:
                    domtree.writexml(file, encoding='utf-8')
            case _:
                raise TypeError("WordCoreXml.__set_property(core_property) core_property should be WordCoreProperty"
                                f"(not {type(core_property)})")

    def __get_property(self, property_name: str) -> str | None:
        match property_name:
            case str():
                if not self.xml_file_path.exists():
                    self.__create_core_xml()
                domtree = xml.dom.minidom.parse(str(self.xml_file_path.absolute()))
                core_file = domtree.documentElement

                core_property_name = f"{self.__get_tag_namespace(property_name)}:{property_name}"
                try:
                    return core_file.getElementsByTagName(core_property_name)[0].childNodes[0].data
                except IndexError:
                    return None
            case _:
                raise TypeError(f"WordCoreXml.__get_property(property_name) property_name should be str"
                                f"(not {type(property_name)})")

    @property
    def creator(self) -> str | None:
        return self.__get_property("creator")

    @creator.setter
    def creator(self, value: str | None) -> None:
        match value:
            case str():
                self.__set_property(WordCoreProperty("creator", value))
            case None:
                self.__set_property(WordCoreProperty("creator", ""))
            case _:
                raise TypeError(f"WordCoreXml.creator should be str (not {type(value)})")

    @property
    def last_modified_by(self) -> str | None:
        return self.__get_property("lastModifiedBy")

    @last_modified_by.setter
    def last_modified_by(self, value: str | None) -> None:
        match value:
            case str():
                self.__set_property(WordCoreProperty("lastModifiedBy", value))
            case None:
                self.__set_property(WordCoreProperty("lastModifiedBy", ""))
            case _:
                raise TypeError(f"WordCoreXml.last_modified_by should be str (not {type(value)})")

    @property
    def revision(self) -> int | None:
        revision = self.__get_property("revision")
        if revision is None:
            return None
        else:
            return int(revision)

    @revision.setter
    def revision(self, value: int | None) -> None:
        match value:
            case int():
                self.__set_property(WordCoreProperty("revision", str(value)))
            case None:
                self.__set_property(WordCoreProperty("revision", ""))
            case _:
                raise TypeError(f"WordCoreXml.revision should be int (not {type(value)})")

    def __init__(self, xml_file_path: pathlib.Path):
        self.xml_file_path = xml_file_path


class WordAppXml:
    """Class to work with app.xml file"""
    def __create_app_xml(self):
        domtree = xml.dom.minidom.parse(str(DEFAULT_WORD_APP_XML_FILEPATH.absolute()))

        if not self.xml_file_path.absolute().parent.exists():
            pathlib.Path(str(self.xml_file_path.absolute().parent)).mkdir()
        with open(self.xml_file_path.absolute(), mode="w") as file:
            domtree.writexml(file, encoding='utf-8')

        content_types = WordContentTypesXml(pathlib.Path(self.xml_file_path.parent.parent, "[Content_Types].xml"))
        content_types.add_information_about_app()

        rels = WordRelsXml(pathlib.Path(self.xml_file_path.parent.parent, "_rels", ".rels"))
        rels.add_information_about_app()

    def __set_property(self, app_property: WordAppProperty) -> None:
        app_tags_subsequence = (
            "Template", "TotalTime", "Pages", "Words", "Characters",
            "Application", "DocSecurity", "Lines", "Paragraphs",
            "ScaleCrop", "Company", "LinksUpToDate", "CharactersWithSpaces",
            "SharedDoc", "HyperlinksChanged", "AppVersion"
        )
        match app_property:
            case WordAppProperty("TotalTime" | "Application" as property_name, str() | None as property_value):
                if not self.xml_file_path.exists():
                    self.__create_app_xml()
                domtree = xml.dom.minidom.parse(str(self.xml_file_path.absolute()))
                core_file = domtree.documentElement

                try:
                    if core_file.getElementsByTagName(property_name)[0].childNodes.length == 0:
                        text_node = domtree.createTextNode(property_value)
                        core_file.getElementsByTagName(property_name)[0].childNodes.append(text_node)
                    else:
                        core_file.getElementsByTagName(property_name)[0].childNodes[0].data = property_value
                except IndexError:
                    new_property = domtree.createElement(property_name)
                    new_property.appendChild(domtree.createTextNode(property_value))

                    after_property = None
                    for core_tag in app_tags_subsequence[app_tags_subsequence.index(property_name) + 1:]:
                        try:
                            after_property = core_file.getElementsByTagName(core_tag)[0]
                            break
                        except IndexError:
                            continue
                    core_file.insertBefore(new_property, after_property)
                with open(self.xml_file_path, "w", encoding='utf-8') as file:
                    domtree.writexml(file, encoding='utf-8')
            case _:
                raise TypeError("WordXmlApp.__set_property(app_property) app_property should be WordAppProperty "
                                f'(not {type(app_property)})')

    def __get_property(self, app_property: str) -> str | None:
        match app_property:
            case str("TotalTime" | "Application" as property_name):
                if not self.xml_file_path.exists():
                    self.__create_app_xml()
                domtree = xml.dom.minidom.parse(str(self.xml_file_path.absolute()))
                core_file = domtree.documentElement

                try:
                    return core_file.getElementsByTagName(property_name)[0].childNodes[0].data
                except IndexError:
                    return None
            case _:
                raise TypeError('WordAppXml.__get_property(app_property) app_property should '
                                'be str("TotalTime" | "Application") '
                                f'(not {type(app_property)})')

    @property
    def application(self) -> str | None:
        return self.__get_property("Application")

    @application.setter
    def application(self, value: str | None) -> None:
        match value:
            case str():
                self.__set_property(WordAppProperty("Application", value))
            case None:
                self.__set_property(WordAppProperty("Application", ""))
            case _:
                raise TypeError("WordAppXml.application should be str"
                                f'(not {type(value)})')

    @property
    def total_time(self) -> int | None:
        total_time = self.__get_property("TotalTime")
        if total_time is None:
            return None
        else:
            return int(total_time)

    @total_time.setter
    def total_time(self, value: int | None) -> None:
        match value:
            case int():
                self.__set_property(WordAppProperty("TotalTime", str(value)))
            case None:
                self.__set_property(WordAppProperty("TotalTime", ""))
            case _:
                raise TypeError("WordAppXml.total_time should be int"
                                f'(not {type(value)})')

    def __init__(self, xml_file_path: pathlib.Path):
        match xml_file_path:
            case pathlib.Path():
                self.xml_file_path = xml_file_path
            case _:
                raise TypeError("WordAppXml(xml_file_path) xml_file_path should be pathlib.Path"
                                f'(not {type(xml_file_path)})')


class Metadata:
    @property
    def _temp_folder_path(self) -> pathlib.Path:
        return pathlib.Path(f"temp_{self.__filepath.name}")

    def __remove_temp_folder(self):
        shutil.rmtree(self._temp_folder_path.absolute())

    def __extract_all(self):
        with zipfile.ZipFile(self.__filepath, "r") as zip_file:
            zip_file.extractall(f"temp_{self.__filepath.name}")

    def __pack_all(self):
        with zipfile.ZipFile(self.__filepath, 'w', compression=zipfile.ZIP_DEFLATED) as myzip:
            temp_folder = self._temp_folder_path.name
            for root, dirs, files in os.walk(temp_folder):
                for f in files:
                    myzip.write(os.path.join(root, f), os.path.join(root.removeprefix(temp_folder), f))

    @property
    def filepath(self) -> pathlib.Path:
        return self.__filepath

    @property
    def application_name(self) -> str | None:
        try:
            self.__extract_all()
            property_xml_file = pathlib.Path(self._temp_folder_path, "docProps", "app.xml")
            app = WordAppXml(property_xml_file)
            application_name = app.application
            self.__remove_temp_folder()
            return application_name
        except PropertyNotFoundError:
            return None

    @application_name.setter
    def application_name(self, value: str | None) -> None:
        match value:
            case str() | None:
                self.__extract_all()
                property_xml_file = pathlib.Path(self._temp_folder_path, "docProps", "app.xml")
                app = WordAppXml(property_xml_file)
                app.application = value
                self.__pack_all()
                self.__remove_temp_folder()
            case _:
                raise TypeError(f"Metadata.application_name should be str or None (not {type(value)})")

    @property
    def editing_time(self) -> int:
        self.__extract_all()
        property_xml_file = pathlib.Path(self._temp_folder_path, "docProps", "app.xml")
        app = WordAppXml(property_xml_file)
        editing_time = app.total_time
        self.__remove_temp_folder()
        return editing_time

    @editing_time.setter
    def editing_time(self, value: int | None):
        match value:
            case int() | None:
                self.__extract_all()
                property_xml_file = pathlib.Path(self._temp_folder_path, "docProps", "app.xml")
                app = WordAppXml(property_xml_file)
                app.total_time = value
                self.__pack_all()
                self.__remove_temp_folder()
            case _:
                raise TypeError(f"Metadata.editing_time should be int or None (not {type(value)})")

    @property
    def creator(self) -> str | None:
        self.__extract_all()
        property_xml_file = pathlib.Path(self._temp_folder_path, "docProps", "core.xml")
        core = WordCoreXml(property_xml_file)
        creator = core.creator
        self.__remove_temp_folder()
        return creator

    @creator.setter
    def creator(self, value: str | None) -> None:
        match value:
            case str() | None:
                self.__extract_all()
                property_xml_file = pathlib.Path(self._temp_folder_path, "docProps", "core.xml")
                core = WordCoreXml(property_xml_file)
                core.creator = value
                self.__pack_all()
                self.__remove_temp_folder()
            case _:
                raise TypeError(f"Metadata.creator should be str or None (not {type(value)})")

    @property
    def last_modified_by(self) -> str | None:
        self.__extract_all()
        property_xml_file = pathlib.Path(self._temp_folder_path, "docProps", "core.xml")
        core = WordCoreXml(property_xml_file)
        last_modified_by = core.last_modified_by
        self.__remove_temp_folder()
        return last_modified_by

    @last_modified_by.setter
    def last_modified_by(self, value: str | None):
        match value:
            case str() | None:
                self.__extract_all()
                property_xml_file = pathlib.Path(self._temp_folder_path, "docProps", "core.xml")
                core = WordCoreXml(property_xml_file)
                core.last_modified_by = value
                self.__pack_all()
                self.__remove_temp_folder()
            case _:
                raise TypeError(f"Metadata.last_modified_by should be str or None (not {type(value)})")

    @property
    def revision(self) -> int | None:
        self.__extract_all()
        property_xml_file = pathlib.Path(self._temp_folder_path, "docProps", "core.xml")
        core = WordCoreXml(property_xml_file)
        revision = core.revision
        self.__remove_temp_folder()
        return revision

    @revision.setter
    def revision(self, value: int | None):
        match value:
            case int() | None:
                self.__extract_all()
                property_xml_file = pathlib.Path(self._temp_folder_path, "docProps", "core.xml")
                core = WordCoreXml(property_xml_file)
                core.revision = value
                self.__pack_all()
                self.__remove_temp_folder()
            case _:
                raise TypeError(f"Metadata.revision should be int or None (not {type(value)})")

    def __init__(self, filepath: pathlib.Path):
        self.__filepath = filepath
