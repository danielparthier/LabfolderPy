"""
Module: entries

This module provides the `Entry` class and associated functions for interacting with
LabFolder entries via its API. The `Entry` class allows for creating, retrieving,
updating, and managing entries, including their metadata and associated elements.

Classes:
    Entry: Represents a LabFolder entry and provides methods to interact with it.

Functions:
    _get_entry(entry: Entry, entry_id: str, user_info: LabFolderUserInfo) -> list:
        Retrieves the details of a LabFolder entry, including its elements, based on the entry ID.

Dependencies:
    requests: Used for making HTTP requests to the LabFolder API.
    LabFolderUserInfo: Represents user authentication and API access information.
    parse_data_element: Parses individual data elements of an entry.
"""

import requests
from labfolder.classes.labfolder_access import LabFolderUserInfo
from labfolder.classes.data_elements import parse_data_element


class Entry:
    """Represents a LabFolder entry.

    Args:
        user_info (LabFolderUserInfo): User authentication and API access information.
        entry_id (str, optional): The unique identifier for the entry. Defaults to "".
        raw (bool, optional): If True, fetches raw data for the entry. Defaults to False.

    Attributes:
        entry_id (str): The unique identifier for the entry.
        title (str): The title of the entry.
        author_id (str): The ID of the author of the entry.
        project_id (str): The ID of the project associated with the entry.
        tags (list): A list of tags associated with the entry.
        creation_date (str): The creation date of the entry.
        elements (list): A list of elements associated with the entry.
    """

    def __init__(self, user_info: LabFolderUserInfo, entry_id="", raw=False):
        """Initializes an instance of the Entry class.

        Args:
            user_info (LabFolderUserInfo): User authentication and API access information.
            entry_id (str, optional): The unique identifier for the entry. Defaults to "".
            raw (bool, optional): If True, fetches raw data for the entry. Defaults to False.
        """
        self.entry_id = entry_id
        self.title = ""
        self.author_id = ""
        self.project_id = ""
        self.tags = []
        self.creation_date = ""
        self.elements = []
        if entry_id != "":
            self.get_entry(raw=raw, user_info=user_info)

    def get_entry(self, user_info: LabFolderUserInfo, raw=False):
        """Retrieves and processes the elements of an entry using the provided user information.

        Args:
            user_info (LabFolderUserInfo): User info for API access.
            raw (bool, optional): If True, returns raw elements. Defaults to False.

        Returns:
            None: Updates `self.elements` with retrieved and optionally processed elements.
        """
        self.elements = _get_entry(self, self.entry_id, user_info)
        if not raw:
            self.elements = [
                parse_data_element(element, user_info) for element in self.elements
            ]

    def add_title(self, title):
        """Assigns a title to the current instance.

        Args:
            title (str): The title to assign to the entry.
        """
        self.title = title

    def add_author_id(self, author_id):
        """Assigns an author ID to the current instance.

        Args:
            author_id (str or int): The unique identifier of the author to be assigned.
        """
        self.author_id = author_id

    def add_project_id(self, project_id):
        """Assigns a project ID to the current instance.

        Args:
            project_id (Any): The identifier to associate with the project.
        """
        self.project_id = project_id

    def add_tags(self, tags):
        """Add one or more tags to the entry.

        Args:
            tags (str or list): A single tag as a string, or a list of tags to add.

        Notes:
            Ensures that tags are unique after addition.
            If a list is provided, all elements are added. If a string is provided,
            it is added as a single tag.
        """
        if isinstance(tags, list):
            self.tags.extend(tags)
        if isinstance(tags, str):
            self.tags.append(tags)
        self.tags = list(set(self.tags))

    def add_creation_date(self, creation_date):
        """Adds a creation date to the entry.

        Args:
            creation_date (Any): The date to assign as the creation date for the entry.
        """
        self.creation_date = creation_date

    def add_elements(self, elements):
        """Add elements to the elements list.

        Args:
            elements (object): The element or elements to be added to the elements list.
                This can be a single object or a collection of objects.

        Notes:
            This method appends the provided `elements` as a single item to the
            `elements` list. If a collection (e.g., list) is passed, it will be added
            as a single entry, not extended.
        """
        self.elements.append(elements)

    def make_entry(self, user_info: LabFolderUserInfo):
        """Creates a new entry in the LabFolder system using the provided user information.

        Args:
            user_info (LabFolderUserInfo): An object containing the user's API address,
                authentication token, and user ID.

        Returns:
            None

        Side Effects:
            Sends a POST request to the LabFolder API to create a new entry with the
            current object's attributes.
            If successful, sets self.entry_id to the ID of the newly created entry.
            Prints an error message if the entry creation fails.

        Raises:
            requests.RequestException: If the HTTP request fails due to network issues or timeouts.
        """
        endpoint = f"{user_info.api_address}entries"
        entry = {
            "title": self.title,
            "author_id": user_info.user_id,
            "project_id": self.project_id,
            "tags": self.tags,
            "elements": self.elements,
            #     "locked": False
        }
        response = requests.post(
            endpoint, headers=user_info.auth_token, json=entry, timeout=10
        )
        if response.status_code == 201:
            response_json = response.json()
            self.entry_id = response_json["id"]
        else:
            print(f"Error creating entry: {response.status_code}")

    def update_entry(self, user_info: LabFolderUserInfo) -> None:
        """
        Updates an existing entry in the LabFolder system using the provided user
        information.

        Args:
            user_info (LabFolderUserInfo): An object containing the API address and
            authentication token.

        Returns:
            None

        Notes:
            If the entry_id is not set (empty string), the function prints a warning
            and returns without making a request. The entry is updated with the current
            values of title, author_id, project_id, tags, and elements. The 'locked'
            field is always set to False during the update.
        """
        if self.entry_id == "":
            print("Entry ID not set.")
            return None
        endpoint = f"{user_info.api_address}entries/{self.entry_id}"
        entry = {
            "title": self.title,
            "author_id": self.author_id,
            "project_id": self.project_id,
            "tags": self.tags,
            "elements": self.elements,
            "locked": False,
        }
        response = requests.put(
            endpoint, headers=user_info.auth_token, json=entry, timeout=10
        )
        if response.status_code != 200:
            print(f"Error updating entry: {response.status_code}")
        else:
            print("Entry updated successfully.")
        return None


def _get_entry(entry: Entry, entry_id: str, user_info: LabFolderUserInfo) -> list:
    """
    Fetches and populates an Entry object with metadata and associated elements
    from the LabFolder API.

    Args:
        entry (Entry): The Entry object to populate with data.
        entry_id (str): The unique identifier of the entry to retrieve.
        user_info (LabFolderUserInfo): User authentication and API address information.

    Returns:
        list: A list of dictionaries, each representing a detailed element
        associated with the entry.

    Notes:
        Supports element types: FILE, IMAGE, TEXT, TABLE, DATA_ELEMENT_GROUP,
        DATA, WELL_PLATE. Prints a message for unsupported element types.
        Populates the provided Entry object with metadata fields such as
        author_id, creation_date, tags, project_id, and title.
    """
    endpoint = f"{user_info.api_address}entries/{entry_id}"
    entry_dict = requests.get(endpoint, headers=user_info.auth_token, timeout=10).json()
    entry.author_id = entry_dict["author_id"]
    entry.creation_date = entry_dict["creation_date"]
    entry.tags = entry_dict["tags"]
    entry.project_id = entry_dict["project_id"]
    entry.title = entry_dict["title"]
    elements = []
    for element in entry_dict["elements"]:
        if element["type"] == "FILE":
            elements.append(
                requests.get(
                    user_info.api_address + f"elements/file/{element['id']}",
                    headers=user_info.auth_token,
                    timeout=10,
                ).json()
            )
        elif element["type"] == "IMAGE":
            elements.append(
                requests.get(
                    user_info.api_address + f"elements/image/{element['id']}",
                    headers=user_info.auth_token,
                    timeout=10,
                ).json()
            )
        elif element["type"] == "TEXT":
            elements.append(
                requests.get(
                    user_info.api_address + f"elements/text/{element['id']}",
                    headers=user_info.auth_token,
                    timeout=10,
                ).json()
            )
        elif element["type"] == "TABLE":
            elements.append(
                requests.get(
                    user_info.api_address + f"elements/table/{element['id']}",
                    headers=user_info.auth_token,
                    timeout=10,
                ).json()
            )
        elif element["type"] == "DATA_ELEMENT_GROUP":
            elements.append(
                requests.get(
                    user_info.api_address + f"elements/data/{element['id']}",
                    headers=user_info.auth_token,
                    timeout=10,
                ).json()
            )
        elif element["type"] == "DATA":
            elements.append(
                requests.get(
                    user_info.api_address + f"elements/data/{element['id']}",
                    headers=user_info.auth_token,
                    timeout=10,
                ).json()
            )
        elif element["type"] == "WELL_PLATE":
            elements.append(
                requests.get(
                    user_info.api_address + f"elements/well-plate/{element['id']}",
                    headers=user_info.auth_token,
                    timeout=10,
                ).json()
            )
        else:
            print(f"Element type {element['type']} not supported.")
    return elements
