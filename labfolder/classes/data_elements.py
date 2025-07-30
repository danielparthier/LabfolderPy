"""
A module for managing data elements in LabFolder, including loading, writing,
and updating data elements via the LabFolder API.
"""

import io
from copy import deepcopy

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import requests
from PIL import Image

from labfolder.classes.labfolder_access import LabFolderUserInfo


class DataElement:
    """A class representing a data element in LabFolder.

    Provides methods to load, write, and update data elements via the LabFolder API.

    Args:
        element_id (str, optional): The unique identifier for the data element.
            Defaults to an empty string.
        description (str, optional): A description of the data element.
            Defaults to an empty string.

    Attributes:
        type (str): The type of the element, set to 'DATA'.
        element_id (str): The unique identifier for the data element.
        description (str): The description of the data element.
    """

    def __init__(self, element_id="", description=""):
        """Initialize a new data element with the specified ID and description.

        Args:
            element_id (str, optional): The unique identifier for the data element.
                Defaults to an empty string.
            description (str, optional): A textual description of the data element.
                Defaults to an empty string.
        """
        self.type = "DATA"
        self.element_id = element_id
        self.description = description

    def load_data(self, user_info: LabFolderUserInfo, element_id=""):
        """Loads data for the specified element ID from the LabFolder API.

        Updates the object's attributes with retrieved data.

        Args:
            user_info (LabFolderUserInfo): An object containing user authentication
                information and API address.
            element_id (str, optional): The ID of the data element to load. If not
                provided, uses the object's `element_id` attribute.

        Returns:
            None or dict: Returns None if no element ID is provided. Otherwise,
                updates the object's `description` and `element_id` attributes with
                the retrieved data.
        """
        if element_id == "":
            element_id = self.element_id
        if element_id == "":
            print("No data ID provided.")
            return None
        data = requests.get(
            f"{user_info.api_address}elements/data/{element_id}",
            headers=user_info.auth_token,
            timeout=10,
        ).json()
        self.description = data["description"]
        self.element_id = data["id"]
        return None

    def write_to_labfolder(self, user_info: LabFolderUserInfo, entry_id=""):
        """Writes the current data element to Labfolder.

        Args:
            user_info (LabFolderUserInfo): An object containing the user's Labfolder
                API address and authentication token.
            entry_id (str, optional): The ID of the Labfolder entry to which the data
                element should be written. Default is an empty string.

        Returns:
            None

        Notes:
            If `entry_id` is not provided, the function prints a message and returns None.
            On successful write (HTTP 201), sets `self.element_id` to the returned element ID.
            Otherwise, prints an error message and the response text.
        """
        if entry_id == "":
            print("No entry_id provided.")
            return None
        response = requests.post(
            f"{user_info.api_address}elements/data",
            headers=user_info.auth_token,
            json={"entry_id": entry_id, "description": self.description},
            timeout=10,
        )
        status = _handle_response(self, response, return_status=True)
        if status:
            self.element_id = response.json()["id"]
        return None

    def update_on_labfolder(self, user_info: LabFolderUserInfo):
        """Updates the description of the data element on Labfolder via a PUT request.

        Args:
            user_info (LabFolderUserInfo): An object containing the user's Labfolder
                API address and authentication token.

        Returns:
            None: This method does not return anything.

        Notes:
            - If the description is empty, the method prints a message and returns without
              making a request.
            - Prints a message indicating whether the update was successful or not.
        """
        if self.description == "":
            print("No data to write.")
            return None
        response = requests.put(
            f"{user_info.api_address}elements/data/{self.element_id}",
            headers=user_info.auth_token,
            json={"id": self.element_id, "description": self.description},
            timeout=10,
        )
        _handle_response(self, response)
        return None

    def __repr__(self):
        """Return a string representation of the DataElement instance.

        Returns:
            str: A formatted string displaying the type and description of the DataElement.
        """
        return f"DataElement(type={self.type}, description={self.description})"


class DescriptiveDataElement:
    """A class representing a descriptive data element with a title, ID, and description.

    Args:
        title (str, optional): The title of the data element. Defaults to an empty string.
        element_id (str, optional): The unique identifier for the data element.
            Defaults to an empty string.
        description (str, optional): A textual description of the data element.
            Defaults to an empty string.

    Attributes:
        type (str): The type of the data element, set to 'DESCRIPTIVE_DATA_ELEMENT'.
        title (str): The title of the data element.
        element_id (str): The unique identifier for the data element.
        description (str): The description of the data element.
    """

    def __init__(self, title="", element_id="", description=""):
        self.type = "DESCRIPTIVE_DATA_ELEMENT"
        self.title = title
        self.element_id = element_id
        self.description = description

    def to_dict(self):
        """Convert the object's attributes to a dictionary.

        Returns:
            dict: A dictionary containing the type, title, and description of the object.
        """
        return {"type": self.type, "title": self.title, "description": self.description}

    def __repr__(self):
        """Return a string representation of the DataElement instance.

        Returns:
            str: A formatted string displaying the type, title, and description.
        """
        return f"DataElement(type={self.type}, title={self.title}, description={self.description})"


class FileElement:
    """Represents a file element with a specific identifier.

    Args:
        element_id (str): The unique identifier for the file element.

    Attributes:
        type (str): The type of the element, always set to 'FILE'.
        element_id (str): The unique identifier for the file element.
    """

    def __init__(self, element_id):
        """Initialize a new instance of the class with the specified element ID.

        Args:
            element_id (Any): The unique identifier for the data element.
        """
        self.type = "FILE"
        self.element_id = element_id

    def __repr__(self):
        """Return a string representation of the DataElement instance.

        Returns:
            str: A string describing the type and element_id of the DataElement.
        """
        return f"DataElement(type={self.type}, title={self.element_id})"


class ImageElement:
    """ImageElement represents an image data element for integration with LabFolder.

    Attributes:
        type (str): The type of the data element, set to 'IMAGE'.
        title (str): The title of the image element.
        element_id (str): The unique identifier for the image element.
        original_file_content_type (str): The MIME type of the original image file.
        creation_date (str): The creation date of the image element.
        image (PIL.Image or None): The loaded image object, or None if not loaded.
        owner_id (str): The ID of the owner of the image element.
    """

    def __init__(self, title="", element_id="", original_file_content_type="image/png"):
        """Initialize an IMAGE data element with optional metadata.

        Args:
            title (str, optional): The title of the image element. Default is empty string.
            element_id (str, optional): The unique identifier for the element.
                Default is empty string.
            original_file_content_type (str, optional): The MIME type of the original file.
                Default is 'image/png'.
        """
        self.type = "IMAGE"
        self.title = title
        self.element_id = element_id
        self.original_file_content_type = original_file_content_type
        self.creation_date = ""
        self.image = None
        self.owner_id = ""

    def load_image(self, user_info: LabFolderUserInfo, element_id=""):
        """Loads an image element from the LabFolder API and updates the object's attributes.

        Args:
            user_info (LabFolderUserInfo): An object containing user authentication
                information and API address.
            element_id (str, optional): The ID of the image element to load. If not provided,
                uses the object's `element_id` attribute.

        Returns:
            None

        Notes:
            - Updates the object's `title`, `owner_id`, `creation_date`, and `image` attributes.
            - Prints a message and returns None if no image ID is provided.
        """

        if element_id == "":
            element_id = self.element_id
        if element_id == "":
            print("No image ID provided.")
            return None
        image_info = requests.get(
            user_info.api_address + f"elements/image/{element_id}",
            headers=user_info.auth_token,
            timeout=10,
        ).json()
        image = requests.get(
            user_info.api_address + f"elements/image/{element_id}/original-data",
            headers=user_info.auth_token,
            timeout=10,
        )
        self.title = image_info["title"]
        self.owner_id = image_info["owner_id"]
        self.creation_date = image_info["creation_date"]
        self.image = Image.open(io.BytesIO(image.content))
        return None

    def show_image(self):
        """Displays the image stored in the object using matplotlib.

        Returns:
            None: This method does not return any value.

        Notes:
            If the image attribute is None, prints a message and returns None.
            Otherwise, displays the image without axes.
        """
        if self.image is None:
            print("No image to show.")
            return None
        plt.imshow(self.image)
        plt.axis("off")
        plt.show()
        return None

    # TODO: Implement write_to_labfolder - it probably has to be written as file (?)

    def __repr__(self):
        """Return a string representation of the DataElement instance.

        Returns:
            str: A string describing the DataElement, including its type, id, title,
                and original file content type.
        """
        return (
            f"DataElement(type={self.type}, id={self.element_id}, "
            f"title={self.title}, original_file_content_type={self.original_file_content_type})"
        )


class TextElement:
    """A class representing a text element for interaction with the LabFolder API.

    Args:
        content (str, optional): The textual content of the element. Defaults to empty string.
        element_id (str, optional): The unique identifier for the text element.
            Defaults to empty string.

    Attributes:
        type (str): The type of the element, set to 'TEXT'.
        element_id (str): The unique identifier for the text element.
        content (str): The textual content of the element.
    """

    def __init__(self, content="", element_id=""):
        """Initialize a new instance of the class.

        Args:
            content (str, optional): The textual content to be stored in the element.
                Defaults to an empty string.
            element_id (str, optional): The unique identifier for the element.
                Defaults to an empty string.
        """
        self.type = "TEXT"
        self.element_id = element_id
        self.content = content
        self.id = ""

    def load_text(self, user_info: LabFolderUserInfo, element_id=""):
        """Load the text content of a data element from the LabFolder API.

        Args:
            user_info (LabFolderUserInfo): An object containing user authentication
                information and API address.
            element_id (str, optional): The ID of the text element to load. If not provided,
                uses the instance's `element_id`.

        Returns:
            None or dict: Returns None if no element ID is provided or if the request fails.
            Otherwise, updates the instance's `content` and `element_id` attributes
            with the retrieved data.

        Notes:
        -----
        This method sends a GET request to the LabFolder API to retrieve the text
        content of the specified element.
        """
        if element_id == "":
            element_id = self.element_id
        if element_id == "":
            print("No text ID provided.")
            return None
        text = requests.get(
            f"{user_info.api_address}elements/text/{element_id}",
            headers=user_info.auth_token,
            timeout=10,
        ).json()
        self.content = text["content"]
        self.element_id = text["id"]
        return None

    def write_to_labfolder(self, user_info: LabFolderUserInfo, entry_id=""):
        """Writes the content of the object to Labfolder as a text element.

        Args:
            user_info (LabFolderUserInfo): An object containing the user's authentication
                token and API address for Labfolder.
            entry_id (str, optional): The ID of the Labfolder entry to which the text
                will be written. If not provided or empty, the method will not proceed.

        Returns:
            None

        Notes:
        -----
        If the `entry_id` is not provided, the method prints a message and returns `None`.
        On successful creation (HTTP 201), the method updates the object's `id`
        attribute with the ID returned by Labfolder.
        Otherwise, it prints an error message with the status code and response text.
        """
        if entry_id == "":
            print("No entry_id provided.")
            return None
        response = requests.post(
            f"{user_info.api_address}elements/text",
            headers=user_info.auth_token,
            json={"entry_id": entry_id, "content": self.content},
            timeout=10,
        )
        status = _handle_response(self, response, return_status=True)
        if status:
            self.id = response.json()["id"]
        return None

    def update_on_labfolder(self, user_info: LabFolderUserInfo):
        """Updates the text content of the element on Labfolder using the provided user information.

        Args:
            user_info (LabFolderUserInfo): An object containing the user's Labfolder
                API address and authentication token.

        Returns:
            None: This method does not return any value.

        Notes:
        -----
        If the content is empty, the method prints a message and returns without
        making a request.
        Prints the result of the update operation, including the status code and
        response text if the update fails.
        """
        if self.content == "":
            print("No text to write.")
            return None
        response = requests.put(
            f"{user_info.api_address}elements/text/{self.element_id}",
            headers=user_info.auth_token,
            json={"id": self.element_id, "content": self.content},
            timeout=10,
        )
        _handle_response(self, response)
        return None

    def __repr__(self):
        """Return a string representation of the DataElement instance.

        Returns:
            str: A string describing the DataElement, including its type, id, and content.
        """
        return f"DataElement(type={self.type}, id={self.element_id}, content={self.content})"


class DataElementGroup:
    """
    A class representing a group of data elements for Labfolder integration.

    Args:
        title (str, optional): The title of the data element group. Default is an
            empty string.
        children (list, optional): A list of child data elements. Default is None,
            which initializes an empty list.

    Attributes:
        type (str): The type identifier for the data element group.
        title (str): The title of the data element group.
        element_id (str): The unique identifier for the data element group in
            Labfolder.
        children (list): The list of child data elements.
    """

    def __init__(self, title="", children=None):
        """Initialize a DATA_ELEMENT_GROUP instance.

        Args:
            title (str, optional): The title of the data element group.
            Defaults to an empty string.
            children (list, optional): A list of child elements to be
            included in the group. If not provided or not a list,
            an empty list is used.
        """
        self.type = "DATA_ELEMENT_GROUP"
        self.title = title
        self.element_id = ""
        if isinstance(children, list):
            self.children = children
        else:
            self.children = []

    def add_child(self, child):
        """Add a child element to the current object's list of children.

        Args:
            child (object): The child element to be added to the list of children.

        Returns:
            None
        """
        self.children.append(child)

    def to_dict(self):
        """Convert the object and its children into a dictionary representation.

        Returns:
            dict: A dictionary containing the object's type, title, and a list of its
                children's dictionary representations.
        """
        return {
            "type": self.type,
            "title": self.title,
            "children": [child.to_dict() for child in self.children],
        }

    def write_to_labfolder(self, user_info: LabFolderUserInfo, entry_id=""):
        """
        Write the data element to Labfolder by creating a new data element group.

        Args:
            user_info (LabFolderUserInfo): An object containing user authentication
            information and API address for Labfolder.
            entry_id (str, optional): The ID of the Labfolder entry to which the data
            element should be written. If not provided or empty, the function will
            print a message and return None.

        Returns:
            None

        Notes:
            If the data element group is successfully written to Labfolder (HTTP status
            code 201), the `element_id` attribute of the object is updated with the
            returned ID. Otherwise, an error message and the response text are printed.
        """
        if entry_id == "":
            print("No entry_id provided.")
            return None
        content = {
            "entry_id": entry_id,
            "data_elements": [self.to_dict()],
            "locked": False,
        }
        response = requests.post(
            f"{user_info.api_address}elements/data",
            headers=user_info.auth_token,
            json=content,
            timeout=10,
        )
        _handle_response(self, response)
        return None

    def update_on_labfolder(self, user_info: LabFolderUserInfo):
        """
        Update the data element group on Labfolder using the provided user information.

        Args:
            user_info (LabFolderUserInfo): An object containing the API address and
            authentication token required to access Labfolder.

        Returns:
            None

        Notes:
            - If `element_id` is empty, the update is not performed and a message is
              printed.
            - Sends a PUT request to the Labfolder API to update the data element group.
            - Prints a message indicating whether the update was successful or not,
              based on the response status code.
        """
        if self.element_id == "":
            print("No data element group ID provided.")
            return None
        content = {
            "id": self.element_id,
            "title": self.title,
            "data_elements": [self.to_dict()],
            "locked": False,
        }
        response = requests.put(
            f"{user_info.api_address}elements/data/{self.element_id}",
            headers=user_info.auth_token,
            json=content,
            timeout=10,
        )
        _handle_response(self, response)
        return None

    def __labfolder_dict__(self):
        """Generate a dictionary representation of the object for labfolder export.

        Returns:
            dict: A dictionary containing the object's type, title, and a list of its
                children's dictionary representations.
        """
        return {
            "type": self.type,
            "title": self.title,
            "children": [child.to_dict() for child in self.children],
        }

    def __repr__(self):
        """Return a string representation of the DataElementGroup object.

        Returns:
            str: A string showing the title and children of the DataElementGroup.
        """
        return f"DataElementGroup(title={self.title}, children={self.children})"


class TableElement:
    """
    TableElement represents a table element in Labfolder, supporting loading,
    conversion, and synchronization with Labfolder's API.

    Args:
        user_info (LabFolderUserInfo): The user information object containing
            authentication and API details.
        element_id (str, optional): The unique identifier for the table element
            (default is '').
        entry_id (str, optional): The entry identifier associated with the table
            (default is '').
        table (dict, optional): The table data, either as a dictionary or pandas
            DataFrame (default is {}).
        import_as_pd (bool, optional): Whether to import the table as a pandas
            DataFrame (default is True).
        header (bool, optional): Whether the first row should be treated as a
            header (default is True).

    Attributes:
        type (str): The type of the element, always 'TABLE'.
        entry_id (str): The entry identifier associated with the table.
        element_id (str): The unique identifier for the table element.
        table (dict): The table data, either as a dictionary or pandas DataFrame.
        creation_date (str): The creation date of the table element.
        owner_id (str): The user ID of the table's owner.
        title (str): The title of the table.
    """

    def __init__(
        self,
        user_info: LabFolderUserInfo,
        element_id="",
        entry_id="",
        table={},
        import_as_pd=True,
        header=True,
    ):
        """
        Initialize a table data element.

        Args:
            user_info (LabFolderUserInfo): The user information object, typically
            representing the current user.
            element_id (str, optional): The unique identifier for the table element.
            Default is an empty string.
            entry_id (str, optional): The identifier for the entry to which this table
            belongs. Default is an empty string.
            table (dict, optional): The table data, represented as a dictionary.
            Default is an empty dictionary.
            import_as_pd (bool, optional): Whether to import the table as a pandas
            DataFrame. Default is True.
            header (bool, optional): Whether the table includes a header row.
            Default is True.

        Notes:
            If `user_info` is provided and `element_id` is not empty, the table is
            loaded and optionally converted to a pandas DataFrame.
            The `owner_id` is set from `user_info.user_id` if available, otherwise
            left empty.
            ----
            Check if `owner_id` can be obtained from another source, especially when
            `user_info` does not represent the table owner.
        """
        self.type = "TABLE"
        self.entry_id = entry_id
        self.element_id = element_id
        self.table = table
        self.creation_date = ""
        self.owner_id = user_info.user_id if user_info is not None else ""
        self.title = ""
        if isinstance(user_info, LabFolderUserInfo) and element_id != "":
            self.load_table(user_info, to_pd=import_as_pd, header=header)
            if import_as_pd:
                self.table_to_pd(header=header)
        # TODO: Check if owner_id can be taken from somewhere else.
        # Case: user_info is not from table owner.

    def load_table(self, user_info: LabFolderUserInfo, to_pd=True, header=True):
        """
        Load a table element from LabFolder using the provided user information.

        Args:
            user_info (LabFolderUserInfo): The user information object containing
            authentication details and API address.
            to_pd (bool, optional): If True, convert the loaded table to a pandas
            DataFrame (default is True).
            header (bool, optional): If True, treat the first row as the header when
            converting to pandas DataFrame (default is True).

        Returns:
            None or pandas.DataFrame: Returns None if the table ID is not provided or
            user_info is invalid. If `to_pd` is True, the table is converted to a
            pandas DataFrame and stored in the object.

        Notes:
            - Updates the object's `entry_id`, `creation_date`, `owner_id`, `title`,
              and `table` attributes.
            - Requires a valid LabFolderUserInfo object for authentication.
        """
        if self.element_id == "":
            print("No table ID provided.")
            return None
        if not isinstance(user_info, LabFolderUserInfo):
            print("Not logged into Labfolder. User information required.")
            return None

        table = requests.get(
            f"{user_info.api_address}elements/table/{self.element_id}",
            headers=user_info.auth_token,
            timeout=10,
        ).json()
        self.entry_id = table["entry_id"]
        self.creation_date = table["creation_date"]
        self.owner_id = table["owner_id"]
        self.title = table["title"]
        self.table = table["content"]["sheets"]
        if to_pd:
            self.table_to_pd(header=header)
        return None

    def write_to_labfolder(
        self, user_info: LabFolderUserInfo, entry_id="", header=True
    ):
        """
        Writes the current table to Labfolder as a new table element.

        Args:
            user_info (LabFolderUserInfo): An object containing user authentication
            information and API address for Labfolder.
            entry_id (str, optional): The ID of the Labfolder entry to which the table
            should be written. If not provided, uses self.entry_id.
            header (bool, default True): Whether to include the header in the exported
            table content.

        Returns:
            None

        Notes:
            - If neither `entry_id` nor `self.entry_id` is provided, the function
              prints a message and returns None.
            - If `self.table` is None, the function prints a message and returns None.
            - If the table cannot be converted to export format, the function prints a
              message and returns None.
            - On successful upload (HTTP 201), sets `self.element_id` to the returned
              element ID.
            - Prints status messages for success or failure.
        """
        if entry_id == "" and self.entry_id != "":
            entry_id = self.entry_id
        elif entry_id == "" and self.entry_id == "":
            print("No entry_id provided.")
            return None
        if self.table is None:
            print("No table to write.")
            return None
        table_content = self.convert_pd_to_export(header=header)
        if table_content is None:
            print("Could not convert table to export format.")
            return None
        response = requests.post(
            f"{user_info.api_address}elements/table",
            headers=user_info.auth_token,
            timeout=10,
            json={
                "entry_id": entry_id,
                "title": self.title,
                "content": table_content,
                "locked": False,
            },
        )
        status = _handle_response(self, response, return_status=True)
        if status:
            self.element_id = response.json()["id"]
        return None

    def update_on_labfolder(self, user_info: LabFolderUserInfo, header=True):
        """
        Updates the table element on Labfolder using the provided user information.

        Args:
            user_info (LabFolderUserInfo): An object containing user authentication
            information and API address for Labfolder.
            header (bool, optional): Whether to include the header in the exported
            table content (default is True).

        Returns:
            None

        Notes:
            - If `self.table` is None, the function prints a message and returns None.
            - If the table cannot be converted to export format, the function prints a
              message and returns None.
            - Sends a PUT request to update the table element on Labfolder.
            - Prints the result of the update operation.
        """
        if self.table is None:
            print("No table to write.")
            return None
        table_content = self.convert_pd_to_export(header=header)
        if table_content is None:
            print("Could not convert table to export format.")
            return None
        response = requests.put(
            f"{user_info.api_address}elements/table/{self.element_id}",
            headers=user_info.auth_token,
            timeout=10,
            json={
                "entry_id": self.entry_id,
                "id": self.element_id,
                "content": table_content,
                "locked": False,
            },
        )
        _handle_response(self, response)
        return None

    def table_to_pd(self, header=True, in_place=True):
        """
        Converts the internal table representation to pandas DataFrames.
        If the table is already converted to pandas DataFrames, the function prints
        a message and returns None. Otherwise, it processes each sheet in the table,
        converting its data to a pandas DataFrame, handling missing values, and
        optionally using the first row as the header.

        Args:
            header (bool, optional): Whether to use the first row of each table as the
            header (default is True).
            in_place (bool, optional): If True, modifies the internal table in place.
            If False, returns a new table with DataFrames (default is True).

        Returns:
            dict or None: If `in_place` is False, returns a dictionary mapping sheet
            names to their corresponding pandas DataFrames. If `in_place` is True,
            modifies the internal table and returns None.

        Notes:
            - If the table is already in pandas DataFrame format, the function prints a
              message and returns None.
            - Handles missing values by filling with NaN and dropping rows/columns that
              are entirely NaN.
            - If `header` is True, the first row is used as column headers and removed
              from the data.
        """
        if any(isinstance(element, pd.DataFrame) for element in self.table.values()):
            print("Table already converted to pandas DataFrame.")
            print(self.table)
            return None

        def table_to_pandas(table, header: bool):
            """Converts a nested dictionary table structure to a pandas DataFrame.

            Args:
                table (dict): A nested dictionary where each key represents a row, and each value
                    is a dictionary
                    mapping column names to dictionaries containing a 'value' key.
                header (bool): If True, the first row of the DataFrame is used as the header
                    (column names).

            Returns:
                pd.DataFrame: A pandas DataFrame constructed from the input table, with empty
                    rows and columns removed.
                    If `header` is True, the first row is used as column headers.

            Notes:
                - Any missing values are filled with `numpy.nan`.
                - Rows and columns that are entirely empty are dropped.
                - The DataFrame index is reset after processing.
            """
            data = []
            for row in table.keys():
                try:
                    data.append(
                        {
                            col: table[row][col].get("value", None)
                            for col in table[row].keys()
                        }
                    )
                except KeyError:
                    data.append({col: None for col in table[row].keys()})
            df = pd.DataFrame(data)
            df.fillna(np.nan, inplace=True)
            df.dropna(axis=0, how="all", inplace=True)
            df.dropna(axis=1, how="all", inplace=True)
            if header:
                df.columns = df.iloc[0].tolist()
                df.drop(0, inplace=True)
            df.sort_index(inplace=True)
            df.reset_index(drop=True, inplace=True)
            return df

        if in_place:
            table = self.table
        else:
            table = deepcopy(self.table)

        table = {
            sheet: table_to_pandas(
                self.table[sheet]["data"]["dataTable"], header=header
            )
            for sheet in self.table.keys()
        }
        if not in_place:
            return table

    def add_sheet(self, sheet_name: str, table: pd.DataFrame) -> None:
        """Add a new sheet to the table dictionary with the given sheet name and DataFrame.

        Args:
            sheet_name (str): The name of the sheet to add.
            table (pandas.DataFrame): The DataFrame to associate with the given sheet name.

        Returns:
            None: This method does not return anything. If the input is not a DataFrame,
            prints a message and returns None.

        Notes:
            If `self.table` is None, it initializes it as an empty dictionary before
            adding the new sheet.
        """

        if not isinstance(table, pd.DataFrame):
            print("Table is not a pandas DataFrame.")
            return None
        if self.table is None:
            self.table = {}
        self.table.update({sheet_name: table})
        return None

    def table_to_dict(self):
        """Converts all pandas DataFrame objects in the `self.table` dictionary to
        dictionaries.

        Iterates over each sheet in `self.table`. If the value associated with a
        sheet is a pandas DataFrame,
        it is converted to a dictionary using the DataFrame's `to_dict()` method.
        If the value is not a DataFrame,
        a message is printed indicating the sheet is not a DataFrame.

        Returns:
            None

        Notes:
            This method modifies `self.table` in place, replacing DataFrame objects
            with their dictionary representations.
        """
        for sheet in self.table.keys():
            if isinstance(self.table[sheet], pd.DataFrame):
                self.table[sheet] = self.table[sheet].to_dict()
            else:
                print(f"Sheet {sheet} is not a pandas DataFrame.")

    def convert_pd_to_export(self, header=True):
        """
        Converts the internal table of pandas DataFrames into a dictionary format
        suitable for export.

        If the table elements are not already pandas DataFrames, attempts to convert
        them. Each sheet in the table is processed to optionally include column headers
        as the first row, replace NaN and infinite values with None, and convert the
        DataFrame to a nested dictionary structure. The resulting dictionary contains
        metadata about each sheet, including its name, row count, column count, and
        the data itself.

        Args:
            header (bool, optional): If True, includes the DataFrame's column headers as the
            first row in the exported data (default is True).

        Returns:
            dict or None: A dictionary with the structure:
            {
                'sheets': {
                sheet_name: {
                    'name': sheet_name,
                    'rowCount': int,
                    'columnCount': int,
                    'data': {
                    'dataTable': dict
                    }
                },
                ...
                }
            }
            Returns None if the table cannot be converted to pandas DataFrames.

        Notes:
            Any exceptions during conversion are caught and handled internally.
        """
        if not all(
            isinstance(element, pd.DataFrame) for element in self.table.values()
        ):
            try:
                self.table_to_pd()
            except (TypeError, ValueError, KeyError, AttributeError) as e:
                print(f"Table could not be converted to pandas DataFrame: {e}")
                return None
        tmp_table = deepcopy(self.table)
        table_content = {"sheets": {}}
        for sheet in tmp_table.keys():
            if header:
                tmp_table[sheet].index = tmp_table[sheet].index + 1
                tmp_table[sheet].loc[0] = tmp_table[sheet].columns
            tmp_table[sheet].sort_index(inplace=True)
            tmp_table[sheet].reset_index(drop=True, inplace=True)
            row_count = tmp_table[sheet].shape[0]
            column_count = tmp_table[sheet].shape[1]
            tmp_table[sheet].columns = range(tmp_table[sheet].shape[1])
            tmp_table[sheet].replace(
                {np.nan: None, np.inf: None, -np.inf: None}, inplace=True
            )
            tmp_table[sheet] = tmp_table[sheet].to_dict(orient="index")
            for row in tmp_table[sheet].keys():
                for col in tmp_table[sheet][row].keys():
                    tmp_table[sheet][row][col] = {"value": tmp_table[sheet][row][col]}
            table_content["sheets"].update(
                {
                    sheet: {
                        "name": sheet,
                        "rowCount": row_count,
                        "columnCount": column_count,
                        "data": {"dataTable": tmp_table[sheet]},
                    }
                }
            )
        return table_content

    def dict_to_table(self, in_place=True):
        """Converts dictionary entries in the table attribute to pandas DataFrames.

        Args:
            in_place (bool, optional): If True, modifies the `table` attribute in place.
            If False, returns a new table with the same structure but with dictionary
            entries converted to DataFrames.

        Returns:
            dict or None: If `in_place` is False, returns a new table dictionary with
            DataFrames. If `in_place` is True, modifies the table in place and
            returns None.

        Notes:
            Only dictionary entries in the table are converted to DataFrames.
            Non-dictionary entries are left unchanged and a message is printed for
            each such entry.
        """
        if in_place:
            table = self.table
        else:
            table = deepcopy(self.table)

        for sheet in table.keys():
            if isinstance(table[sheet], dict):
                table[sheet] = pd.DataFrame(table[sheet])
            else:
                print(f"Sheet {sheet} is not a dictionary.")
        if not in_place:
            return table
        return None

    def to_dict(self, in_place=True):
        """Convert the internal table attribute to a dictionary representation.

        Args:
            in_place (bool, optional): If True, modifies the internal `table` attribute in place.
            If False, operates on a deep copy and returns the converted dictionary.
            Default is True.

        Returns:
            dict or None: If `in_place` is False, returns the converted dictionary
            representation of the table. If `in_place` is True, returns None.

        Notes:
            - Each sheet in the table that is a pandas DataFrame is converted to a
              dictionary using `to_dict()`.
            - Sheets that are already dictionaries are left unchanged.
            - If a sheet is neither a DataFrame nor a dictionary, a message is printed.
        """
        if in_place:
            table = self.table
        else:
            table = deepcopy(self.table)
        for sheet in table.keys():
            if isinstance(table[sheet], pd.DataFrame):
                table[sheet] = table[sheet].to_dict()
            elif isinstance(table[sheet], dict):
                pass
            else:
                print(f"Sheet {sheet} is not a pandas DataFrame or dictionary.")
        if not in_place:
            return table
        return None

    def __repr__(self):
        """Return a string representation of the TableElement instance.

        Returns:
            str: A string describing the TableElement, including its type,
            id, and table.
        """
        return (
            f"TableElement(type={self.type}, id={self.element_id}, table={self.table})"
        )


def _handle_response(
    element,
    response: requests.Response,
    silent: bool = False,
    return_status: bool = False,
) -> bool | None:
    """Handle the response from Labfolder API requests.

    Args:
        response (int): The HTTP status code from the response.
        element_name (str): The name of the element being processed.
        silent (bool): If True, suppresses print statements.
    Returns:
        None
    """
    response_text = ""
    if response == 200 and response.request.method == "PUT":
        response_text = element.type, "updated on Labfolder."
        status = True
    elif response == 201 and response.request.method == "POST":
        response_text = element.type, "written to Labfolder."
        status = True
    else:
        if response.request.method == "POST":
            response_text = (
                f"{element.type} could not be written to Labfolder. "
                f"Status code: {response.status_code}"
            )
        elif response.request.method == "PUT":
            response_text = (
                f"{element.type} could not be updated on Labfolder. "
                f"Status code: {response.status_code}"
            )
        status = False
    if not silent:
        print(response_text)
    if return_status:
        return status
    return None


def parse_data_element(
    element: dict, user_info: LabFolderUserInfo
) -> (
    DataElementGroup
    | FileElement
    | ImageElement
    | TextElement
    | DescriptiveDataElement
    | DataElement
    | TableElement
    | None
):
    """
    Parse a data element from Labfolder and return the corresponding object.

    Args:
        element (dict): A dictionary representing the data element to be parsed.
        user_info (LabFolderUserInfo): An object containing the user's Labfolder
            API address and authentication token.

    Returns:
        DataElementGroup | FileElement | ImageElement | TextElement |
        DescriptiveDataElement | DataElement | TableElement | None
    """
    return_element = None
    t = element.get("element_type", "")
    if t == "DATA_ELEMENT_GROUP":
        group = DataElementGroup(title=element.get("title", ""))
        for child in element.get("children", []):
            c = parse_data_element(child, user_info)
            if c is not None:
                group.add_child(c)
        return group
    if t == "FILE":
        return_element = FileElement(element_id=element.get("id", ""))
    if t == "IMAGE":
        return_element = ImageElement(
            element_id=element.get("id", ""),
            title=element.get("title", ""),
            original_file_content_type=element.get(
                "original_file_content_type", "image/png"
            ),
        )
    if t == "TEXT":
        return_element = TextElement(
            element_id=element.get("id", ""), content=element.get("content", "")
        )
    if t == "DESCRIPTIVE_DATA":
        return_element = DescriptiveDataElement(
            element_id=element.get("id", ""),
            title=element.get("title", ""),
            description=element.get("description", ""),
        )
    if t == "DATA":
        return_element = DataElement(
            element_id=element.get("id", ""), description=element.get("description", "")
        )
    if t == "TABLE":
        return_element = TableElement(
            user_info=user_info,
            element_id=element.get("id", ""),
            entry_id=element.get("entry_id", ""),
        )
    if return_element is None:
        print(f"Unknown element type: {element}")
    return return_element
