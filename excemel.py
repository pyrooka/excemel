import os
import sys
import json
import copy
import typing
import argparse
import operator
import functools
import xml.dom.minidom as md
import xml.etree.ElementTree as et

import openpyxl


# Name of the config file.
CONFIG_NAME = "config.json"
# Merge key.
MERGE_KEY = "__MERGE__"
# Default config struct.
DEFAULT_CONFIG = """
{
  "order": "row",
  "from": 1,
  "struct": {
    "Main": {
        "columns": {"col": 1}
    }
  }
}
"""

def parse_args() -> dict:
    """Parses the command line arguments.

    Returns:
        dict: the command line arguments as a dictionary
    """

    # Parser.
    parser = argparse.ArgumentParser(description="Generate XML from Excel files base on the defined structure.")

    # Name arguments.
    parser.add_argument("--create-config", action="store_true", default=False, help="Create a default config.")
    parser.add_argument("input_file", action="store", default=None, help="input Excel file", nargs="?")
    parser.add_argument("output_file", action="store", default=None, help="output XML file", nargs="?")

    # Little trick to make the print help usable from the caller function.
    args = vars(parser.parse_args())
    args["_print_help_function"] = parser.print_help

    return args

def create_default_config():
    """Creates the default config file with the struct. Doesn't overwrite existing ones.
    """

    # Check if the config exists.
    config_exists = os.path.isfile(CONFIG_NAME)
    if config_exists:
        print("Config file is already exists. To generate the default remove the current one.")
        return

    # If not, write it.
    with open(CONFIG_NAME, "w") as file:
        file.write(DEFAULT_CONFIG)

def read_config(config_name: str) -> dict:
    """Reads the config struct JSON and parse it.

    Returns:
        dict: config struct as a dictionary
    """

    config_exists = os.path.isfile(config_name)
    if not config_exists:
        raise Exception(f"the config file ({CONFIG_NAME}) isn't exists. Please create a new one with the --create-config flag")

    with open(config_name, "r") as file:
        text = file.read()

        # Parse JSON to dict.
        try:
            struct = json.loads(text)
            return struct
        except:
            raise Exception("config file is not valid JSON")

def get_worksheet(filename: str) -> openpyxl.worksheet.worksheet.Worksheet:
    """Loads the active worksheet from an excel file.

    Args:
        filename (str): name of the excel file

    Returns:
        openpyxl.worksheet.worksheet.Worksheet: active worksheet in the excel file
    """

    # Open the excel workbook.
    try:
        workbook = openpyxl.load_workbook(filename)
    except Exception as e:
        raise Exception("cannot open the excel file: " + str(e))

    # Return the active worksheet.
    return workbook.active

def get_path_recursive(struct: typing.Union[dict, list], current_path: list, paths: dict):
    """Iterates over the struct recursively and if an element found saves it to the path.

    Args:
        struct (typing.Union[dict, list]): various type for element finding
        current_path (list): where we are at the moment
        paths (dict): saves the path of the element here
    """

    if type(struct) is dict:
        if "col" in struct:
            if "merge" in struct:
                paths["merge"] = current_path.copy()

            paths[struct["col"]] = current_path.copy()
            current_path.pop()
        else:
            for key in struct:
                current_path.append(key)
                get_path_recursive(struct[key], current_path, paths)

    elif type(struct) is list:
        current_path.append(0)
        get_path_recursive(struct[0], current_path, paths)

def set_nested_value(struct: dict, path: list, value: str):
    """Sets the value on a path for the given struct.

    Args:
        struct (dict): base
        path (list): value location in the struct
        value (str): value for set
    """

    # Get the last key from the path and remove it also.
    last_key = path.pop()
    # Get the object contains our key.
    element = functools.reduce(operator.getitem, path, struct)
    # Set the value for the key in the object.
    element[last_key] = value

def create_final_struct(structs: list) -> dict:
    """Creates the final struct from the existing ones.

    Args:
        structs (list): list of created structs

    Returns:
        dict: the final struct
    """

    # Get the first one.
    final_struct = structs[0]

    # Iterate over the others.
    for i in range(1, len(structs)):
        # Merge the new struct into the one before.
        merge_structs(final_struct, structs[i])

    return final_struct

def merge_structs(struct_one: typing.Union[dict, list], struct_two: typing.Union[dict, list]):
    """Merges the second struct into the first one.

    Args:
        struct_one (typing.Union[dict, list]): merge into this
        struct_two (typing.Union[dict, list]): merge this
    """

    if type(struct_one) is dict:
        for key in struct_two:
            if type(struct_two[key]) is list:
                merge_structs(struct_one[key], struct_two[key])
            elif key not in struct_one:
                struct_one[key] = struct_two[key]
            else:
                merge_structs(struct_one[key], struct_two[key])

    elif type(struct_one) is list:
        first_one_elem = struct_one[0]
        first_two_elem = struct_two[0]

        if type(first_one_elem) is list:
            merge_structs(first_one_elem, first_two_elem)
        else:
            # Dict should be contains only one element.
            # NOTE: check it!!!
            key = list(first_one_elem.keys())[0]
            elem_one = first_one_elem[key]
            elem_two = first_two_elem[key]

            if MERGE_KEY in elem_one:
                merge_by = elem_one[MERGE_KEY]
                if elem_one[merge_by] == elem_two[merge_by]:
                    merge_structs(first_one_elem, first_two_elem)
                    return

            struct_one += struct_two

def build_xml_recursive(struct: typing.Union[dict, list, str], parent: et.Element):
    """Build XML from the struct recursively.

    Args:
        struct (typing.Union[dict, list, str]): Various type of object can be passed.
        parent (et.Element): The previous XML element object which is the parent for the current one.
    """

    # Get the type of the struct.
    struct_type = type(struct)
    # If it's a dictionary
    if struct_type is dict:
        # iterate over its items.
        for key, val in struct.items():
            # Exclude our merge key.
            if key == MERGE_KEY:
                continue

            # Create a sub element for the parent with the current object
            se = et.SubElement(parent, key)
            # and start the whole process again with it.
            build_xml_recursive(val, se)
    # If the struct is a list
    elif struct_type is list:
        # iterate over its elements
        for s in struct:
            # and call this function again with the same parent.
            build_xml_recursive(s, parent)

    # Otherwise save the object as a text for the parent.
    else:
        parent.text = str(struct)

def main():
    # Get the arguments from the CLI.
    args = parse_args()

    # Helper function for print usage.
    print_help = args["_print_help_function"]

    # Should we create a new config?
    if args["create_config"] is True:
        create_default_config()
        return

    # Check if the input and output file provided.
    if args["input_file"] is None or args["output_file"] is None:
        print("Not enough file.")
        print_help()
        return

    # Read the config file.
    try:
        config = read_config(CONFIG_NAME)
        order = config["order"]
        from_value = config["from"]
        struct = config["struct"]
    except Exception as e:
        print("Error while reading the config: " + str(e))
        return

    # Load the excel.
    try:
        ws = get_worksheet(args["input_file"])
    except Exception as e:
        print("Error while loading the worksheet: " + str(e))
        return

    # Get the location of the cells.
    temp_path = []
    paths = {}
    get_path_recursive(struct, temp_path, paths)

    structs = []

    for i, row in enumerate(ws.rows):
        if i+1 < from_value:
            continue

        new_struct = struct.copy()

        if "merge" in paths:
            merge_path = paths["merge"].copy()
            merge_by = merge_path.pop()
            merge_path.append(MERGE_KEY)
            set_nested_value(new_struct, merge_path, merge_by)

        for j, col in enumerate(row):
            path_key = j + 1
            if path_key in paths:
                set_nested_value(new_struct, paths[path_key].copy(), col.value)

        structs.append(copy.deepcopy(new_struct))

    done = create_final_struct(structs)

    main_key = list(done.keys())[0]
    xml = et.Element(main_key)
    build_xml_recursive(done[main_key], xml)

    # Prettify.
    dom = md.parseString(et.tostring(xml, encoding="unicode"))
    pretty_xml = dom.toprettyxml()

    with open(args["output_file"], "w") as file:
        file.write(pretty_xml)

if __name__ == "__main__":
    main()