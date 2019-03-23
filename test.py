import pathlib
import difflib

from excemel import *

def main():
    test_dir = pathlib.Path("./test")

    sub_dirs = [sub_dir for sub_dir in test_dir.iterdir() if sub_dir.is_dir()]

    for sub_dir in sub_dirs:
        # Read the config file.
        try:
            config = read_config(pathlib.PurePath(sub_dir, "config.json"))
            order = config["order"]
            from_value = config["from"]
            struct = config["struct"]
        except Exception as e:
            print("Error while reading the config: " + str(e))
            return

        # Load the excel.
        try:
            ws = get_worksheet(pathlib.PurePath(sub_dir, "test1.xlsx"))
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

        with open(pathlib.PurePath(sub_dir, "test1.xml"), "r") as file:
            content = file.read()

            print(pretty_xml)

            if content != pretty_xml:
                print("TEST FAILED!")
                diff = difflib.ndiff(content, pretty_xml)
                print([li for li in diff if li[0] != ' '])
            else:
                print("TEST WAS SUCCESS!")

if __name__ == "__main__":
    main()