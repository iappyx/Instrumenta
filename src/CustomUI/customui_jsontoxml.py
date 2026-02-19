"""
JSON -> customUI XML converter  (with $ref resolution)
======================================================

Reads the JSON produced by xml_to_json.py and generates a valid
customUI.xml.

$ref resolution
---------------
When the layout contains a stub like:
    { "$ref": "BoldSplitButton", "_tag": "splitButton" }

the generator:
  1. Looks up definitions["BoldSplitButton"].
  2. Deep-copies the definition.
  3. Determines the correct "id" value for the current view:
       - regular / one-tab view  ->  id stays as the base id
       - TabView tabs            ->  id becomes "TabView" + base_id
  4. Emits the element with that id (and all other attrs / children
     from the definition).

View detection
--------------
The generator walks the layout.  When it enters a <tab> element, it
checks whether that tab's id starts with "TabView" to decide whether
to apply the prefix when resolving $refs.

Empty supertip / screentip values are omitted from the output XML so
the file stays clean until you actually fill them in.
"""

import copy
import json
import sys
from xml.dom import minidom
from xml.etree import ElementTree as ET

NS = "http://schemas.microsoft.com/office/2006/01/customui"
TABVIEW_PREFIX = "TabView"

INTERNAL_KEYS = {"_tag", "_children", "$ref"}
OMIT_IF_EMPTY = {"supertip", "screentip"}
OMIT_VALUES = {""}              # empty string only

# Tab-level ids that belong to the TabView multi-tab variant
TABVIEW_TAB_IDS = {
    "TabViewInstrumentaText",
    "TabViewInstrumentaShapes",
    "TabViewInstrumentaTables",
    "TabViewInstrumentaAdvanced",
}


def resolve_ref(stub: dict, definitions: dict, use_tabview: bool) -> dict:
    """
    Expand a $ref stub into a full element dict.
    Applies the TabView id prefix when use_tabview is True.
    """
    bid = stub["$ref"]
    tag = stub["_tag"]

    if bid not in definitions:
        raise KeyError(f"$ref '{bid}' not found in definitions")

    node = copy.deepcopy(definitions[bid])
    node["_tag"] = tag      # tag in layout takes precedence (should match)

    # Fix up the id
    if "id" in node:
        node["id"] = (TABVIEW_PREFIX + bid) if use_tabview else bid

    return node


def build_elem(node: dict, definitions: dict, parent=None,
               use_tabview: bool = False):
    """
    Recursively build an ET.Element from a layout node.
    Resolves $ref stubs on the fly.
    """
    # Resolve ref if needed
    if "$ref" in node:
        node = resolve_ref(node, definitions, use_tabview)

    tag = node["_tag"]
    qualified = f"{{{NS}}}{tag}"

    elem = ET.SubElement(parent, qualified) if parent is not None \
        else ET.Element(qualified)

    # Write attributes
    for k, v in node.items():
        if k in INTERNAL_KEYS:
            continue
        if k in OMIT_IF_EMPTY and v in OMIT_VALUES:
            continue
        # Translate explicit marker back to the actual unicode character
        v = v.replace("&#x200B;", "\u200b")
        elem.set(k, v)

    # Determine tabview context for children
    child_tabview = use_tabview
    if tag == "tab":
        tab_id = node.get("id", "")
        child_tabview = tab_id in TABVIEW_TAB_IDS

    # Recurse
    for child in node.get("_children", []):
        build_elem(child, definitions, parent=elem,
                   use_tabview=child_tabview)

    return elem


def pretty_xml(elem: ET.Element) -> str:
    raw = ET.tostring(elem, encoding="unicode")
    parsed = minidom.parseString(raw)
    result = parsed.toprettyxml(indent="   ", encoding="UTF-8").decode("utf-8")
    # Python's XML serialiser writes the raw U+200B character; replace it
    # with the explicit entity so the file stays readable and unambiguous.
    result = result.replace("\u200b", "&#x200B;")
    return result


def convert(src_path: str, dst_path: str):
    with open(src_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    definitions: dict = data["definitions"]
    layout: dict = data["layout"]

    # Register namespace so ET doesn't invent a prefix
    ET.register_namespace("", NS)

    root_elem = build_elem(layout, definitions, use_tabview=False)

    xml_str = pretty_xml(root_elem)

    with open(dst_path, "w", encoding="utf-8") as f:
        f.write(xml_str)

    print(f"Generated  {src_path}  ->  {dst_path}")


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "CustomUI.json"
    dst = sys.argv[2] if len(sys.argv) > 2 else "CustomUI_output.xml"
    convert(src, dst)


if __name__ == "__main__":
    main()
