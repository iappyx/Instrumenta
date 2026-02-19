"""
customUI XML -> JSON converter  (with $ref deduplication)
=========================================================

Output JSON structure
---------------------
{
  "definitions": {
      "<base_id>": { "_tag": "...", "id": "<base_id>", ...all attrs...,
                     "supertip": "", "screentip": "",
                     "_children": [...] }          <- optional
  },
  "layout": {
      "_tag": "customUI",
      ...root attrs...,
      "_children": [
          { "_tag": "ribbon", "_children": [
              { "_tag": "tabs", "_children": [
                  { "_tag": "tab", ...,  "_children": [
                      { "_tag": "group", ..., "_children": [
                          { "$ref": "BoldSplitButton", "_tag": "splitButton" },
                          ...
                      ]},
                      ...
                  ]}
              ]}
          ]}
      ]
  }
}

Deduplication rules
-------------------
* Elements that carry an "id" attribute are deduplicated by their *base id*
  (TabView prefix stripped).
* The first occurrence is stored verbatim in definitions[base_id] (with the
  base id as the canonical id value).
* Every subsequent occurrence – in any tab / view – is replaced in the layout
  by a lightweight stub:  { "$ref": "<base_id>", "_tag": "<tag>" }
* Elements with only an idMso (built-in Office controls) are also deduplicated
  the same way (idMso itself is the key; it never has a TabView prefix).
* Structural / container elements without any id (ribbon, tabs, tab, group,
  box, separator, menuSeparator, buttonGroup, …) are kept inline in the
  layout.

Tooltip placeholders
--------------------
Empty "supertip" and "screentip" keys are added to every element that
supports them so they are easy to fill in without knowing the spec.
"""

import json
import sys
import xml.etree.ElementTree as ET

NS = "http://schemas.microsoft.com/office/2006/01/customui"
TABVIEW_PREFIX = "TabView"

# Office ribbon elements that accept supertip / screentip
SUPPORTS_TOOLTIP = {
    "button", "toggleButton", "splitButton", "gallery", "comboBox",
    "control", "menu", "dropDown", "editBox", "checkBox", "labelControl",
}

# Elements that have an identity of their own (id or idMso) and can be
# deduplicated.  Everything else is treated as structural/container.
DEDUPLICATED_TAGS = {
    "button", "toggleButton", "splitButton", "gallery", "comboBox",
    "control", "menu", "dropDown", "editBox", "checkBox", "labelControl",
    "separator", "menuSeparator", "buttonGroup",
}


def strip_ns(tag: str) -> str:
    return tag.split("}", 1)[1] if tag.startswith("{") else tag


def base_id(eid: str) -> str:
    """Strip the TabView prefix to obtain the canonical base id."""
    return eid[len(TABVIEW_PREFIX):] if eid.startswith(TABVIEW_PREFIX) else eid


def get_elem_id(attribs: dict):
    """Return (key, value, canonical_base) for the element's identity, or None."""
    if "id" in attribs:
        eid = attribs["id"]
        return ("id", eid, base_id(eid))
    if "idMso" in attribs:
        eid = attribs["idMso"]
        return ("idMso", eid, eid)          # idMso is never prefixed
    return None


def elem_to_dict(elem, definitions: dict, seen: set):
    """
    Recursively convert an XML element to a JSON-serialisable dict.
    Side-effects: populates `definitions` and `seen`.
    Returns either:
      - a full node dict  (structural/container elements)
      - a $ref stub dict  (deduplicated leaf/compound elements)
    """
    tag = strip_ns(elem.tag)
    attribs = dict(elem.attrib)

    # Preserve zero-width space label as explicit marker so round-trip is lossless
    if attribs.get("label") == "\u200b":
        attribs["label"] = "&#x200B;"

    # Add tooltip placeholders to supported elements
    if tag in SUPPORTS_TOOLTIP:
        attribs.setdefault("supertip", "")
        attribs.setdefault("screentip", "")

    # Recurse into children regardless – needed to build definitions
    children = [elem_to_dict(ch, definitions, seen) for ch in elem]

    identity = get_elem_id(attribs)

    if identity and tag in DEDUPLICATED_TAGS:
        id_key, id_val, bid = identity

        if bid in seen:
            # Already defined → emit a $ref stub in the layout
            return {"$ref": bid, "_tag": tag}

        # First time → store canonical definition
        seen.add(bid)
        definition = {"_tag": tag}
        # Write all attributes with canonical id
        for k, v in attribs.items():
            if k == "id":
                definition["id"] = bid          # normalise to base id
            else:
                definition[k] = v
        if children:
            definition["_children"] = children
        definitions[bid] = definition

        # Layout gets a stub too (keeps the layout clean and uniform)
        return {"$ref": bid, "_tag": tag}

    # Structural / container element → keep fully inline in layout
    node = {"_tag": tag}
    node.update(attribs)
    if children:
        node["_children"] = children
    return node


def convert(src_path: str, dst_path: str):
    tree = ET.parse(src_path)
    root = tree.getroot()

    definitions: dict = {}
    seen: set = set()

    # Root element (customUI) is structural
    root_tag = strip_ns(root.tag)
    layout_root = {"_tag": root_tag}
    layout_root.update(root.attrib)
    layout_root["_children"] = [
        elem_to_dict(ch, definitions, seen) for ch in root
    ]

    output = {
        "definitions": definitions,
        "layout": layout_root,
    }

    with open(dst_path, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print(f"Exported  {src_path}  ->  {dst_path}")
    print(f"  {len(definitions)} unique element definitions")
    # Count stubs
    raw = json.dumps(output)
    stubs = raw.count('"$ref"')
    print(f"  {stubs} $ref stubs in layout  ({stubs - len(definitions) + stubs} bytes saved)")


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "CustomUI.xml"
    dst = sys.argv[2] if len(sys.argv) > 2 else "CustomUI.json"
    convert(src, dst)


if __name__ == "__main__":
    main()
