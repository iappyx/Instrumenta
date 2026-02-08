import xml.etree.ElementTree as ET
from collections import defaultdict

XML_FILE = "CustomUI.xml"
OUTPUT_FILE = "generated_features.txt"

NS = {"ns": "http://schemas.microsoft.com/office/2006/01/customui"}

def build_parent_map(root):
    return {c: p for p in root.iter() for c in p}

def get_parent(element, parent_map):
    return parent_map.get(element)

def find_ancestor(element, tag, parent_map):
    parent = get_parent(element, parent_map)
    while parent is not None:
        if parent.tag.endswith(tag):
            return parent
        parent = get_parent(parent, parent_map)
    return None

def get_splitbutton_friendly_label(split, ns):
    if split is None:
        return None

    # 1. If splitbutton has its own label, use it
    if "label" in split.attrib:
        return split.get("label")

    # 2. Otherwise, check first child for idMso
    first_child = next(iter(split), None)
    if first_child is not None:
        if "label" in first_child.attrib:
            return first_child.get("label")
        if "idMso" in first_child.attrib:
            return first_child.get("idMso")

    return None


def extract_features(root, parent_map):
    """Extract ALL features (normal + tabview)."""
    features = {}

    for ctrl in root.findall(".//*[@onAction]", NS):
        ctrl_id = ctrl.get("id")
        if ctrl_id is None or ctrl_id.startswith("idMso"):
            continue

        ctrl_label = ctrl.get("label")
        ctrl_action = ctrl.get("onAction")

        # Find group
        group = find_ancestor(ctrl, "group", parent_map)
        group_label = group.get("label") if group is not None else ""

        # Find menu
        menu = find_ancestor(ctrl, "menu", parent_map)
        menu_label = menu.get("label") if menu is not None else None

        # Find splitbutton
        split = find_ancestor(ctrl, "splitButton", parent_map)
        split_label = get_splitbutton_friendly_label(split, NS)

        # Find tab
        tab = find_ancestor(ctrl, "tab", parent_map)
        tab_label = tab.get("label") if tab is not None else ""

        # Build hierarchy prefix
        if ctrl_id.startswith("TabView"):
            prefix = f"{tab_label}"
        else:
            prefix = tab_label  # normal view

        # Correct hierarchy order:
        # Tab > Group > Menu > Splitbutton
        hierarchy = prefix

        if group_label:
            hierarchy += f" > {group_label}"

        if menu_label:
            hierarchy += f" > Inside menu {menu_label}"

        if split_label:
            hierarchy += f" > Inside splitbutton {split_label}"

        features[ctrl_id] = {
            "label": ctrl_label,
            "action": ctrl_action,
            "group": group_label,
            "hierarchy": hierarchy
        }

    return features

# Parse XML
tree = ET.parse(XML_FILE)
root = tree.getroot()
parent_map = build_parent_map(root)

# Extract all features
all_features = extract_features(root, parent_map)

# Split into normal and tabview
normal = {fid: data for fid, data in all_features.items() if not fid.startswith("TabView")}
tabview = {fid: data for fid, data in all_features.items() if fid.startswith("TabView")}

# Match tabview features to normal ones
tabview_map = defaultdict(list)
for fid, data in tabview.items():
    base_id = fid.replace("TabView", "", 1)
    tabview_map[base_id].append(data["hierarchy"])

# Output
lines = []
for fid, data in normal.items():
    normal_h = data["hierarchy"]
    tab_h = "; ".join(tabview_map.get(fid, []))

    line = (
        f'AddFeature "{fid}", "{data["label"]}", "{data["action"]}", '
        f'"{data["group"]}", "{normal_h}", "{tab_h}"'
    )
    lines.append(line)

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"Done. Output written to {OUTPUT_FILE}")

