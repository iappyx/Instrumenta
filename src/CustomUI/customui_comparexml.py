"""
compare_customui.py  -  Semantic comparison of two customUI XML files
=====================================================================

Parses both files into element trees and compares them structurally,
ignoring cosmetic differences like:
  - self-closing vs. explicit closing tags  (<x /> vs <x></x>)
  - attribute order
  - whitespace-only text nodes / indentation
  - the zero-width-space (&#x200B;) vs empty-string label equivalence

Usage:
    python compare_customui.py customUI.xml CustomUI_output.xml
"""

import sys
import xml.etree.ElementTree as ET

ZWS = "\u200b"   # zero-width space  (&#x200B;)


def short_tag(tag: str) -> str:
    """Strip namespace URI for readable output."""
    return tag.split("}", 1)[1] if tag.startswith("{") else tag


def normalise_value(v: str) -> str:
    """Treat zero-width space the same as an empty string."""
    return "" if v == ZWS else v


def normalise_attribs(attribs: dict) -> dict:
    return {k: normalise_value(v) for k, v in attribs.items()}


def elem_signature(elem: ET.Element) -> str:
    """Short human-readable label for an element, used in diff paths."""
    tag = short_tag(elem.tag)
    attrs = normalise_attribs(elem.attrib)
    ident = attrs.get("id") or attrs.get("idMso") or ",".join(
        f"{k}={attrs[k]}" for k in sorted(attrs)
    )
    return f"{tag}[{ident}]"


def compare_trees(a: ET.Element, b: ET.Element, path: str = "root") -> list:
    """
    Recursively compare two elements.
    Returns a list of difference descriptions (empty list = identical).
    """
    diffs = []

    # Tag
    if a.tag != b.tag:
        diffs.append(f"{path}: tag mismatch  <{short_tag(a.tag)}>  vs  <{short_tag(b.tag)}>")
        return diffs   # no point going deeper

    # Attributes
    a_attrs = normalise_attribs(a.attrib)
    b_attrs = normalise_attribs(b.attrib)

    for k in sorted(set(a_attrs) | set(b_attrs)):
        av = a_attrs.get(k)
        bv = b_attrs.get(k)
        if av != bv:
            diffs.append(f"{path}  @{k}:  {repr(av)}  vs  {repr(bv)}")

    # Children
    a_ch = list(a)
    b_ch = list(b)

    if len(a_ch) != len(b_ch):
        diffs.append(f"{path}: child count  {len(a_ch)}  vs  {len(b_ch)}")

    for i, (ca, cb) in enumerate(zip(a_ch, b_ch)):
        diffs.extend(compare_trees(ca, cb, path=f"{path} > [{i}]{elem_signature(ca)}"))

    return diffs


def main():
    if len(sys.argv) < 3:
        print("Usage: python compare_customui.py <file_a.xml> <file_b.xml>")
        sys.exit(1)

    path_a, path_b = sys.argv[1], sys.argv[2]

    try:
        a = ET.parse(path_a).getroot()
    except Exception as e:
        print(f"ERROR reading {path_a}: {e}")
        sys.exit(1)

    try:
        b = ET.parse(path_b).getroot()
    except Exception as e:
        print(f"ERROR reading {path_b}: {e}")
        sys.exit(1)

    diffs = compare_trees(a, b)

    if not diffs:
        print(f"✓  Files are semantically identical.")
        print(f"   {path_a}")
        print(f"   {path_b}")
    else:
        print(f"✗  {len(diffs)} difference(s) found:\n")
        for d in diffs:
            print(f"  • {d}")
        sys.exit(1)


if __name__ == "__main__":
    main()
