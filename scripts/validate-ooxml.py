#!/usr/bin/env python3
"""Validate the XML parts of .xlsx files against the ECMA-376 / ISO/IEC 29500 schemas.

Usage: OOXML_SCHEMA_DIR=<dir> python3 scripts/validate-ooxml.py <xlsx files...>

<dir> must contain ISO-IEC29500-4_2016/*.xsd (sml.xsd, dml-chart.xsd, ...) and
ecma/fouth-edition/opc-*.xsd. The schemas are not shipped with this repository.

Exit codes: 0 = every part valid, 1 = validation errors, 2 = setup problem (no schemas, no lxml, no files).
The only error ignored is the known false positive for xml:space on <t> (the strict schema omits it,
Excel writes and expects it).
"""
import os
import sys
import zipfile

HINT = ("download ECMA-376 Part 1/4 schemas from ecma-international.org and set OOXML_SCHEMA_DIR "
        "(layout: <dir>/ISO-IEC29500-4_2016/*.xsd and <dir>/ecma/fouth-edition/opc-*.xsd)")


def fail_setup(msg):
    print(f"validate-ooxml: {msg}", file=sys.stderr)
    sys.exit(2)


root = os.environ.get("OOXML_SCHEMA_DIR", "")
if not root:
    fail_setup(f"OOXML_SCHEMA_DIR is not set; {HINT}")
ISO = os.path.join(root, "ISO-IEC29500-4_2016")
OPC = os.path.join(root, "ecma", "fouth-edition")
if not os.path.isfile(os.path.join(ISO, "sml.xsd")) or not os.path.isfile(os.path.join(OPC, "opc-relationships.xsd")):
    fail_setup(f"schemas not found under {root}; {HINT}")
try:
    from lxml import etree
except ImportError:
    fail_setup("python module lxml is not installed (pip install lxml)")
if len(sys.argv) < 2:
    fail_setup("usage: python3 scripts/validate-ooxml.py <xlsx files...>")

_cache = {}


def schema(path):
    if path not in _cache:
        _cache[path] = etree.XMLSchema(etree.parse(path))
    return _cache[path]


def schema_for(part):
    if part == "[Content_Types].xml":
        return os.path.join(OPC, "opc-contentTypes.xsd")
    if part.endswith(".rels"):
        return os.path.join(OPC, "opc-relationships.xsd")
    if part == "docProps/app.xml":
        return os.path.join(ISO, "shared-documentPropertiesExtended.xsd")
    if part.startswith("xl/charts/"):
        return os.path.join(ISO, "dml-chart.xsd")
    if part.startswith("xl/drawings/") and part.endswith(".xml"):
        return os.path.join(ISO, "dml-spreadsheetDrawing.xsd")
    if part.startswith("xl/theme/"):
        return os.path.join(ISO, "dml-main.xsd")
    if part.startswith("xl/") and part.endswith(".xml"):
        return os.path.join(ISO, "sml.xsd")
    return None  # docProps/core.xml (Dublin Core imports), media


def is_known_false_positive(err):
    return "{http://www.w3.org/XML/1998/namespace}space" in err.message and "}t'" in err.message


errors = 0
parts = 0
for f in sys.argv[1:]:
    if not os.path.isfile(f):
        fail_setup(f"no such file: {f}")
    with zipfile.ZipFile(f) as z:
        for name in z.namelist():
            xsd = None if name.endswith("/") else schema_for(name)
            if xsd is None:
                continue
            parts += 1
            sc = schema(xsd)
            try:
                doc = etree.fromstring(z.read(name))
            except etree.XMLSyntaxError as e:
                errors += 1
                print(f"{f}:{name}:{e.lineno}: not well-formed: {e.msg}")
                continue
            if sc.validate(doc):
                continue
            for e in sc.error_log:
                if is_known_false_positive(e):
                    continue
                errors += 1
                print(f"{f}:{name}:{e.line}: {e.message[:300]}")

print(f"validate-ooxml: {parts} parts in {len(sys.argv) - 1} files validated, {errors} errors")
sys.exit(1 if errors else 0)
