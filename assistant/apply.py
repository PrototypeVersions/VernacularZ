"""Apply a Vernacular date operation to an editable PowerPoint copy."""
import argparse
from datetime import date
import json
from pathlib import Path
import re
import sys
from xml.dom import minidom
from zipfile import ZIP_DEFLATED, ZipFile
from io import BytesIO


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
MONTHS = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)
MONTH_PATTERN = "|".join(
    month[:3] + "(?:" + month[3:] + ")?" if len(month) > 3 else month
    for month in MONTHS
)
DATE = re.compile(
    r"\b(?:\d{4}-\d{2}-\d{2}|(?:" + MONTH_PATTERN +
    r")\s+\d{1,2}(?:,\s*\d{4})?)\b",
    re.IGNORECASE,
)


def parse_xml(data):
    if re.search(rb"<!DOCTYPE|<!ENTITY", data, re.IGNORECASE):
        raise ValueError("The file uses an unsupported XML declaration.")
    return minidom.parseString(data)


def slides(package):
    presentation = parse_xml(package.read("ppt/presentation.xml"))
    relationships = parse_xml(package.read("ppt/_rels/presentation.xml.rels"))
    links = {
        relation.getAttribute("Id"): relation.getAttribute("Target")
        for relation in relationships.getElementsByTagNameNS(REL, "Relationship")
        if relation.getAttribute("Type").endswith("/slide")
        and relation.getAttribute("TargetMode") != "External"
    }
    slide_ids = presentation.getElementsByTagNameNS(P, "sldId")
    if not 1 <= len(slide_ids) <= 100:
        raise ValueError("Choose a deck containing 1 to 100 slides.")
    for index, slide_id in enumerate(slide_ids, 1):
        link = links.get(slide_id.getAttributeNS(R, "id"), "")
        name = link[1:] if link.startswith("/") else "ppt/" + link
        if not re.fullmatch(r"ppt/slides/[^/]+\.xml", name):
            raise ValueError("A slide points to an unsupported file location.")
        yield index, name, parse_xml(package.read(name))


def text(run):
    return "".join(
        child.data for child in run.childNodes
        if child.nodeType in (child.TEXT_NODE, child.CDATA_SECTION_NODE)
    )


def normalized(value):
    return " ".join(value.split()).lower()


def locate(package, operation):
    prefix = normalized(operation["prefix"])
    suffix = normalized(operation["suffix"])
    matches = []
    for index, name, document in slides(package):
        for paragraph in document.getElementsByTagNameNS(A, "p"):
            runs = list(paragraph.getElementsByTagNameNS(A, "t"))
            content = "".join(map(text, runs))
            dates = list(DATE.finditer(content))
            if len(dates) != 1:
                continue
            match = dates[0]
            if (
                normalized(content[:match.start()]) == prefix
                and normalized(content[match.end():]) == suffix
            ):
                matches.append((index, name, document, runs, match))
    if len(matches) != 1:
        raise ValueError(f"The target field must occur once. Found {len(matches)} matches.")
    return matches[0]


def render_date(new_date, previous):
    if re.match(r"^\d{4}-", previous):
        return new_date.isoformat()
    month = MONTHS[new_date.month - 1]
    if len(previous.split()[0]) <= 3:
        month = month[:3]
    year = f", {new_date.year}" if re.search(r"\d{4}$", previous) else ""
    return f"{month} {new_date.day}{year}"


def edit_run(runs, match, replacement):
    offset = 0
    for run in runs:
        content = text(run)
        if offset <= match.start() and match.end() <= offset + len(content):
            start = match.start() - offset
            end = match.end() - offset
            while run.firstChild:
                run.removeChild(run.firstChild)
            run.appendChild(run.ownerDocument.createTextNode(content[:start] + replacement + content[end:]))
            return
        offset += len(content)
    raise ValueError("The target date is split across formatting runs.")


def apply(source, operation, new_date):
    if operation.get("operation") != "replace-powerpoint-date" or operation.get("version") != 1:
        raise ValueError("This is not a supported Vernacular operation.")
    if source.suffix.lower() != ".pptx" or source.stat().st_size > 15_000_000:
        raise ValueError("Choose an editable .pptx under 15 MB.")
    with ZipFile(source) as original:
        files = [entry for entry in original.infolist() if not entry.is_dir()]
        if len(files) > 3000 or sum(file.file_size for file in files) > 70_000_000:
            raise ValueError("The uncompressed PowerPoint file is too large.")
        names = [file.filename for file in files]
        if len(set(names)) != len(names) or any(
            name.lower().endswith("vbaproject.bin") or name.lower().startswith("_xmlsignatures/")
            for name in names
        ):
            raise ValueError("The deck has unsupported entries, macros, or digital signatures.")
        index, name, document, runs, match = locate(original, operation)
        before = match.group()
        after = render_date(new_date, before)
        if before == after:
            raise ValueError("Choose a date that differs from the date in this file.")
        edit_run(runs, match, after)
        expected_slide = document.toxml(encoding="utf-8")

        output = BytesIO()
        with ZipFile(output, "w", compression=ZIP_DEFLATED) as edited:
            for file in files:
                part = file.filename
                edited.writestr(part, expected_slide if part == name else original.read(file))
        with ZipFile(output) as verified:
            if verified.namelist() != names or verified.read(name) != expected_slide:
                raise ValueError("The edited slide could not be verified.")
            if any(verified.read(part) != original.read(part) for part in names if part != name):
                raise ValueError("Verification found an unexpected change outside the edited slide.")
            _, _, _, _, result = locate(verified, operation)
            if result.group() != after:
                raise ValueError("The requested date could not be read back.")
    return output.getvalue(), {"slide": index, "before": before, "after": after, "unchanged_parts": len(files) - 1}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("date", type=date.fromisoformat, help="New date in YYYY-MM-DD format")
    parser.add_argument("output", type=Path, help="Path for a new .pptx copy")
    parser.add_argument("--operation", type=Path, default=Path(__file__).with_name("operation.json"))
    args = parser.parse_args()
    try:
        if args.output.suffix.lower() != ".pptx":
            raise ValueError("Use the .pptx extension for the new copy.")
        operation = json.loads(args.operation.read_text(encoding="utf-8"))
        result, evidence = apply(args.input, operation, args.date)
        with args.output.open("xb") as destination:
            destination.write(result)
    except Exception as error:
        print(f"Could not apply the operation: {error}", file=sys.stderr)
        return 1
    print(json.dumps(evidence))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
