"""Remove drawings / pictures / shapes from an .xlsx so the app can load it quickly and safely."""
import re
import zipfile


def strip_drawings(src, dst):
    zin = zipfile.ZipFile(src)
    drop = [n for n in zin.namelist()
            if n.startswith(("xl/drawings/", "xl/media/", "xl/embeddings/", "xl/activeX/"))
            or re.match(r"xl/(comments\d*\.xml|threadedComments/.*)", n)
            or re.match(r"xl/worksheets/_rels/.*\.rels", n) and False]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            name = item.filename
            if name in drop:
                continue
            data = zin.read(name)
            if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                t = data.decode("utf-8")
                t = re.sub(r"<drawing\b[^>]*/>", "", t)
                t = re.sub(r"<legacyDrawing\b[^>]*/>", "", t)
                t = re.sub(r"<legacyDrawingHF\b[^>]*/>", "", t)
                t = re.sub(r"<picture\b[^>]*/>", "", t)
                data = t.encode("utf-8")
            elif re.match(r"xl/worksheets/_rels/.*\.rels$", name):
                t = data.decode("utf-8")
                t = re.sub(r"<Relationship\b[^>]*Type=\"[^\"]*/(drawing|vmlDrawing|image|comments)\"[^>]*/>", "", t)
                data = t.encode("utf-8")
            elif name == "[Content_Types].xml":
                t = data.decode("utf-8")
                t = re.sub(r"<Override\b[^>]*PartName=\"/xl/(drawings|media|embeddings|comments)[^\"]*\"[^>]*/>", "", t)
                data = t.encode("utf-8")
            zout.writestr(item, data)
