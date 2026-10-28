"""Replace the deck background with the reference academic style (white, grey title band,
gold bar under the title, grey bar at the bottom, BAU logo bottom-left). Slide content is untouched.

usage: python3 -I rebackground.py <unpacked_deck_dir> <config.json>
config keys: gold_y (in), logo_slide, covers [slide numbers], to_content [slides moved to the CONTENT layout],
             nudges [[slide, text the shape starts with, old_y, new_y, new_h or null], ...]
"""
import json
import re
import shutil
import sys
from pathlib import Path

D = Path(sys.argv[1])
CFG = json.loads(Path(sys.argv[2]).read_text())
GOLD_Y = float(CFG.get("gold_y", 1.22))
EMU = 914400
SW, SH = 12192000, 6858000  # 13.333 x 7.5 in


def e(v):
    return int(round(v * EMU))


_id = [9000]


def nid():
    _id[0] += 1
    return _id[0]


def rect(name, x, y, w, h, fill, line="<a:ln><a:noFill/></a:ln>"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'{fill}{line}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>')


def grad(c1, c2, alpha2=None):
    a = f'<a:alpha val="{alpha2}"/>' if alpha2 is not None else ""
    return (f'<a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="{c1}"/></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="{c2}">{a}</a:srgbClr></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>')


def logo(rid):
    h = 0.42
    w = h * 250 / 214
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{nid()}" name="BAU logo"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr><a:xfrm><a:off x="{e(0.12)}" y="{e(CFG.get("logo_y", 7.03))}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')


GOLD = lambda y: rect("Gold bar", 0, e(y), SW, e(0.07), grad("766000", "FFDE53"), '<a:ln w="3240"><a:solidFill><a:srgbClr val="E8BC00"/></a:solidFill></a:ln>')
BOTTOM_AT = lambda y: rect("Bottom bar", 0, e(y), SW, e(0.07), grad("333333", "D0D0D0"), '<a:ln w="3240"><a:solidFill><a:srgbClr val="ABABAB"/></a:solidFill></a:ln>')
BOTTOM_Y = float(CFG.get("bottom_y", 6.93))
BOTTOM = BOTTOM_AT(BOTTOM_Y)
TITLE_BAND = rect("Title band", 0, 0, SW, e(GOLD_Y), grad("D0D0D0", "FFFFFF", 0))
WHITE_BG = '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
LOGO_RID = "rIdBauLogo"


def add_rel(rels_path, rid, typ, target):
    t = rels_path.read_text()
    if f'Id="{rid}"' not in t:
        t = t.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{typ}" Target="{target}"/></Relationships>')
        rels_path.write_text(t)


IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"


def insert_shapes(xml, shapes):
    # right after the group properties, i.e. behind every other shape
    i = xml.index("</p:grpSpPr>") + len("</p:grpSpPr>")
    return xml[:i] + shapes + xml[i:]


def set_bg(xml, bg):
    xml = re.sub(r"<p:bg>.*?</p:bg>", "", xml, flags=re.S)
    if bg:
        xml = re.sub(r"(<p:cSld[^>]*>)", r"\1" + bg, xml, count=1)
    return xml


# --- logo image ---
media = D / "ppt" / "media"
logo_src = None
for rel in (D / "ppt/slides/_rels").glob(f"slide{CFG.get('logo_slide', 2)}.xml.rels"):
    m = re.search(r'Target="\.\./media/([^"]+)"', rel.read_text())
    logo_src = m.group(1) if m else None
assert logo_src, "BAU logo not found on the logo slide"
shutil.copyfile(media / logo_src, media / "bau_logo.jpeg")

# --- master: white background + content-slide decorations ---
mp = D / "ppt/slideMasters/slideMaster1.xml"
mx = set_bg(mp.read_text(), WHITE_BG)
mx = insert_shapes(mx, TITLE_BAND + GOLD(GOLD_Y) + BOTTOM + logo(LOGO_RID))
mrels = D / "ppt/slideMasters/_rels/slideMaster1.xml.rels"
add_rel(mrels, LOGO_RID, IMG, "../media/bau_logo.jpeg")

# --- layouts: drop their own backgrounds so the master shows through ---
for lp in sorted((D / "ppt/slideLayouts").glob("slideLayout*.xml")):
    lp.write_text(set_bg(lp.read_text(), None))

# --- title-style layout (gold bar across the middle), for title / section / closing slides ---
def make_title_style(lp, with_logo, bottom_y=None):
    bottom_y = BOTTOM_Y if bottom_y is None else bottom_y
    x = lp.read_text()
    x = set_bg(x, WHITE_BG)
    if 'showMasterSp="0"' not in x:
        x = x.replace("<p:sldLayout ", '<p:sldLayout showMasterSp="0" ', 1)
    x = insert_shapes(x, GOLD(2.22) + BOTTOM_AT(bottom_y) + (logo(LOGO_RID) if with_logo else ""))
    lp.write_text(x)
    if with_logo:
        add_rel(lp.parent / "_rels" / (lp.name + ".rels"), LOGO_RID, IMG, "../media/bau_logo.jpeg")

# L2 "TITLE" (section and thank-you slides)
make_title_style(D / "ppt/slideLayouts/slideLayout2.xml", with_logo=True)

# New layout 8 = copy of L3 (keeps its title placeholder) for the two cover slides
src, dst = D / "ppt/slideLayouts/slideLayout3.xml", D / "ppt/slideLayouts/slideLayout8.xml"
dst.write_text(re.sub(r'<p:cSld name="[^"]*"', '<p:cSld name="COVER"', src.read_text(), count=1))
shutil.copyfile(D / "ppt/slideLayouts/_rels/slideLayout3.xml.rels", D / "ppt/slideLayouts/_rels/slideLayout8.xml.rels")
make_title_style(dst, with_logo=False, bottom_y=7.38)

ct = D / "[Content_Types].xml"
t = ct.read_text()
if "/ppt/slideLayouts/slideLayout8.xml" not in t:
    t = t.replace("</Types>", '<Override PartName="/ppt/slideLayouts/slideLayout8.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/></Types>')
ct.write_text(t)
add_rel(mrels, "rIdCoverLayout", LAYOUT, "../slideLayouts/slideLayout8.xml")
ids = [int(v) for v in re.findall(r'<p:sldLayoutId id="(\d+)"', mx)] + [int(v) for v in re.findall(r'<p:sldMasterId id="(\d+)"', (D / "ppt/presentation.xml").read_text())]
mx = mx.replace("</p:sldLayoutIdLst>", f'<p:sldLayoutId id="{max(ids) + 1}" r:id="rIdCoverLayout"/></p:sldLayoutIdLst>')
mp.write_text(mx)

# cover slides use the new layout; listed slides move to the CONTENT layout
for n in CFG.get("covers", [1, 2]):
    rp = D / f"ppt/slides/_rels/slide{n}.xml.rels"
    rp.write_text(re.sub(r"\.\./slideLayouts/slideLayout\d+\.xml", "../slideLayouts/slideLayout8.xml", rp.read_text()))
for n in CFG.get("to_content", []):
    rp = D / f"ppt/slides/_rels/slide{n}.xml.rels"
    rp.write_text(re.sub(r"\.\./slideLayouts/slideLayout\d+\.xml", "../slideLayouts/slideLayout3.xml", rp.read_text()))

# slides keep their own content; drop any slide-level background override
for sp in (D / "ppt/slides").glob("slide*.xml"):
    x = sp.read_text()
    if "<p:bg>" in x:
        sp.write_text(set_bg(x, None))
# Small nudges so titles end above the gold bar and captions start below it.
# (slide number, text the shape starts with, old y, new y, new height or None) - inches
NUDGES = CFG.get("nudges", [])
for n, start, old_y, new_y, new_h, *rest in NUDGES:
    new_x = rest[0] if rest else None
    sp = D / f"ppt/slides/slide{n}.xml"
    x = sp.read_text()
    hits = 0
    def fix(m):
        global hits
        b = m.group(0)
        if not "".join(re.findall(r"<a:t>([^<]*)</a:t>", b)).startswith(start):
            return b
        o = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"\s*/>\s*<a:ext cx="(\d+)" cy="(\d+)"', b)
        if abs(int(o.group(2)) / EMU - old_y) > 0.02:
            return b
        hits += 1
        cy = e(new_h) if new_h else int(o.group(4))
        x_off = e(new_x) if new_x is not None else int(o.group(1))
        cx = int(o.group(3)) - (x_off - int(o.group(1)))  # keep the right edge when moving right
        return b.replace(o.group(0), f'<a:off x="{x_off}" y="{e(new_y)}"/><a:ext cx="{cx}" cy="{cy}"')
    x = re.sub(r"<p:(?:sp|graphicFrame)>.*?</p:(?:sp|graphicFrame)>", fix, x, flags=re.S)
    assert hits == 1, (n, start, hits)
    sp.write_text(x)
print("done; gold bar at", GOLD_Y, "in")
