"""Replace fragile PowerPoint Unicode math with typeset SVG equations."""

from pathlib import Path
from tempfile import TemporaryDirectory
import shutil
import subprocess
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt
from PIL import Image


PPTX = Path.home() / "Desktop" / "Algorithm_Framework_English.pptx"
ASSETS = Path(__file__).resolve().parents[1] / "out" / "algorithm_math"
SCALE = 0.6
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"

EQUATIONS = {
    "Full policy math": (r"a_0 \sim \pi_0(\cdot \mid o_0)", (437, 266, 131, 25)),
    "Mask math": (r"o_0^{-n}", (458, 493, 90, 26)),
    "Counterfactual action math": (r"a_0^{-n} \sim \pi_0", (638, 493, 127, 26)),
    "Critic math": (r"Q_0^{-n}", (1214, 187, 190, 25)),
    "Payment math": (
        r"p_n = Q_0^{-n,*} - Q_0^{-n}\!\left(\Pi_{-n}(a_0^*)\right)",
        (1130, 305, 360, 30),
    ),
}


def run(command, cwd):
    result = subprocess.run(command, cwd=cwd, text=True, capture_output=True)
    if result.returncode:
        raise RuntimeError(f"{' '.join(command)}\n{result.stdout[-3000:]}\n{result.stderr[-3000:]}")


def render_equation(name, formula):
    ASSETS.mkdir(parents=True, exist_ok=True)
    stem = name.lower().replace(" ", "_")
    with TemporaryDirectory(prefix="algorithm-math-") as temp:
        temp_path = Path(temp)
        tex_path = temp_path / f"{stem}.tex"
        tex_path.write_text(
            "\\documentclass[border=0pt]{standalone}\n"
            "\\usepackage{amsmath,amssymb}\n"
            "\\begin{document}\n"
            "\\fontsize{14}{16}\\selectfont$\\displaystyle "
            + formula
            + "$\n\\end{document}\n",
            encoding="utf-8",
        )
        run(["xelatex", "-interaction=nonstopmode", "-halt-on-error", tex_path.name], temp_path)
        pdf_path = temp_path / f"{stem}.pdf"
        run(["xelatex", "-no-pdf", "-interaction=nonstopmode", "-halt-on-error", tex_path.name], temp_path)
        xdv_path = temp_path / f"{stem}.xdv"
        svg_path = ASSETS / f"{stem}.svg"
        png_base = ASSETS / stem
        run(
            ["dvisvgm", "--no-fonts", "--exact-bbox", f"--output={svg_path}", str(xdv_path)],
            temp_path,
        )
        run(
            ["pdftocairo", "-png", "-transp", "-r", "300", "-singlefile", str(pdf_path), str(png_base)],
            temp_path,
        )
        return svg_path, png_base.with_suffix(".png")


def set_label(shape, text, x, y, w, h, size):
    shape.left, shape.top = Pt(x * SCALE), Pt(y * SCALE)
    shape.width, shape.height = Pt(w * SCALE), Pt(h * SCALE)
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = False
    frame.margin_left = frame.margin_right = Pt(0)
    frame.margin_top = frame.margin_bottom = Pt(0)
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = text
    run.font.name = "Arial"
    run.font.size = Pt(size * SCALE)
    run.font.bold = False
    run.font.color.rgb = RGBColor(43, 33, 23)


def add_text(slide, name, text, box, size=14):
    x, y, w, h = box
    shape = slide.shapes.add_textbox(Pt(x * SCALE), Pt(y * SCALE), Pt(w * SCALE), Pt(h * SCALE))
    shape.name = name
    set_label(shape, text, x, y, w, h, size)


def add_picture(slide, name, png_path, box):
    x, y, max_w, max_h = box
    with Image.open(png_path) as image:
        aspect = image.width / image.height
    width = min(max_w, max_h * aspect)
    height = width / aspect
    left = x + (max_w - width) / 2
    top = y + (max_h - height) / 2
    picture = slide.shapes.add_picture(
        str(png_path), Pt(left * SCALE), Pt(top * SCALE),
        width=Pt(width * SCALE), height=Pt(height * SCALE),
    )
    picture.name = name


def embed_svg_parts(pptx_path, svg_by_name):
    with TemporaryDirectory(prefix="algorithm-pptx-") as temp:
        rebuilt = Path(temp) / "rebuilt.pptx"
        with zipfile.ZipFile(pptx_path, "r") as source:
            entries = {item.filename: source.read(item.filename) for item in source.infolist()}

        slide = etree.fromstring(entries["ppt/slides/slide1.xml"])
        rels = etree.fromstring(entries["ppt/slides/_rels/slide1.xml.rels"])
        content_types = etree.fromstring(entries["[Content_Types].xml"])
        if not content_types.xpath("ct:Default[@Extension='svg']", namespaces={"ct": CT_NS}):
            etree.SubElement(content_types, f"{{{CT_NS}}}Default", Extension="svg", ContentType="image/svg+xml")

        relation_ids = [
            int(node.get("Id")[3:])
            for node in rels.findall(f"{{{REL_NS}}}Relationship")
            if node.get("Id", "").startswith("rId") and node.get("Id", "")[3:].isdigit()
        ]
        next_id = max(relation_ids, default=0) + 1
        for name, svg_path in svg_by_name.items():
            picture = slide.xpath(
                ".//p:pic[p:nvPicPr/p:cNvPr[@name=$name]]",
                namespaces={"p": P_NS},
                name=name,
            )
            if len(picture) != 1:
                raise RuntimeError(f"Expected one SVG picture for {name}, found {len(picture)}")
            blip = picture[0].find(f".//{{{A_NS}}}blip")
            if blip is None:
                raise RuntimeError(f"Missing blip for {name}")

            rid = f"rId{next_id}"
            next_id += 1
            media_name = svg_path.name
            entries[f"ppt/media/{media_name}"] = svg_path.read_bytes()
            etree.SubElement(
                rels, f"{{{REL_NS}}}Relationship",
                Id=rid,
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                Target=f"../media/{media_name}",
            )
            ext_list = etree.SubElement(blip, f"{{{A_NS}}}extLst")
            ext = etree.SubElement(ext_list, f"{{{A_NS}}}ext", uri=SVG_EXT_URI)
            svg_blip = etree.SubElement(ext, f"{{{SVG_NS}}}svgBlip", nsmap={"asvg": SVG_NS})
            svg_blip.set(f"{{{R_NS}}}embed", rid)

        entries["ppt/slides/slide1.xml"] = etree.tostring(slide, encoding="UTF-8", xml_declaration=True, standalone=True)
        entries["ppt/slides/_rels/slide1.xml.rels"] = etree.tostring(rels, encoding="UTF-8", xml_declaration=True, standalone=True)
        entries["[Content_Types].xml"] = etree.tostring(content_types, encoding="UTF-8", xml_declaration=True, standalone=True)
        with zipfile.ZipFile(rebuilt, "w", compression=zipfile.ZIP_DEFLATED) as output:
            for name, data in entries.items():
                output.writestr(name, data)
        shutil.copyfile(rebuilt, pptx_path)


def main():
    if not PPTX.is_file():
        raise FileNotFoundError(PPTX)
    rendered = {name: render_equation(name, formula) for name, (formula, _) in EQUATIONS.items()}
    presentation = Presentation(PPTX)
    slide = presentation.slides[0]
    for shape in list(slide.shapes):
        if shape.name in EQUATIONS or shape.name == "Critic description":
            shape._element.getparent().remove(shape._element)

    by_name = {shape.name: shape for shape in slide.shapes}
    set_label(by_name["Full policy label"], "UAV policy", 430, 230, 145, 25, 15)
    set_label(by_name["Mask label"], "Mask user n", 430, 459, 145, 26, 15)
    set_label(by_name["Counterfactual action label"], "Counterfactual", 627, 459, 150, 26, 15)
    set_label(by_name["Critic label"], "Counterfactual critic", 1100, 166, 420, 22, 15)
    add_text(slide, "Critic description", "Evaluate the removal environment", (1100, 211, 420, 20), 14)
    set_label(by_name["Comparison label"], "Compare long-term values", 1100, 269, 420, 25, 15)
    set_label(by_name["Update label"], "Update actors and critics\nUser and UAV policies", 1255, 487, 235, 58, 15)
    set_label(by_name["Payment and next state"], "Payment and state update\nCharge users, update state", 80, 472, 220, 58, 12.5)

    for name, (_, box) in EQUATIONS.items():
        add_picture(slide, name, rendered[name][1], box)
    presentation.save(PPTX)
    embed_svg_parts(PPTX, {name: paths[0] for name, paths in rendered.items()})
    print(f"Updated {PPTX} with {len(EQUATIONS)} embedded SVG equations")


if __name__ == "__main__":
    main()
