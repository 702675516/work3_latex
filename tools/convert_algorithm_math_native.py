"""Convert the diagram's SVG equation objects to editable Office Math."""

from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory
import argparse
import shutil
import zipfile

from lxml import etree
from latex2mathml.converter import convert as latex_to_mathml


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_PPTX = ROOT / "out" / "equation_command_test.pptx"
OFFICE_XSLT = Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL")
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"

FORMULAS = {
    "Full policy math": r"a_0 \sim \pi_0(\cdot \mid o_0)",
    "Mask math": r"o_0^{-n}",
    "Counterfactual action math": r"a_0^{-n} \sim \pi_0",
    "Critic math": r"Q_0^{-n}",
    "Payment math": r"p_n = Q_0^{-n,*} - Q_0^{-n}\!\left(\Pi_{-n}(a_0^*)\right)",
}
FONT_SIZES = {
    "Full policy math": 1000,
    "Mask math": 1000,
    "Counterfactual action math": 1000,
    "Critic math": 1000,
    "Payment math": 1100,
}
EMU_PER_LAYOUT_UNIT = 7620
PAYMENT_LAYOUT = {
    "Critic label": (1180, 174, 205, 27),
    "Critic math": (1360, 172, 52, 30),
    "Payment math": (1130, 293, 360, 30),
}


def load_template():
    with zipfile.ZipFile(TEMPLATE_PPTX) as package:
        root = etree.fromstring(package.read("ppt/slides/slide1.xml"))
    nodes = root.xpath(".//mc:AlternateContent", namespaces={"mc": MC_NS})
    if len(nodes) != 1:
        raise RuntimeError("The PowerPoint equation template is unavailable")
    return nodes[0]


def formula_omml(formula, transform):
    mathml = etree.fromstring(latex_to_mathml(formula).encode("utf-8"))
    transformed = transform(mathml)
    equation = transformed.getroot()
    if equation.tag != f"{{{M_NS}}}oMath":
        raise RuntimeError(f"Unexpected Office Math root: {equation.tag}")
    return equation


def adjust_payment_layout(slide):
    for name, (x, y, width, height) in PAYMENT_LAYOUT.items():
        shapes = slide.xpath(
            ".//p:sp[p:nvSpPr/p:cNvPr[@name=$name]] | "
            ".//p:pic[p:nvPicPr/p:cNvPr[@name=$name]]",
            namespaces={"p": P_NS},
            name=name,
        )
        if not shapes:
            raise RuntimeError(f"Missing payment layout shape: {name}")
        for shape in shapes:
            transform = shape.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm")
            if transform is None:
                raise RuntimeError(f"Missing geometry for {name}")
            offset = transform.find(f"{{{A_NS}}}off")
            extent = transform.find(f"{{{A_NS}}}ext")
            for node, values in ((offset, (x, y)), (extent, (width, height))):
                for attribute, value in zip(("x", "y") if node is offset else ("cx", "cy"), values):
                    node.set(attribute, str(value * EMU_PER_LAYOUT_UNIT))


def find_shape(slide, name, kind="sp"):
    shape = slide.xpath(
        f".//p:{kind}[p:nvSpPr/p:cNvPr[@name=$name]]" if kind == "sp" else
        f".//p:{kind}[p:nvCxnSpPr/p:cNvPr[@name=$name]]",
        namespaces={"p": P_NS},
        name=name,
    )
    if len(shape) != 1:
        raise RuntimeError(f"Expected one {kind} named {name}, found {len(shape)}")
    return shape[0]


def set_shape_name(shape, name):
    c_nv_pr = shape.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
    if c_nv_pr is None:
        c_nv_pr = shape.find(f"{{{P_NS}}}nvCxnSpPr/{{{P_NS}}}cNvPr")
    c_nv_pr.set("name", name)


def set_shape_geometry(shape, x, y, width, height):
    xfrm = shape.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm")
    if xfrm is None:
        xfrm = shape.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm")
    offset = xfrm.find(f"{{{A_NS}}}off")
    extent = xfrm.find(f"{{{A_NS}}}ext")
    offset.set("x", str(x * EMU_PER_LAYOUT_UNIT))
    offset.set("y", str(y * EMU_PER_LAYOUT_UNIT))
    extent.set("cx", str(width * EMU_PER_LAYOUT_UNIT))
    extent.set("cy", str(height * EMU_PER_LAYOUT_UNIT))


def set_text_shape(shape, text, x, y, width, height):
    set_shape_geometry(shape, x, y, width, height)
    tx_body = shape.find(f"{{{P_NS}}}txBody")
    old_paragraph = tx_body.find(f"{{{A_NS}}}p")
    p_pr = old_paragraph.find(f"{{{A_NS}}}pPr")
    old_run = old_paragraph.find(f"{{{A_NS}}}r")
    run_pr = old_run.find(f"{{{A_NS}}}rPr") if old_run is not None else None
    end_pr = old_paragraph.find(f"{{{A_NS}}}endParaRPr")
    for paragraph in tx_body.findall(f"{{{A_NS}}}p"):
        tx_body.remove(paragraph)
    for line in text.split("\n"):
        paragraph = etree.SubElement(tx_body, f"{{{A_NS}}}p")
        if p_pr is not None:
            paragraph.append(deepcopy(p_pr))
        run = etree.SubElement(paragraph, f"{{{A_NS}}}r")
        if run_pr is not None:
            run.append(deepcopy(run_pr))
        etree.SubElement(run, f"{{{A_NS}}}t").text = line
        if end_pr is not None:
            paragraph.append(deepcopy(end_pr))


def assign_shape_id(slide, shape):
    ids = [int(value) for value in slide.xpath(".//p:cNvPr/@id", namespaces={"p": P_NS})]
    c_nv_pr = shape.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
    if c_nv_pr is None:
        c_nv_pr = shape.find(f"{{{P_NS}}}nvCxnSpPr/{{{P_NS}}}cNvPr")
    c_nv_pr.set("id", str(max(ids, default=0) + 1))


def update_diagram_labels(slide):
    labels = {
        "Learning heading": "4. Per-agent SAC learning",
        "Full buffer note": "Store transitions in per-agent replay buffers",
        "Counterfactual buffer note": "Store counterfactual transition in the UAV replay buffer",
        "Critic label": "UAV SAC critic",
    }
    for name, text in labels.items():
        shape = find_shape(slide, name)
        paragraphs = shape.find(f"{{{P_NS}}}txBody").findall(f"{{{A_NS}}}p")
        runs = [run for paragraph in paragraphs for run in paragraph.findall(f"{{{A_NS}}}r")]
        if not runs:
            raise RuntimeError(f"Missing text in shape {name}")
        runs[0].find(f"{{{A_NS}}}t").text = text
        for run in runs[1:]:
            run.getparent().remove(run)

    branch_label = find_shape(slide, "Branch sampling")
    set_text_shape(
        branch_label,
        "Select training branch κ",
        80, 238, 220, 58,
    )

    layout = {
        "Replay buffers": ("UAV replay buffer", 1090, 474, 125, 62),
        "SAC update": ("UAV SAC update", 1247, 474, 283, 62),
        "Replay label": ("UAV buffer label", "UAV replay\nbuffer", 1095, 480, 115, 50),
        "Update label": ("UAV update label", "Update UAV actor\nand critics", 1252, 480, 273, 50),
    }
    clones = {}
    for old_name, values in layout.items():
        shape = find_shape(slide, old_name)
        clone = deepcopy(shape)
        if old_name in ("Replay buffers", "SAC update"):
            new_name, x, y, width, height = values
            set_shape_name(shape, new_name)
            set_shape_geometry(shape, x, y, width, height)
            set_shape_name(clone, "User replay buffer" if old_name == "Replay buffers" else "User SAC update")
            set_shape_geometry(clone, x, 550, width, height)
        else:
            new_name, text, x, y, width, height = values
            set_shape_name(shape, new_name)
            set_text_shape(shape, text, x, y, width, height)
            set_shape_name(clone, "User buffer label" if old_name == "Replay label" else "User update label")
            user_text = "User n replay\nbuffer" if old_name == "Replay label" else "Update user n actor and critics\nfor each n ∈ N"
            set_text_shape(clone, user_text, x, 556, width, height)
        assign_shape_id(slide, clone)
        shape.addnext(clone)
        clones[old_name] = clone

    arrow = find_shape(slide, "Replay to update_0", "cxnSp")
    set_shape_name(arrow, "UAV replay to update")
    arrow_xfrm = arrow.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm")
    arrow_xfrm.find(f"{{{A_NS}}}off").set("y", str(505 * EMU_PER_LAYOUT_UNIT))
    user_arrow = deepcopy(arrow)
    set_shape_name(user_arrow, "User replay to update")
    user_arrow.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm/{{{A_NS}}}off").set("y", str(581 * EMU_PER_LAYOUT_UNIT))
    assign_shape_id(slide, user_arrow)
    arrow.addnext(user_arrow)

    for shape in slide.xpath(".//p:sp | .//p:cxnSp", namespaces={"p": P_NS}):
        names = shape.xpath("./p:nvSpPr/p:cNvPr/@name | ./p:nvCxnSpPr/p:cNvPr/@name", namespaces={"p": P_NS})
        if names and (names[0] == "Feedback note" or names[0].startswith("Learning feedback loop")):
            shape.getparent().remove(shape)


def convert_slide(package_entries):
    slide = etree.fromstring(package_entries["ppt/slides/slide1.xml"])
    template = load_template()
    transform = etree.XSLT(etree.parse(str(OFFICE_XSLT)))

    for name, formula in FORMULAS.items():
        font_size = str(FONT_SIZES[name])
        pictures = slide.xpath(
            ".//p:pic[p:nvPicPr/p:cNvPr[@name=$name]]",
            namespaces={"p": P_NS},
            name=name,
        )
        if len(pictures) != 1:
            raise RuntimeError(f"Expected one picture named {name}, found {len(pictures)}")
        picture = pictures[0]
        picture_id = picture.find(f"{{{P_NS}}}nvPicPr/{{{P_NS}}}cNvPr").get("id")
        xfrm = picture.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm")
        if xfrm is None:
            raise RuntimeError(f"Missing geometry for {name}")

        native = deepcopy(template)
        choice_shape = native.find(f"{{{MC_NS}}}Choice/{{{P_NS}}}sp")
        if choice_shape is None:
            raise RuntimeError("Missing native equation shape")
        c_nv_pr = choice_shape.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
        c_nv_pr.set("id", picture_id)
        c_nv_pr.set("name", name)
        for ext_list in c_nv_pr.findall(f"{{{A_NS}}}extLst"):
            c_nv_pr.remove(ext_list)

        old_xfrm = choice_shape.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm")
        old_xfrm.getparent().replace(old_xfrm, deepcopy(xfrm))
        math_wrapper = choice_shape.find(f".//{{{A14_NS}}}m")
        math_para = math_wrapper.find(f"{{{M_NS}}}oMathPara")
        old_equation = math_para.find(f"{{{M_NS}}}oMath")
        math_para.replace(old_equation, formula_omml(formula, transform))

        paragraph = choice_shape.find(f"{{{P_NS}}}txBody/{{{A_NS}}}p")
        end = paragraph.find(f"{{{A_NS}}}endParaRPr")
        if end is not None:
            end.set("sz", font_size)
        paragraph_properties = etree.Element(f"{{{A_NS}}}pPr", algn="ctr")
        paragraph.insert(0, paragraph_properties)
        for run in math_para.findall(f".//{{{M_NS}}}r"):
            run_properties = etree.Element(f"{{{A_NS}}}rPr", lang="en-US", sz=font_size)
            etree.SubElement(run_properties, f"{{{A_NS}}}latin", typeface="Cambria Math")
            run.insert(0, run_properties)

        fallback = native.find(f"{{{MC_NS}}}Fallback")
        fallback.clear()
        fallback_picture = deepcopy(picture)
        blip = fallback_picture.find(f".//{{{A_NS}}}blip")
        if blip is not None:
            for extension_list in blip.findall(f"{{{A_NS}}}extLst"):
                blip.remove(extension_list)
        fallback.append(fallback_picture)
        picture.getparent().replace(picture, native)

    adjust_payment_layout(slide)
    update_diagram_labels(slide)
    package_entries["ppt/slides/slide1.xml"] = etree.tostring(
        slide, encoding="UTF-8", xml_declaration=True, standalone=True
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=Path.home() / "Desktop" / "Algorithm_Framework_English.pptx")
    parser.add_argument("--output", type=Path, default=ROOT / "out" / "Algorithm_Framework_English_native_test.pptx")
    args = parser.parse_args()
    with zipfile.ZipFile(args.input, "r") as package:
        entries = {item.filename: package.read(item.filename) for item in package.infolist()}
    convert_slide(entries)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="algorithm-native-equations-") as temp:
        target = Path(temp) / "native.pptx"
        with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as package:
            for name, data in entries.items():
                package.writestr(name, data)
        shutil.copyfile(target, args.output)
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
