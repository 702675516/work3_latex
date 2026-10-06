from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.util import Pt


SCALE = 0.6
INK = RGBColor(43, 33, 23)
FLOW = RGBColor(51, 65, 85)
BLUE = RGBColor(37, 99, 166)
NODE = RGBColor(247, 249, 251)
NODE_LINE = RGBColor(100, 116, 139)
BLUE_FILL = RGBColor(234, 242, 250)
BLUE_LINE = RGBColor(57, 115, 168)
GOLD_FILL = RGBColor(251, 243, 229)
GOLD_LINE = RGBColor(173, 128, 52)


def pt(value):
    return Pt(value * SCALE)


def set_label(shape, text, size=12.5, bold=False, color=INK):
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.margin_left = Pt(1.5)
    frame.margin_right = Pt(1.5)
    frame.margin_top = Pt(0)
    frame.margin_bottom = Pt(0)
    for i, line in enumerate(text.split("\n")):
        paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        paragraph.alignment = PP_ALIGN.CENTER
        run = paragraph.add_run()
        run.text = line
        run.font.name = "Arial"
        run.font.size = pt(size)
        run.font.bold = bold
        run.font.color.rgb = color


def add_box(slide, name, x, y, w, h, text, fill, stroke, size=12.5):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, pt(x), pt(y), pt(w), pt(h)
    )
    shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = stroke
    shape.line.width = pt(1.2)
    try:
        shape.adjustments[0] = 0.06
    except (IndexError, TypeError):
        pass
    set_label(shape, text, size=size)
    return shape


def add_text(slide, name, text, x, y, w, h, size=13, bold=False, color=INK):
    shape = slide.shapes.add_textbox(pt(x), pt(y), pt(w), pt(h))
    shape.name = name
    shape.fill.background()
    shape.line.fill.background()
    set_label(shape, text, size=size, bold=bold, color=color)
    return shape


def add_path(slide, name, points, color=FLOW, dashed=False, weight=1.8):
    for index, ((x1, y1), (x2, y2)) in enumerate(zip(points, points[1:])):
        line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, pt(x1), pt(y1), pt(x2), pt(y2)
        )
        line.name = f"{name}_{index}"
        line.line.color.rgb = color
        line.line.width = pt(weight)
        if dashed:
            line.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        if index == len(points) - 2:
            line._element.spPr.ln.append(
                parse_xml(f'<a:tailEnd {nsdecls("a")} type="triangle" w="med" len="med"/>')
            )


def main():
    desktop = Path.home() / "Desktop"
    path = desktop / "Algorithm_Framework_English.pptx"
    if not path.is_file():
        raise FileNotFoundError(path)
    presentation = Presentation(path)
    slide = presentation.slides[0]

    remove_names = {
        "State to report_0",
        "Report to branch_0",
        "Branch to full rollout_0",
        "Branch to counterfactual rollout_0",
        "System state",
        "User reports",
        "Branch sampling",
        "State label",
        "Report label",
        "Branch label",
        "Branch outcome note",
        "Branch full label",
        "Branch counterfactual label",
        "Payment transfer label",
        "Online UAV decision",
        "Online MEC execution",
        "Online price settlement",
        "Next state",
        "Observe and report",
        "Payment and next state",
    }
    remove_prefixes = (
        "State to report_",
        "Observe to branch sample_",
        "Report to branch sample_",
        "Branch sample to UAV decision_",
        "UAV decision to execution_",
        "Execution to settlement_",
        "Settlement to next state_",
        "State transition feedback_",
        "Branch to full rollout_",
        "Branch to counterfactual rollout_",
        "Payment to online settlement_",
    )
    for shape in list(slide.shapes):
        if shape.name in remove_names or shape.name.startswith(remove_prefixes):
            shape._element.getparent().remove(shape._element)

    for shape in slide.shapes:
        if shape.name == "Online heading":
            shape.left, shape.top = pt(64), pt(102)
            shape.width, shape.height = pt(285), pt(38)
            set_label(shape, "1. Online interaction", size=18, bold=True)
        elif shape.name == "Rollout heading":
            shape.left, shape.top = pt(399), pt(102)
            shape.width, shape.height = pt(610), pt(38)
            set_label(shape, "2. Branch-specific training rollouts", size=18, bold=True)
        elif shape.name == "Full policy label":
            set_label(shape, "UAV policy\na_0 ~ π_0(. | o_0)", size=15)
        elif shape.name == "Mask label":
            set_label(shape, "Mask user n\no_0^(-n)", size=15)
        elif shape.name == "Counterfactual action label":
            set_label(shape, "Counterfactual\na_0^(-n) ~ π_0", size=15)
        elif shape.name == "Critic label":
            set_label(shape, "Counterfactual critic Q_0^(-n)\nEvaluate the removal environment", size=16)
        elif shape.name == "Comparison label":
            set_label(
                shape,
                "Compare long-term values\n"
                "p_n = Q_0^(-n,*)\n"
                "- Q_0^(-n)(Π_(-n)(a_0^*))",
                size=14,
            )
        elif shape.name == "Update label":
            set_label(
                shape,
                "Update actors and critics\nUser policies, UAV policy, Q_0^(-n)",
                size=15,
            )

    for name, y, text, fill, stroke in [
        ("Observe and report", 160, "Observe state and report types\ns(t), type reports", NODE, NODE_LINE),
        ("Branch sampling", 238, "Sample training branch κ\nκ=0 (1−ε), κ=n (ε/N)", GOLD_FILL, GOLD_LINE),
        ("Online UAV decision", 316, "Select UAV action\nTask admission and location", BLUE_FILL, BLUE_LINE),
        ("Online MEC execution", 394, "Execute MEC service\nOffloading and UAV flight", NODE, NODE_LINE),
        ("Payment and next state", 472, "Payment and state update\nCharge users, update state", GOLD_FILL, GOLD_LINE),
    ]:
        add_box(slide, name, 80, y, 220, 58, text, fill, stroke, size=12.5)

    for name, p1, p2 in [
        ("Observe to branch sample", (190, 218), (190, 238)),
        ("Branch sample to UAV decision", (190, 296), (190, 316)),
        ("UAV decision to execution", (190, 374), (190, 394)),
        ("Execution to settlement", (190, 452), (190, 472)),
    ]:
        add_path(slide, name, [p1, p2])

    add_path(
        slide,
        "Branch to full rollout",
        [(300, 250), (350, 250), (350, 261), (420, 261)],
    )
    add_path(
        slide,
        "Branch to counterfactual rollout",
        [(300, 276), (366, 276), (366, 490), (420, 490)],
    )

    presentation.save(path)
    print(f"Updated {path}")
    print(f"Slide shapes: {len(slide.shapes)}")


if __name__ == "__main__":
    main()
