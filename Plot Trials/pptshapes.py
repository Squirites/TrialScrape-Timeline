from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.shapes.shapetree import GroupShapes

def makeYearList(start_year, end_year):
    year_list = []
    year_range =  range(start_year, end_year +1)
    for i in range(len(year_range)):
        if i == 0:
            year = "<" + str(year_range[i])
            year_list.append(year)
        elif i == len(year_range) - 1:
            year = str(year_range[i]) + "+"
            year_list.append(year)
        else:
            year_list.append(str(year_range[i]))
    return year_list

def makeTrialShape(shapes,trials,L, t, w, i):
    top = t
    gs = []
    shape1 = shapes.add_shape(MSO_SHAPE.RECTANGLE, L, t, w, Inches(0.25))
    shape1.fill.solid()
    shape_text = shape1.text_frame.paragraphs[0]
    shape_text_font = shape_text.add_run()
    shape_text_font = shape_text.font
    if trials["Phase"][i] == "['PHASE4']":
        shape1.fill.fore_color.rgb = RGBColor(220, 220, 220)
    if trials["Phase"][i] == "['PHASE3']":
        shape1.fill.fore_color.rgb = RGBColor(4, 4, 100)
    elif trials["Phase"][i] == "['PHASE2', 'PHASE3']":
        shape1.fill.fore_color.rgb = RGBColor(7, 7, 177)
    elif trials["Phase"][i] == "['PHASE2']":
        shape1.fill.fore_color.rgb = RGBColor(7, 161, 36)
    elif trials["Phase"][i] == "['PHASE1', 'PHASE2']":
        shape1.fill.fore_color.rgb = RGBColor(10, 216, 49)
    elif trials["Phase"][i] == "['PHASE1']":
        shape1.fill.fore_color.rgb = RGBColor(197, 201, 25)
    elif trials["Phase"][i] == "Observational":
        shape1.fill.fore_color.rgb = RGBColor(255, 210, 164)
    if trials["Status"][i] == "ACTIVE_NOT_RECRUITING":
        shape2 = shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, (L + w + Inches(0.1)), t, Inches(0.25), Inches(0.25))
        shape2.fill.solid()
        shape2.fill.fore_color.rgb = RGBColor(68, 114, 196)
        shape2.fill.fore_color.brightness = 0
    elif trials["Status"][i] == "COMPLETED":
        shape3 = shapes.add_shape(MSO_SHAPE.STAR_4_POINT, (L + w + Inches(0.1)), t, Inches(0.25), Inches(0.25))
        shape3.fill.solid()
        shape3.fill.fore_color.rgb = RGBColor(64, 200, 67)
        shape3.fill.fore_color.brightness = 0
    elif trials["Status"][i] == "TERMINATED":
        shape4 = shapes.add_shape(MSO_SHAPE.EXPLOSION1, (L + w + Inches(0.1)), t, Inches(0.25), Inches(0.25))
        shape4.fill.solid()
        shape4.fill.fore_color.rgb = RGBColor(246, 18, 18)
        shape4.fill.fore_color.brightness = 0
    elif trials["Status"][i] == "RECRUITING" or trials["Status"][i] == "ENROLLING_BY_INVITATION":
        shape5 = shapes.add_shape(MSO_SHAPE.DIAMOND, (L + w + Inches(0.1)), t, Inches(0.25), Inches(0.25))
        shape5.fill.solid()
        shape5.fill.fore_color.rgb = RGBColor(223, 238, 26)
        shape5.fill.fore_color.brightness = 0
    elif trials["Status"][i] == "NOT_YET_RECRUITING":
        shape6 = shapes.add_shape(MSO_SHAPE.DONUT, (L + w + Inches(0.1)), t, Inches(0.25), Inches(0.25))
        shape6.fill.solid()
        shape6.fill.fore_color.rgb = RGBColor(15, 106, 133)
        shape6.fill.fore_color.brightness = 0
    elif trials["Status"][i] == "SUSPENDED" or trials["Status"][i] == "UNKNOWN" or trials["Status"][i] == "WITHDRAWN":
        shape7 = shapes.add_shape(MSO_SHAPE.DOWN_ARROW, (L + w + Inches(0.1)), t, Inches(0.25), Inches(0.25))
        shape7.fill.solid()
        shape7.fill.fore_color.rgb = RGBColor(127, 127, 127)
        shape7.fill.fore_color.brightness = 0
    shape_text_font.color.rgb = RGBColor(55, 66, 74)
    shape1.fill.fore_color.brightness = 0

    gs.append(shape1)
    if len(gs) > 1:
        group_s = shapes.add_group_shape()
        for i in gs:
            group_s.shapes._spTree.append(i._element)

        sh1_bottom = shape1.top + shape1.height

        if check_overlap(shape1, shape2):
            shape2.top = sh1_bottom + Inches(0.03125)
            top = shape2.top
    else:
        top = t
    return top
# make header shape
def makeYearShape(prs,L2, t2, w2, h2, year, i):
    slide = prs.slides[i]
    shapes = slide.shapes
    shape = shapes.add_shape(MSO_SHAPE.RECTANGLE, L2, t2, w2, h2)
    shape.fill.background()
    shape.line.fill.background()
    shape.text = year
    shape_text = shape.text_frame.paragraphs[0]
    shape_text_font = shape_text.add_run()
    shape_text_font = shape_text.font
    shape_text_font.color.rgb = RGBColor(0, 135, 124)
    shape_text_font.bold = True
    shape.text_frame.paragraphs[0].font.size = Pt(10)

# function to make shape for trials
def makeHeaders(prs,Lt, t2, wt, ht,i):
    slide = prs.slides[i]
    shapes = slide.shapes

    comp = shapes.add_textbox(Lt, t2, wt, ht)
    comptf = comp.text_frame
    f = comptf.paragraphs[0]
    run = f.add_run()
    c = "Sponsor"
    run.text = c
    font = run.font
    font.name = "Calibri"
    font.size = Pt(14)
    font.color.rgb = RGBColor(31, 73, 125)
    font.bold = True

    Lt = Lt + Inches(1)

    reg = shapes.add_textbox(Lt, t2, wt, ht)
    regtf = reg.text_frame
    f = regtf.paragraphs[0]
    run = f.add_run()
    r = "Regimen"
    run.text = r
    font = run.font
    font.name = "Calibri"
    font.size = Pt(14)
    font.color.rgb = RGBColor(31, 73, 125)
    font.bold = True

    setting = shapes.add_textbox(Inches(2.49), t2, wt, ht)
    settingtf = setting.text_frame
    f = settingtf.paragraphs[0]
    run = f.add_run()
    r = "Setting"
    run.text = r
    font = run.font
    font.name = "Calibri"
    font.size = Pt(14)
    font.color.rgb = RGBColor(31, 73, 125)
    font.bold = True

    ident = shapes.add_textbox(Inches(3.5), t2, wt, ht)
    identtf = ident.text_frame
    f = identtf.paragraphs[0]
    run = f.add_run()
    r = "Identifier"
    run.text = r
    font = run.font
    font.name = "Calibri"
    font.size = Pt(14)
    font.color.rgb = RGBColor(31, 73, 125)
    font.bold = True

# function to add line
def makeLine(prs,begin_x, end_x, end_y, year,i):
    s = []
    slide = prs.slides[i]
    shapes = slide.shapes
    line = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, begin_x, Inches(1.15), end_x, end_y)
    s.append(line)
    if year % 2 != 0:
        line.line.color.rgb = RGBColor(192, 192, 192)
    else:
        line.line.color.rgb = RGBColor(4, 4, 88)
    line.line.dash_style = MSO_LINE.ROUND_DOT
    line.line.fill.fore_color.brightness = 0.6
    group_s = shapes.add_group_shape()
    for j in s:
        group_s.shapes._spTree.append(j._element)
    cursor_sp = shapes[0]._element
    cursor_sp.addprevious(group_s._element)

# function to make background shape for text
def makeBackShape(prs,end_y, i):
    s = []
    slide = prs.slides[i]
    shapes = slide.shapes
    shape = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.25), Inches(1.15), Inches(4.19), end_y - Inches(1.3))
    s.append(shape)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(242, 242, 242)
    shape.fill.fore_color.brightness = 0
    group_b = shapes.add_group_shape()
    for j in s:
        group_b.shapes._spTree.append(j._element)
    cursor_sp = shapes[1]._element
    cursor_sp.addprevious(group_b._element)

# function to add text
def makeText(shapes, trials,Lt, t, wt, ht,i):
    comp = shapes.add_textbox(Lt, t, wt, ht)
    comptf = comp.text_frame
    f = comptf.paragraphs[0]
    run = f.add_run()
    c = trials["Sponsor"][i]
    run.text = c
    font = run.font
    font.name = "Calibri"
    font.size = Pt(8)
    font.bold = False

    reg = shapes.add_textbox(Inches(1.14), t, wt, ht)
    regtf = reg.text_frame
    f = regtf.paragraphs[0]
    run = f.add_run()
    r = trials["Therapy"][i]
    run.text = r
    font = run.font
    font.name = "Calibri"
    font.size = Pt(8)
    font.bold = False

    m = regtf.add_paragraph()
    m.text = trials["Mechanism of Action"][i]
    m.font.size = Pt(7)

    Lt = Lt + Inches(1.125)

    MoA = shapes.add_textbox(Inches(2.4), t, wt, ht)
    MoAtf = MoA.text_frame
    f = MoAtf.paragraphs[0]
    run = f.add_run()
    r = trials["Setting"][i]
    run.text = r
    font = run.font
    font.name = "Calibri"
    font.size = Pt(8)
    font.bold = False

    p = MoAtf.add_paragraph()
    p.text = trials["Indication"][i]
    p.font.size = Pt(7)

    Identifier = shapes.add_textbox(Inches(3.52), t, wt, ht)
    Identifier = Identifier.text_frame
    f = Identifier.paragraphs[0]
    run = f.add_run()
    if trials["Name"][i] != "":
        r = trials["Name"][i]
    else:
        r = trials["Registry Code"][i]
    run.text = r
    font = run.font
    font.name = "Calibri"
    font.size = Pt(8)
    font.bold = False
    hlink = run.hyperlink
    hlink.address = trials["Link"][i]

    s = Identifier.add_paragraph()
    s.text = str(trials["Enrollment"][i])
    s.font.size = Pt(7)

# function to add an outline to highlight a key trial
def keyTrialShape(shapes,L, t, w):
    shape1 = shapes.add_shape(MSO_SHAPE.RECTANGLE, L, t, w, Inches(0.33))
    shape1.fill.background()

def lowestShape(prs,i):
    x = []
    slide = prs.slides[i]
    shapes = slide.shapes
    for shape in shapes:
        x.append(shape.top / 914400)
    return (max(x))
