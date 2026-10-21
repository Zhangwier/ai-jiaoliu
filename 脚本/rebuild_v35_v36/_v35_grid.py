from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
prs=Presentation(); prs.slide_width=Inches(13.3333); prs.slide_height=Inches(7.5)
s=prs.slides.add_slide(prs.slide_layouts[6])
# markers at 1in grid
for i in range(14):
    r=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(i),Inches(0),Pt(2),Inches(0.3))
    r.fill.solid(); r.fill.fore_color.rgb=RGBColor(0,0,0); r.line.fill.background()
for j in range(8):
    r=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(0),Inches(j),Inches(0.3),Pt(2))
    r.fill.solid(); r.fill.fore_color.rgb=RGBColor(200,0,0); r.line.fill.background()
# box exactly 0..11in x 2..3in
b=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(1),Inches(2),Inches(11),Inches(1))
b.fill.background(); b.line.color.rgb=RGBColor(0,120,0); b.line.width=Pt(2)
prs.save(r'_v35\grid.pptx'); print('ok')
