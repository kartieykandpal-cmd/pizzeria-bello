from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path

ROOT = Path('/Users/kartikeykandpal/Documents/Codex/2026-10-22/mak')
WORK = ROOT / 'work'
OUT = ROOT / 'outputs' / 'Pizzeria_Bello_Project_Document.docx'
CREAM=(247,242,233); INK=(41,38,31); RED=(189,73,50); GREEN=(86,102,72); MUTED=(130,125,112); LINE=(223,216,202)

def font(size, serif=False, bold=False):
    names = (['Georgia Bold.ttf','Georgia.ttf'] if serif and bold else ['Georgia.ttf'] if serif else ['Arial Bold.ttf','Arial.ttf'] if bold else ['Arial.ttf'])
    for name in names:
        p=Path('/System/Library/Fonts/Supplemental')/name
        if p.exists(): return ImageFont.truetype(str(p), size)
    return ImageFont.load_default()

def txt(draw, xy, text, size, fill=INK, serif=False, bold=False): draw.text(xy,text,font=font(size,serif,bold),fill=fill)

def pizza_art(draw, box, seed=0):
    x,y,w,h=box
    draw.rectangle((x,y,x+w,y+h),fill=(226,212,184))
    # warm restaurant-table background and a simple illustrated, baked pizza
    cx=x+w*.52; cy=y+h*.52; r=min(w,h)*.43
    draw.ellipse((cx-r*.82,cy+r*.72,cx+r*.82,cy+r*1.02),fill=(117,91,63))
    draw.ellipse((cx-r,cy-r,cx+r,cy+r),fill=(214,164,83),outline=(244,221,164),width=15)
    draw.ellipse((cx-r+15,cy-r+15,cx+r-15,cy+r-15),fill=(190,67,43))
    draw.ellipse((cx-r+28,cy-r+28,cx+r-28,cy+r-28),fill=(236,197,111))
    toppings=[(-.36,-.2),(.08,-.36),(.36,-.06),(-.16,.12),(.18,.28),(-.37,.33),(.02,.02),(.36,.34)]
    for i,(dx,dy) in enumerate(toppings):
        tx0=cx+dx*r; ty0=cy+dy*r
        if (i+seed)%3==0:
            draw.ellipse((tx0-12,ty0-7,tx0+12,ty0+7),fill=(153,50,37))
        elif (i+seed)%3==1:
            draw.ellipse((tx0-10,ty0-5,tx0+10,ty0+5),fill=(69,117,63))
        else:
            draw.ellipse((tx0-11,ty0-8,tx0+11,ty0+8),fill=(247,228,177))
    for k in range(3):
        fx=x+28+k*48
        draw.ellipse((fx,y+h-45,fx+20,y+h-24),fill=(79,117,62))

def website_capture(top=True):
    im=Image.new('RGB',(1200,760),CREAM); d=ImageDraw.Draw(im)
    d.rectangle((0,0,1200,31),fill=GREEN); txt(d,(370,9),'HAND-TOSSED, WOOD-FIRED, MADE WITH AMORE',11,(255,249,238),bold=True)
    d.rectangle((0,32,1200,104),fill=CREAM); d.line((0,103,1200,103),fill=LINE,width=1)
    d.ellipse((74,49,112,87),outline=RED,width=2); txt(d,(87,53),'B',26,RED,serif=True,bold=True); txt(d,(121,53),'PIZZERIA',11,INK,bold=True); txt(d,(121,67),'BELLO',11,INK,bold=True)
    txt(d,(500,63),'Our menu     Our story     Find us',13,INK)
    d.rectangle((1020,49,1125,88),fill=RED); txt(d,(1035,63),'Order online  ↗',12,(255,255,255),bold=True)
    if top:
        txt(d,(75,155),'GOOD FOOD. GOOD MOOD.',12,GREEN,bold=True)
        txt(d,(75,205),'A little slice',60,INK,serif=True,bold=True); txt(d,(75,270),'of Italy.',60,RED,serif=True,bold=True)
        txt(d,(78,365),'Big-hearted pizza, made from scratch and pulled',15,MUTED)
        txt(d,(78,390),'hot from the oven. Gather your people —',15,MUTED)
        txt(d,(78,415),"we've got dinner covered.",15,MUTED)
        d.rectangle((78,465,240,512),fill=INK); txt(d,(96,481),'Explore the menu  →',12,(255,255,255),bold=True)
        txt(d,(78,556),'★★★★★  4.9/5 from our lovely locals',12,(173,119,44),bold=True)
        pizza_art(d,(590,133,530,500),seed=1)
        d.ellipse((555,151,645,241),fill=(228,180,78)); txt(d,(574,177),'FRESH FROM',10,INK,bold=True); txt(d,(589,194),'OUR OVEN',10,INK,bold=True)
        txt(d,(600,692),'01 / 04     THE BELLO CLASSIC',11,MUTED,bold=True)
    else:
        txt(d,(76,137),'THE GOOD STUFF',12,GREEN,bold=True)
        txt(d,(76,165),'Made to be shared.',42,INK,serif=True,bold=True)
        txt(d,(76,225),'Long-fermented dough, the good cheese, and toppings that know their place.',14,MUTED)
        cards=[('MARgherita','San Marzano tomato, fior di latte, basil','$14'),('PEPPERONI PICCANTE','Spicy salami, mozzarella, hot honey','$17'),('GARDEN PARTY','Roasted courgette, olives, whipped ricotta','$16')]
        for i,(title,desc,price) in enumerate(cards):
            x=76+i*352; pizza_art(d,(x,280,324,265),seed=i)
            d.rectangle((x+12,293,x+129,319),fill=CREAM); txt(d,(x+21,301),['THE CLASSIC','A CROWD FAVOURITE','GREEN & GOOD'][i],9,GREEN,bold=True)
            txt(d,(x,565),title.title(),18,INK,serif=True,bold=True); txt(d,(x,595),desc,10,MUTED); txt(d,(x+276,564),price,16,RED,serif=True,bold=True)
        d.line((76,637,1122,637),fill=LINE,width=2); txt(d,(76,660),'48-hour slow-fermented dough. That’s the secret.',12,MUTED)
        d.rectangle((940,649,1122,694),outline=INK,width=1); txt(d,(964,664),'See the full menu  →',11,INK,bold=True)
    return im

def set_cell_shading(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)

for name,top in [('website-top.png',True),('website-menu.png',False)]:
    website_capture(top).save(WORK/name)

doc=Document(); sec=doc.sections[0]
sec.top_margin=Inches(.65); sec.bottom_margin=Inches(.6); sec.left_margin=Inches(.75); sec.right_margin=Inches(.75)
styles=doc.styles
styles['Normal'].font.name='Arial'; styles['Normal'].font.size=Pt(10); styles['Normal'].font.color.rgb=RGBColor(*INK)
styles['Title'].font.name='Georgia'; styles['Title'].font.size=Pt(30); styles['Title'].font.bold=True; styles['Title'].font.color.rgb=RGBColor(*INK)
title_ppr=styles['Title']._element.pPr
if title_ppr is not None:
    for node in list(title_ppr):
        if node.tag==qn('w:pBdr'): title_ppr.remove(node)
styles['Heading 1'].font.name='Georgia'; styles['Heading 1'].font.size=Pt(19); styles['Heading 1'].font.bold=True; styles['Heading 1'].font.color.rgb=RGBColor(*GREEN)

p=doc.add_paragraph(style='Title'); p.paragraph_format.space_after=Pt(3); p.add_run('Pizzeria Bello Website Project')
if p._p.pPr is not None:
    for node in list(p._p.pPr):
        if node.tag==qn('w:pBdr'): p._p.pPr.remove(node)
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(9); r=p.add_run('A responsive restaurant website concept'); r.italic=True; r.font.size=Pt(12); r.font.color.rgb=RGBColor(*MUTED)
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(10); p.add_run('Project overview').bold=True
p=doc.add_paragraph('Pizzeria Bello is a warm, modern website concept for a neighborhood pizza restaurant. It introduces the restaurant, highlights its menu, and helps visitors find its location and opening hours. The layout is designed to work on both desktop and mobile screens.')
p.paragraph_format.space_after=Pt(12)
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(4); p.add_run().add_picture(str(WORK/'website-top.png'),width=Inches(6.8))
p=doc.add_paragraph('Website preview: homepage and hero section'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(4)
for r in p.runs: r.italic=True; r.font.size=Pt(8); r.font.color.rgb=RGBColor(*MUTED)
p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(2); p.add_run('The homepage uses a simple call to action, restaurant branding, and a large food visual to introduce the concept at a glance.').font.size=Pt(9)

doc.add_page_break()
p=doc.add_paragraph('Menu and site features',style='Heading 1'); p.paragraph_format.space_after=Pt(8)
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(3); p.add_run().add_picture(str(WORK/'website-menu.png'),width=Inches(6.8))
p=doc.add_paragraph('Website preview: featured pizza menu'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(9)
for r in p.runs: r.italic=True; r.font.size=Pt(8); r.font.color.rgb=RGBColor(*MUTED)
features=[('Menu cards','Three featured pizzas with descriptions and sample prices.'),('Restaurant story','A short introduction to the restaurant and its food.'),('Visit details','Sample address, opening hours, directions, and tap-to-call phone link.'),('Responsive navigation','A compact navigation menu for smaller screens.'),('Visual style','Cream backgrounds, tomato-red accents, green details, and editorial typography.')]
for label,desc in features:
    p=doc.add_paragraph(); p.paragraph_format.left_indent=Inches(.12); p.paragraph_format.space_after=Pt(4)
    r=p.add_run(label+'  '); r.bold=True; r.font.color.rgb=RGBColor(*RED); p.add_run(desc)
p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(8); p.paragraph_format.space_after=Pt(0)
r=p.add_run('Customization note  '); r.bold=True; r.font.color.rgb=RGBColor(*GREEN)
p.add_run('Restaurant contact details, location, prices, and menu items are sample content and can be replaced for the final presentation.')
for r in p.runs: r.font.size=Pt(9)

footer=sec.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.RIGHT
r=footer.add_run('PIZZERIA BELLO  ·  COLLEGE PROJECT'); r.font.size=Pt(8); r.font.color.rgb=RGBColor(*MUTED)
doc.core_properties.title='Pizzeria Bello Website Project'; doc.core_properties.subject='College project overview with website previews'
doc.save(OUT)
print(OUT)
