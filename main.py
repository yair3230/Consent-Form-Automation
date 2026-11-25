DESCRIPTION = "הכניסו כאן את התיאור של הניסוי שלכם"  # בניסוי תתבקש\י...

import copy
import os
from datetime import date
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

TEMPLATE_FILE = "דוגמא לטופס הסכמה מדעת.docx"


def get_template_rPr(paragraph):
    if paragraph.runs and paragraph.runs[0]._r.rPr is not None:
        return copy.deepcopy(paragraph.runs[0]._r.rPr)
    return None


def add_styled_run(paragraph, text, template_rPr=None, underline=False, bold=False, italic=False):
    run = paragraph.add_run(text)
    if template_rPr is not None:
        if run._r.rPr is not None:
            run._r.remove(run._r.rPr)
        run._r.insert(0, copy.deepcopy(template_rPr))
    else:
        run.font.name = "David"
        run.font.size = Pt(12)

    rPr = run._r.get_or_add_rPr()
    if rPr.find(qn("w:rtl")) is None:
        rPr.append(parse_xml(f'<w:rtl {nsdecls("w")}/>'))
    run.underline = underline
    run.bold = bold
    run.italic = italic
    return run


def replace_paragraph_runs(paragraph, segments):
    """Replaces text inside a template paragraph while preserving its RTL <w:pPr>."""
    p = paragraph._p
    template_rPr = get_template_rPr(paragraph)

    for child in list(p):
        if not child.tag.endswith("pPr"):
            p.remove(child)

    for text, underline, bold in segments:
        add_styled_run(paragraph, text, template_rPr=template_rPr, underline=underline, bold=bold)


def insert_signature_table(doc, target_p, full_name_val, date_val, name_label, template_rPr=None):
    """
    Inserts a borderless 2x3 RTL table directly before `target_p` with:
    - Column 1 (Right): Printed Full Name
    - Column 2 (Center): Full Name as Italicized Autograph
    - Column 3 (Left): Auto-filled Date
    """
    table = doc.add_table(rows=2, cols=3)
    table.autofit = False
    target_p._p.addprevious(table._tbl)

    tblPr = table._tbl.tblPr
    for tag in ("w:bidiVisual", "w:tblBorders", "w:jc", "w:tblW", "w:tblLayout"):
        existing = tblPr.find(qn(tag))
        if existing is not None:
            tblPr.remove(existing)

    section = doc.sections[0]
    usable_twips = int((section.page_width - section.left_margin - section.right_margin) / 635)
    col_twips = [int(usable_twips * 0.36), int(usable_twips * 0.38), int(usable_twips * 0.26)]

    tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{sum(col_twips)}" w:type="dxa"/>'))
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
    ))

    grid_cols = table._tbl.tblGrid.findall(qn("w:gridCol"))
    for grid_col, w in zip(grid_cols, col_twips):
        grid_col.set(qn("w:w"), str(w))

    # Each cell tuple: (text, underline, italic)
    rows_content = [
        [
            (f"   {full_name_val}   ", True, False),
            (f"   {full_name_val}   ", True, True),
            (f"   {date_val}   ", True, False),
        ],
        [
            (name_label, False, False),
            ("חתימה", False, False),
            ("תאריך", False, False),
        ],
    ]

    for r_idx, row_items in enumerate(rows_content):
        row = table.rows[r_idx]
        for c_idx, ((text, is_underlined, is_italic), w) in enumerate(zip(row_items, col_twips)):
            cell = row.cells[c_idx]
            tcPr = cell._tc.get_or_add_tcPr()
            existing_tcW = tcPr.find(qn("w:tcW"))
            if existing_tcW is not None:
                tcPr.remove(existing_tcW)
            tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w}" w:type="dxa"/>'))

            cell_p = cell.paragraphs[0]
            cell_pPr = cell_p._p.get_or_add_pPr()
            cell_pPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
            cell_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_styled_run(
                cell_p,
                text,
                template_rPr=template_rPr,
                underline=is_underlined,
                italic=is_italic,
            )


def fill_from_template(template_path, data):
    doc = docx.Document(template_path)
    paragraphs = list(doc.paragraphs)
    after_exp_desc = False
    sig_index = 0
    i = 0

    while i < len(paragraphs):
        p = paragraphs[i]
        txt = p.text

        if "שם משפחה:" in txt:
            replace_paragraph_runs(p, [
                ("שם משפחה: ", False, False),
                (f" {data['last_name']} ", True, False),
                ("     שם פרטי: ", False, False),
                (f" {data['first_name']} ", True, False),
                ("     מס' ת.ז.: ", False, False),
                (f" {data['id_num']} ", True, False),
            ])

        elif "כתובת:" in txt:
            replace_paragraph_runs(p, [
                ("כתובת: ", False, False),
                (f" {data['address']} ", True, False),
                ("                  מיקוד: ", False, False),
                (f" {data['zip_code']} ", True, False),
            ])

        elif "הוסבר לי" in txt:
            replace_paragraph_runs(p, [
                ('ב. מצהיר/ה בזה כי הוסבר לי ע"י (שם החוקר/ת) כדלהלן: ', False, False),
                (f" {data['researcher_name']} ", True, False),
                (
                '  כי המחקר מתבצע במסגרת הקורס "מעבדה ניסויית בפסיכולוגיה", וכי המחקר אושר לביצוע על ידי ועדת האתיקה של היחידה ללימודים בינתחומיים.',
                False, False),
            ])

        elif "בניסוי תתבקש/י" in txt:
            replace_paragraph_runs(p, [
                ("בניסוי תתבקש/י ", False, False),
                (f" {data['exp_description']} ", True, False),
            ])
            after_exp_desc = True

        elif after_exp_desc and txt.strip() and set(txt.strip()) <= {"_"}:
            replace_paragraph_runs(p, [])
            after_exp_desc = False

        elif "____" in txt:
            after_exp_desc = False
            rPr = get_template_rPr(p)
            if sig_index == 0:
                insert_signature_table(doc, p, data["full_name"], data["date"], "שם המשתתף/ת במחקר", rPr)
            else:
                insert_signature_table(doc, p, data["researcher_name"], data["date"], "שם החוקר המסביר", rPr)
            sig_index += 1

            p._p.getparent().remove(p._p)
            if i + 1 < len(paragraphs) and "חתימה" in paragraphs[i + 1].text:
                paragraphs[i + 1]._p.getparent().remove(paragraphs[i + 1]._p)
                i += 1
        else:
            after_exp_desc = False

        i += 1

    return doc


def build_from_scratch_rtl(data):
    """Fallback if template file is not in the same folder."""
    doc = docx.Document()
    normal_pPr = doc.styles["Normal"]._element.get_or_add_pPr()
    normal_pPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
    for section in doc.sections:
        section._sectPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))

    def add_rtl_p(segments, center=False):
        p = doc.add_paragraph()
        pPr = p._p.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for text, underline, bold in segments:
            add_styled_run(p, text, underline=underline, bold=bold)
        return p

    add_rtl_p([("טופס הסכמה מדעת להשתתפות במחקר", True, True)], center=True)
    add_rtl_p([('אני הח"מ:', False, True)])
    add_rtl_p([
        ("שם משפחה: ", False, False), (f" {data['last_name']} ", True, False),
        ("     שם פרטי: ", False, False), (f" {data['first_name']} ", True, False),
        ("     מס' ת.ז.: ", False, False), (f" {data['id_num']} ", True, False),
    ])
    add_rtl_p([
        ("כתובת: ", False, False), (f" {data['address']} ", True, False),
        ("                  מיקוד: ", False, False), (f" {data['zip_code']} ", True, False),
    ])
    add_rtl_p([("א. מצהיר/ה בזה כי אני מסכים/ה להשתתף במחקר כמפורט במסמך זה.", False, False)])
    add_rtl_p([
        ('ב. מצהיר/ה בזה כי הוסבר לי ע"י (שם החוקר/ת) כדלהלן: ', False, False),
        (f" {data['researcher_name']} ", True, False),
        (
        ' כי המחקר מתבצע במסגרת הקורס "מעבדה ניסויית בפסיכולוגיה", וכי המחקר אושר לביצוע על ידי ועדת האתיקה של היחידה ללימודים בינתחומיים.',
        False, False),
    ])
    add_rtl_p([("תאור הניסוי:", False, True)])
    add_rtl_p([("בניסוי תתבקש/י ", False, False), (f" {data['exp_description']} ", True, False)])
    add_rtl_p([("   3. כי אני חפשי/ה להפסיק בכל עת את השתתפותי בניסוי וכי איני צפוי לשום סנקציה אם אבחר לעשות כן.",
                False, False)])
    add_rtl_p([("   4. כי מובטחת סודיות באשר לזהותי האישית בפרסומים מדעיים.", False, False)])
    add_rtl_p([("   5. כי בכל בעיה הקשורה למחקר אוכל לפנות לחוקר/ת הראשית של המחקר", False, False)])
    add_rtl_p(
        [('ג. הנני מצהיר/ה בזה כי את הסכמתי הנ"ל נתתי מרצוני החופשי וכי הבינותי את כל האמור לעיל.', False, False)])

    p_sig1 = doc.add_paragraph()
    insert_signature_table(doc, p_sig1, data["full_name"], data["date"], "שם המשתתף/ת במחקר")
    p_sig1._p.getparent().remove(p_sig1._p)

    add_rtl_p([("ד. הצהרת החוקר/ת", False, True)])
    add_rtl_p([(
               '    ההסכמה הנ"ל נתקבלה על ידי וזאת לאחר שהסברתי למשתתף/ת במחקר כל האמור לעיל ווידאתי שכל הסבריי הובנו על ידו/ידה.',
               False, False)])

    p_sig2 = doc.add_paragraph()
    insert_signature_table(doc, p_sig2, data["researcher_name"], data["date"], "שם החוקר המסביר")
    p_sig2._p.getparent().remove(p_sig2._p)

    return doc


def main():
    if DESCRIPTION == "הכניסו כאן את התיאור של הניסוי שלכם":
        print(
            'כדי להשתמש בכלי הנ"ל, כנסו לקוד main.py, לכו לשורה הראשונה, ומלאו את התיאור של הניסוי שלכם בתוך המרכאות.')
        return
    print("--- מילוי טופס הסכמה מדעת להשתתפות במחקר ---")
    last_name = input("שם משפחה: ").strip()
    first_name = input("שם פרטי: ").strip()
    full_name = f"{first_name} {last_name}".strip()
    today_date = date.today().strftime("%d.%m.%Y")

    data = {
        "last_name": last_name,
        "first_name": first_name,
        "full_name": full_name,
        "id_num": input("מס' ת.ז.: ").strip(),
        "address": input("כתובת: ").strip(),
        "zip_code": input("מיקוד: ").strip(),
        "researcher_name": input("שם מלא של החוקר/ת שהסביר/ה: ").strip(),
        "exp_description": DESCRIPTION,
        "date": today_date,
    }

    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, TEMPLATE_FILE)

    if os.path.exists(template_path):
        doc = fill_from_template(template_path, data)
    else:
        doc = build_from_scratch_rtl(data)

    safe_name = "".join(c for c in full_name if c.isalnum() or c in (" ", "-", "_")).strip()
    output_path = os.path.join(script_dir, f"{safe_name or 'טופס_הסכמה_מלא'}.docx")
    doc.save(output_path)
    print(f"\nהמסמך נשמר בהצלחה ({today_date}): {output_path}")


if __name__ == "__main__":
    main()
