from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.platypus import Table, TableStyle, PageBreak, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.graphics.shapes import Drawing, Rect, Line, String
from pypdf import PdfReader, PdfWriter
from copy import deepcopy
import os
import math

# =====================================================
# DIRECTORY SETUP
# =====================================================
INPUT_DIR = "Inputs"
OUTPUT_DIR = "Outputs"

# Ensure directories exist
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# =====================================================
# CONFIG
# =====================================================
USE_PDF_LETTERHEAD = True

# =====================================================
# DEFAULTS
# =====================================================
DEFAULT_BID_NUMBER = "PMWH/DP Funded/Works/NCB/01/2082-83"
DEFAULT_CONTRACT_NAME = "Construction of 25-Bed Hospitals in Chandragiri Municipality and Mahalaxmi Municipality"
DEFAULT_DATE = "April 18, 2026"
DEFAULT_BIDDER_NAME = "Sabin Chaulagain"
DEFAULT_EMPLOYER_NAME = "Paropakar Maternity and Women's Hospital"
DEFAULT_EMPLOYER_ADDRESS = "Thapathali, Kathmandu"
DEFAULT_ORG_NAME = "Eco Builders and Engineers Pvt. Ltd."
DEFAULT_ORG_ADDRESS = "Gokarneshwor-06, Kathmandu"
DEFAULT_EMAIL = "ecobuilders12@gmail.com"


# =====================================================
# LETTERHEAD FILES
# =====================================================
LETTERHEAD_FILE = "letterhead.pdf"   # portrait only
LANDSCAPE_LETTERHEAD_FILE = "letterhead_landscape.pdf"

def get_landscape_letterhead():
    full_path = os.path.join(INPUT_DIR, LANDSCAPE_LETTERHEAD_FILE)
    if os.path.exists(full_path):
        print("📄 Using external landscape letterhead")
        return "PDF"
    else:
        print("⚠ landscape letterhead not found – using auto-generated header")
        return "AUTO"

# =====================================================
# MERGE FUNCTION
# =====================================================
def merge_with_letterhead(letterhead_path, content_path, output_path):
    try:
        content = PdfReader(content_path)
        writer = PdfWriter()

        for i in range(len(content.pages)):
            base = PdfReader(letterhead_path).pages[0]
            base = deepcopy(base)

            base.merge_page(content.pages[i])
            writer.add_page(base)

        with open(output_path, "wb") as f:
            writer.write(f)

        if os.path.exists(content_path):
            os.remove(content_path)

        print(f"✅ Created: {os.path.basename(output_path)}")
    except Exception as e:
        print(f"⚠️ Warning: Could not merge with letterhead ({e}). Outputting standard PDF instead: {os.path.basename(output_path)}")
        if os.path.exists(content_path):
            os.rename(content_path, output_path)

# =====================================================
# STYLES
# =====================================================
styles = getSampleStyleSheet()

date_style = ParagraphStyle("Date", parent=styles["Normal"], alignment=TA_RIGHT, fontSize=8)
body_style = ParagraphStyle("Body", parent=styles["Normal"], fontSize=8, leading=10)
body_style_center = ParagraphStyle("Bodycenter", parent=styles["Normal"], fontSize=8, leading=10, alignment=TA_CENTER)
title_style = ParagraphStyle("Title", parent=styles["Normal"], alignment=TA_CENTER, fontSize=12)
body_style_10 = ParagraphStyle("Body10", parent=styles["Normal"], fontSize=10, leading=14, spaceAfter=6)
date_style_10 = ParagraphStyle("Date10", parent=styles["Normal"], alignment=TA_RIGHT, fontSize=10, spaceAfter=6)
header_title = ParagraphStyle("HdrTitle", parent=styles["Normal"], alignment=TA_CENTER, fontSize=16, fontName="Helvetica-Bold")
header_small = ParagraphStyle("HdrSmall", parent=styles["Normal"], alignment=TA_CENTER, fontSize=10, textColor=colors.grey)

def add_org_chart_signature(canvas, doc):
    canvas.saveState()
    try:
        sig = ImageReader(os.path.join(INPUT_DIR, "bidder_signature.png"))
        stamp = ImageReader(os.path.join(INPUT_DIR, "stamp.png"))

        # Signature (bottom-left)
        canvas.drawImage(
            sig,
            x=2.2*inch,
            y=0.6*inch,
            width=2.2*inch,
            height=1.18*inch,
            mask='auto',
            preserveAspectRatio=True
        )

        # Stamp (bottom-right of signature)
        canvas.drawImage(
            stamp,
            x=3.8*inch,
            y=0.6*inch,
            width=1.4*inch,
            height=1.18*inch,
            mask='auto',
            preserveAspectRatio=True
        )

    except Exception as e:
        pass # Suppress warning to keep output clean

    canvas.restoreState()
    
def draw_landscape_with_signature(canvas, doc):
    draw_landscape_letterhead(canvas, doc)
    add_org_chart_signature(canvas, doc)



def add_mobilization_schedule_stamps(canvas, doc):
    canvas.saveState()
    try:
        img1 = ImageReader(os.path.join(INPUT_DIR, "stamp.png"))
        canvas.drawImage(img1, 
                        x=5.5*inch, 
                        y=5.8*inch, 
                        width=2.5*inch, 
                        height=1.18*inch, 
                        mask='auto',
                        preserveAspectRatio=True)
        
        img2 = ImageReader(os.path.join(INPUT_DIR, "bidder_signature.png"))
        canvas.drawImage(img2,
                        x=4.5*inch,
                        y=5.8*inch,
                        width=2.5*inch,
                        height=1.18*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
    except Exception as e:
        pass
    finally:
        canvas.restoreState()
        
def add_technical_purposal_stamps(canvas, doc):
    canvas.saveState()
    try:
        img1 = ImageReader(os.path.join(INPUT_DIR, "stamp.png"))
        canvas.drawImage(img1, 
                        x=5.5*inch, 
                        y=0.5*inch, 
                        width=2.5*inch, 
                        height=1.18*inch, 
                        mask='auto',
                        preserveAspectRatio=True)
        
        img2 = ImageReader(os.path.join(INPUT_DIR, "bidder_signature.png"))
        canvas.drawImage(img2,
                        x=4.5*inch,
                        y=0.5*inch,
                        width=2.5*inch,
                        height=1.18*inch,
                        mask='auto',
                        preserveAspectRatio=True)
      
    except Exception as e:
        pass
    finally:
        canvas.restoreState()        

def add_poa_stamps(canvas, doc):
    canvas.saveState()
    try:
        img1 = ImageReader(os.path.join(INPUT_DIR, "stamp.png"))
        canvas.drawImage(img1, 
                        x=6.5*inch, 
                        y=3*inch, 
                        width=1.5*inch, 
                        height=1.18*inch, 
                        mask='auto',
                        preserveAspectRatio=True)
    
        img2 = ImageReader(os.path.join(INPUT_DIR, "bidder_signature.png"))
        canvas.drawImage(img2,
                        x=0.5*inch,
                        y=4.2*inch,
                        width=1.5*inch,
                        height=1.55*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
        img3 = ImageReader(os.path.join(INPUT_DIR, "owner_signature.png"))
        canvas.drawImage(img3,
                        x=5.5*inch,
                        y=3*inch,
                        width=1.5*inch,
                        height=1.18*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
    except Exception as e:
        pass
    finally:
        canvas.restoreState()
        
def add_selfdec_stamps(canvas, doc):
    canvas.saveState()
    try:
        img1 = ImageReader(os.path.join(INPUT_DIR, "stamp.png"))
        canvas.drawImage(img1, 
                        x=2.5*inch, 
                        y=2.8*inch, 
                        width=1.5*inch, 
                        height=1.18*inch, 
                        mask='auto',
                        preserveAspectRatio=True)
    
        img2 = ImageReader(os.path.join(INPUT_DIR, "bidder_signature.png"))
        canvas.drawImage(img2,
                        x=0.75*inch,
                        y=4.55*inch,
                        width=1.5*inch,
                        height=1.8*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
    except Exception as e:
        pass
    finally:
        canvas.restoreState()
        
def add_bid_letter_stamp(canvas, doc):
    canvas.saveState()
    try:
        img1 = ImageReader(os.path.join(INPUT_DIR, "stamp.png"))
        canvas.drawImage(img1, 
                        x=1.3*inch, 
                        y=1.8*inch, 
                        width=1.5*inch, 
                        height=1.18*inch, 
                        mask='auto',
                        preserveAspectRatio=True)
        
        img2 = ImageReader(os.path.join(INPUT_DIR, "bidder_signature.png"))
        canvas.drawImage(img2,
                        x=0.5*inch,
                        y=1.8*inch,
                        width=1.5*inch,
                        height=1.18*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
    except Exception as e:
        pass
    finally:
        canvas.restoreState()        
        

def add_paragraphs(elements, texts, style, spacing=10):
    for t in texts:
        elements.append(Paragraph(t, style))
        elements.append(Spacer(1, spacing))


# =====================================================
# TECHNICAL BID
# =====================================================
def create_technical_bid_pdf(bid_number, contract_name, bid_date,
                             bidder_name, org_name, org_address,
                             email_address, employer_name, employer_address):

    safe = bid_number.replace("/", "-")
    temp = os.path.join(OUTPUT_DIR, "temp_tech.pdf")
    final = os.path.join(OUTPUT_DIR, f"Letter_of_Technical_Bid_{safe}.pdf")

    doc = SimpleDocTemplate(temp, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=150)
    elements = []

    elements.append(Paragraph(f"<b>Date: {bid_date}</b>", date_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("<b>LETTER OF TECHNICAL BID</b>", title_style))
    elements.append(Spacer(1, 10))

    add_paragraphs(elements, [
        f"<b>To,<br/>{employer_name},<br/>{employer_address}</b>",
        "We, the undersigned, declare that:",
        "a) We have examined and have no reservation to the Bidding Documents, including Addenda issued in according with instructions to Bidder (ITB) Clause 8;",
        f"b) We offer to execute in conformity with the Bidding Documents the Following Works:<br/><b>{contract_name}</b> and Invitation for Bid No.: <b>{bid_number}</b>",
        "c) Our Bid consisting of the Technical and the Price Bid shall be valid for a period of 180 day from the date fixed for the bid submission deadline in according with the bidding document and it shall remain binding upon us and may be accepted at any time before the expiration of that Period.",
        "d) Our firm, including any subcontractors or suppliers for any part of the Contract, have nationalities from eligible countries in accordance with ITB 4.2 and meet the requirement of ITB3.5 & 3.5",
        "e) We are not participating as a Bidder or as a subcontractor in more than one bid in the bidding process in accordance with ITB 4.3 9 (e). Other than alternative offers submitted in accordance with ITB 13,",
        "f) Our firm, its affiliates or subsidiaries, including any Subcontractor or suppliers for any part of the contract has not been declared ineligible by DP, under the Employers country laws or official regulations or by an act of compliance with a decision of the United Nations Security Council;",
        "g) We are not a government owned entity / we are a government owned entity but meet the Requirements of ITB 4.5;",
        "h) We declare that we including any subcontractor or suppliers for any part of the contract do not have any conflict of interest in accordance with ITB 4.3 and we have not been punished for an offense relating to the concerned profession or business.",
        "i) We declare that we are solely responsible for the authenticity of the documents submitted by us.",
        "j) We agree to permit the Employer/DP or its representative to inspect our accounts and records and other documents relating to the bid submission and to have them audited by auditors appointed by the Employer.",
        "k) If our Bid is accepted we commit to mobilizing key equipment and personnel in accordance with the requirements set forth in section III (Evaluation and Qualification Criteria) and our technical proposal, or as otherwise agreed with the employer.",
        "l) We declare that we have not running contracts more than five (5) in accordance with ITB 4.9",
        f"Name: <b>{bidder_name}</b><br/>In the Capacity of Attorney Person<br/><br/><br/><br/><br/><br/><br/>Signed...<br/>Duly authorized to sign the Bid for and on behalf of <b>{org_name}</b><br/>Date: <b>{bid_date}</b>"
    ], body_style)

    doc.build(elements, onFirstPage=add_bid_letter_stamp)

    merge_with_letterhead(os.path.join(INPUT_DIR, LETTERHEAD_FILE), temp, final)


# =====================================================
# PRICE BID
# =====================================================
def create_price_bid_pdf(bid_number, contract_name, bid_date,
                         bidder_name, org_name, org_address,
                         email_address, employer_name, employer_address):

    safe = bid_number.replace("/", "-")
    temp = os.path.join(OUTPUT_DIR, "temp_price.pdf")
    final = os.path.join(OUTPUT_DIR, f"Letter_of_Price_Bid_{safe}.pdf")

    doc = SimpleDocTemplate(temp, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=150)
    elements = []

    elements.append(Paragraph(f"<b>Date: {bid_date}</b>", date_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("<b>LETTER OF PRICE BID</b>", title_style))
    elements.append(Spacer(1, 10))

    body_texts = [
        f"<b>To,<br/>{employer_name},<br/>{employer_address}</b>",
        "We, the undersigned, declare that:",
        "a) We have examined and have no reservation to the Bidding Documents, including Addenda issued in according with instructions to Bidder (ITB) Clause 8;",
        f"b) We offer to execute in conformity with the Bidding Documents the Following Works:<br/><b>{contract_name}</b> and Invitation for Bid No.: <b>{bid_number}</b>",
        "c) The total price of our Bid, excluding any discount offered in item (d) below is <b>AS PER BOQ</b> or when left blank is the Bid Price indicated in the Bill of Quantities.",
        "d) The discount offered and the methodologies for their application are: <b> None</b>",
        "e) Our Bid shall be valid for a period of <b>150</b> days from the date fixed for the bid submission deadline in according with the bidding document and it shall remain binding upon us and may be accepted at any time before the expiration of that Period.",
        "f) If our bid is accepted, we commit to obtain a performance security in accordance with the Bidding Document:",
        "g) We have paid, or will pay the following commissions, gratuities, or fees with respect to the bidding process or execution of the Contract:",
    ]
    add_paragraphs(elements, body_texts, body_style)

    price_table = [[
        Paragraph(f"<b>Name of recipient</b><br/>................<br/>................", body_style_center),
        Paragraph(f"<b>Address</b><br/>.................<br/>.................", body_style_center),
        Paragraph(f"<b>Reason</b><br/>..................<br/>..................", body_style_center),
        Paragraph(f"<b>Amount</b><br/>..................<br/>..................", body_style_center),
    ]]
    elements.append(Table(price_table, colWidths=[1.7 * inch, 1.7 * inch, 1.7 * inch, 1.7 * inch]))
    elements.append(Spacer(1, 12))

    footer_texts = [        
        "h) We understand that this bid, together with your written acceptance thereof included in your notification of award, shall constitute a binding contract between us, until a formal contract is prepared and executed;",
        "i) We understand that you are not bound to accept the lowest evaluated bid or any other bid that you may receive and",
        "j) We declare that we are solely responsible for the authenticity of the documents submitted by us.",
        "k) We agree to permit the Employer/DP or its representative to inspect our accounts and records and other documents relating to the bid submission and to have them audited by author appointed by the Employer.",
        f"Name: <b>{bidder_name}</b><br/>In the Capacity of Attorney Person<br/><br/><br/><br/><br/><br/><br/><br/>Signed...<br/>Duly authorized to sign the Bid for and on behalf of <b>{org_name}</b><br/>Date: <b>{bid_date}</b>"
    ]
    add_paragraphs(elements, footer_texts, body_style)

    doc.build(elements, onFirstPage=add_bid_letter_stamp)

    merge_with_letterhead(os.path.join(INPUT_DIR, LETTERHEAD_FILE), temp, final)


# =====================================================
# POA + DECLARATION
# =====================================================
def create_poa_with_declaration_pdf(
    bid_number,
    contract_name,
    bid_date,
    employer_name,
    employer_address,
    bidder_name,
    org_name,
    org_address,
    letterhead_file,
    output_name
):

    temp_poa = os.path.join(OUTPUT_DIR, "temp_poa_single.pdf")
    temp_dec = os.path.join(OUTPUT_DIR, "temp_dec_single.pdf")
    temp_merge = os.path.join(OUTPUT_DIR, "temp_merge_pages.pdf")
    full_output_name = os.path.join(OUTPUT_DIR, output_name)


    # --- PAGE 1 ---
    doc1 = SimpleDocTemplate(
        temp_poa, pagesize=A4,
        leftMargin=40, rightMargin=40,
        topMargin=150, bottomMargin=40
    )

    elements = []
    elements.append(Paragraph(f"<b>Date: {bid_date}</b>", date_style_10))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph(f"To,<br/>{employer_name},<br/>{employer_address}", body_style_10))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("<u><b>POWER OF ATTORNEY</b></u>", title_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(
        f"Known all men by these presents that we the undersigned, Board of Directors lawfully authorized to represent and act on behalf of the said firm under the Company Act do hereby authorize <b>{bidder_name}</b> of <b>{org_name}</b> Having its head office at {org_address}, Whose specimen signature appears as given below to run all business activities for, negotiable with the Employer, dealing with running bill , Final bill and all the task of related offices operating by his single signature or to authorize any other person to operate on behalf of his authorization within Nepal and abroad, This undersigned shall acknowledge the legal effects of the signature of the said attorney holder after the signing and sealing of the power of attorney.",
         body_style_10
    ))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Whereas, <b>{employer_name}</b>", body_style_10))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"For: <b>{contract_name}</b> and Invitation for Bids No.: <b>{bid_number}</b>", body_style_10))

    elements.append(Spacer(1, 60))
    elements.append(Paragraph("<b>………………………</b>", body_style_10))
    elements.append(Paragraph("<b>Signature of Authorized Representative</b>", body_style_10))
    elements.append(Paragraph(f"<b>{bidder_name}</b>", body_style_10))

    elements.append(Spacer(1, 35))
    elements.append(Paragraph("………………………", date_style_10))
    elements.append(Paragraph(f"<b>{bidder_name}</b>", date_style_10))
    elements.append(Paragraph("<b>Managing Director</b>", date_style_10))
    elements.append(Paragraph(f"<b>{org_name}</b>", date_style_10))
    elements.append(PageBreak())
    doc1.build(elements, onFirstPage=add_poa_stamps)

    # --- PAGE 2 ---
    doc2 = SimpleDocTemplate(
        temp_dec, pagesize=A4,
        leftMargin=40, rightMargin=40,
        topMargin=150, bottomMargin=40
    )

    elements.append(Paragraph(f"<b>Date: {bid_date}</b>", date_style_10))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(f"<b>To,<br/>{employer_name},<br/>{employer_address}</b>", body_style_10))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(f"<b>For: {contract_name} and Invitation for Bids No.: {bid_number}</b>", body_style_10))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("<u><b>LETTER OF DECLARATION</b></u>", title_style))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Dear Sir,", body_style_10))
    elements.append(Paragraph(
        "This is to certify that we are not ineligible to participate in the bid, has no conflict of interest in the proposed bid procurement proceeding and have not been punished for the profession or businesses related offences.",
        body_style_10
    ))

    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Thanking you.", body_style_10))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Sincerely Yours,", body_style_10))
    elements.append(Spacer(1, 52))
    elements.append(Paragraph("Signature", body_style_10))
    elements.append(Paragraph(f"Name: <b>{bidder_name}</b>", body_style_10))
    elements.append(Paragraph("Designation: <b>Managing Director</b>", body_style_10))
    elements.append(Paragraph(f"Authorized to sign on behalf of : <b>{org_name}</b>", body_style_10))
    elements.append(Spacer(1, 64))
    elements.append(Paragraph("Office stamp of the organization……………………", body_style_10))
    doc2.build(elements, onFirstPage=add_selfdec_stamps)


    # --- MERGE CONTENT PAGES ---
    writer = PdfWriter()
    writer.add_page(PdfReader(temp_poa).pages[0])
    writer.add_page(PdfReader(temp_dec).pages[0])

    with open(temp_merge, "wb") as f:
        writer.write(f)

    os.remove(temp_poa)
    os.remove(temp_dec)

    merge_with_letterhead(os.path.join(INPUT_DIR, letterhead_file), temp_merge, full_output_name)



# =====================================================
# WORK METHODOLOGY
# =====================================================
def create_work_methodology_pdf(bid_number, contract_name, bid_date,
                                org_name, org_address, email_address):

    safe = bid_number.replace("/", "-")
    temp = os.path.join(OUTPUT_DIR, "temp_work.pdf")
    final = os.path.join(OUTPUT_DIR, f"Work_Methodology_{safe}.pdf")

    doc = SimpleDocTemplate(temp, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=150)
    elements = []

    elements.append(Paragraph("<b>WORK METHODOLOGY</b>", title_style))
    elements.append(Spacer(1, 10))
    body_texts = [
        "<b>A. <u><i>Preliminary Site Organization Chart :</i></u></b>",
        "A Preliminary Site Organization Chart Showing the organization set up the considering sufficient in all respect to execute the work as per the specification within the condition of contract is attached herewith. All the personnel proposed for this project will be worked for whole contract period. A brief description of the relation of the staff within the site organization and its relation and business with head office is also detailed here under.",
        "<b>B. <u><i>Narrative Description of site organization Chart :</i></u></b>",
        "The Project Manager is the responsible for successful implementation smooth execution and timely completion of whole work. He will receive information from head office, give suggestion, progress report from administration, technical and account section as well as guidelines, special instruction, suggestion and take timely action by issuing notice, instruction making on the spot-study in a bid achieve timely execution maintaining. cordial relation and good co-operation, However, the company also look after the project as per requirement and visit time to time. The whole site organization under the project manager will be divided informally in two section i.e. Technical, Administrative and Financial section.",
        "<b>(i) <u><i>Technical Section :</i></u></b>",
        "This section shall comprise at least one Project Manager, one Experienced Civil Engineer, one Design Engineer, one Auto CAD operator, sufficient no’s of construction supervisor, mechanic operators etc. This section will study drawing and specifications, take measurement of works, prepare running bills and submits the same to employer for payment on receipt of instructions from employer's representative and consultant regarding changes on include such changes in works. it will also monitor progress and take necessary action in time to keep pace with targeted schedule. This section will also look after machine requirement, its maintenance and running through mechanics and operators its job is also to prepare work schedule manpower schedule materials schedule, design drawing, survey, laboratory testing etc. However it can be more as the work is going to take in company.",
        "<b>(ii) <u><i>Administrative/Account Section :</i></u></b>",
        "This section will deal correspondences maintain records of staff and workmen, appoint personal, look after insurance, compensation, safety measures, attendance, records of leave, afford and maintain transportation facilities, housing, medical and sanitation facilities. it will also increase or decrease substitute staffs, workmen depending on manpower schedule keeping close touch with head office normally all information area received at site-office through administrative section and also this section is though not directly connected to construction-activities, indispensable for construction management's. As the perfect is concerned with directly to the community, this section will comprise 1 Administrative cum procurement manager, one accountant one store keeper, one assistant, guards and peons etc. This section will also provide manpower-material and finance-schedule consultation with technical schedule to the project manager.",
    ]
    add_paragraphs(elements, body_texts, body_style_10)
    elements.append(PageBreak())

    elements.append(Paragraph("PROPOSED WORK METHODOLOGY", title_style))
    elements.append(Paragraph(f"<u><b>Name of the Applicant : {org_name}</b></u>", title_style))
    elements.append(Spacer(1, 10))

    body_texts = [
          "<b><u><i>PROJECT/CONSTRUCITON MANAGER :</i></u></b>",
          "Overall management including construction management supervision correspondence with consultant and employer as well as dealing with consultant and employer and coordination with Company office.",
          "<b><u><i>ADMINISTRATIVE SECTION :</i></u></b>",
          "Personnel administrative of staffs, employment and retrenchment of staffs, management of safety measures for employer’s materials and workers.",
          "<b><u><i>CONSTRUCTION SECTION :</i></u></b>",
          "Surveying, Designing if required execution of fields works, management of works preparation of bills, quality control, planning, scheduling and Progress review to asset requirements of construction materials equipment’s etc.",
          "<b><u><i>ACCOUNT/PROCUREMENT SECTION :</i></u></b>",
          "Surveying, Designing if required execution of fields works, management of works preparation of bills, quality control, planning, scheduling and Progress review to asset requirements of construction materials equipment’s etc.",
          "<b><u><i>Description of Relation between Head office and Site office :</i></u></b>",
          "Head office will be responsible for fulfilling materials, manpower and fund insufficient at site. dealing with owner/consultant at their head office frequent site visit, progress review and assist the site management by solving its various problems, Site management will do everything regarding work execution and contract to head office from time to time.",
    ]
    add_paragraphs(elements, body_texts, body_style_10)
    elements.append(PageBreak())
    
    elements.append(Spacer(1, 15))
    elements.append(Paragraph("<b>CONSTRUCTION PLAN AND METHODOLOGY</b>", title_style))
    elements.append(Spacer(1, 10))

    body_texts = [
        f"<b>{contract_name}, Contract Identification No.:{bid_number}</b>",
        "The contractor upon signing of contract agreement will plan and start to mobilize the necessary Equipment, manpower, and other logistics' within 15 days after signing of the contract. The Insurance cover will be provided from the start date to the end of Defects Liability period.",
        "Detail work program will be prepared and submitted to the Engineer for his approval after the Contract agreement is made.",
        "A site office will be established at a suitable place within the contract location. The site office, Labor camp as well as working area will be kept in a hygienic condition. Provision of toilets for Labor and employees will be made to avoid public nuisance and pollution of watercourses and Air. Depending up on the source of water, adequate drinking water and other water will be Supplied for the use of own staff and labor.",
        "The work shall be carried out in accordance with the relevant quality standards, test procedures or codes of practice as specific in the Technical Specification. Similarly, the method of working To be adopted will be such as to enable the satisfactory completion of the works before scheduled Date completion.",
        "The sequence of construction activities will be so managed that the double handling of materials Will be avoided.",
        "The source of construction material will get approved from the Engineer. Test sample of the Material and shop drawings if any require will be submitted early to allow sufficient time to the Engineer for review and approval.",
        "Safety and security of life and property is one of the very important part of the contract. For this/Fencing, guards, warning signs, safety helmet, adequate light for night work (if required and Permitted) etc. will be provided as required and to the satisfaction of the Engineer. Upon completion of the works the site will be cleaned and all temporary offices, sanitary units, Stores, workshop, plant, tools, rubbish etc. will be removed from the side.",
    ]
    add_paragraphs(elements, body_texts, body_style_10)
    elements.append(PageBreak())
    
    elements.append(Spacer(1, 15))
    elements.append(Paragraph("<b>WORK METHOD STATEMENT</b>", title_style))
    elements.append(Spacer(1, 10))

    body_texts = [
        f"<b>{contract_name}, Contract Identification No.:{bid_number}</b>",
        "The full authority for the management of the whole work will be given to the Contract Manager.",
        "The Contract Manager shall be responsible for the management of technical, administrative, Financial, procurement as well as mechanical. The Contract Manager shall also give the final Decision of major problems. Contract manager can deliver his partial authority to concerned Section chief as and when he feels necessary.",
        "Technical section will be responsible for quality control, contraction supervision, survey, Measurement, preparation of bill etc. Any problems which cannot be solved shall be informed to The Project Manager.",
        "Mechanical Section will be responsible for repair of minor defect of any equipment or vehicles As well as regular servicing.",
        "Financial section will be responsible for management of smooth cash flow, keeping accurate Records of debit and credit. Timely information shall be given to the Site Manager in case of Shortage of money.",
        "Procurement section shall be responsible for purchasing of any kind of material required to the Project. This section will also keep the records of stock. It shall provide procurement program to The financial section on time.",
        "Administrative section shall be responsible for internal administrative control as well as public Relations",
    ]
    add_paragraphs(elements, body_texts, body_style_10)
    elements.append(Paragraph("<b><u>Relations Description of relationship between Head office & Site Office Management</u></b>", title_style))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("The full authority will be given to the Contract Manager. The head office, as requested by the Contract Manager will furnish financial management. Site office will inform all its major Activities including progress of the work to head office from time to tine interval. The head office will solve any major dispute which cannot be solved by the Contract Manager.", body_style))        
    elements.append(PageBreak())

    doc.build(elements, onFirstPage=add_technical_purposal_stamps, onLaterPages=add_technical_purposal_stamps)
    merge_with_letterhead(os.path.join(INPUT_DIR, LETTERHEAD_FILE), temp, final)




# =====================================================
# MOBILIZATION SCHEDULE PDF
# =====================================================

def create_mobilization_schedule_pdf(bid_number, contract_name, email_address, org_name, org_address):
    safe = bid_number.replace("/", "-")
    temp = os.path.join(OUTPUT_DIR, "temp_mobilization.pdf")
    final = os.path.join(OUTPUT_DIR, f"Mobilization_Schedule_{safe}.pdf")
    
    try:
        doc = SimpleDocTemplate(temp, pagesize=A4,
                                rightMargin=30, leftMargin=30,
                                topMargin=140, bottomMargin=18)
        elements = []

        elements.append(Spacer(1, 15))
        elements.append(Paragraph(f"Name of Project: <b>{contract_name}</b>", body_style))
        elements.append(Paragraph(f"Contract Identification No.: <b>{bid_number}</b>", body_style))
        elements.append(Paragraph(f"Proposed By: <b>{org_name}</b>", body_style))
        elements.append(Paragraph("MOBILIZATION SCHEDULE", title_style))
        elements.append(Spacer(1, 10))
                
        table_data = [
            ["", "", "1st week", "2nd week", "3rd week", "4th week"],
            ["S.N.", "Description of Work", "", "", "", ""],
            ["1", "Technical Supervision of site", "", "", "", ""],
            ["2", "Mobilization of Temporary Construction materials", "", "", "", ""],
            ["3", "Mobilization of labor / safety equipment and tools", "", "", "", ""],
            ["4", "Construction of temporary warehouse and labor camp/Temporary toilet", "", "", "", ""],
            ["5", "Mobilization of tools and equipment/material", "", "", "", ""]
        ]
        
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('BACKGROUND', (2, 2), (5, 2), colors.black),    
            ('BACKGROUND', (3, 3), (5, 6), colors.black),   
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ])
        
        table = Table(table_data, colWidths=[0.4*inch, 4.1*inch, 0.7*inch, 0.7*inch, 0.7*inch, 0.7*inch])
        table.setStyle(table_style)
        
        elements.append(table)
        
        elements.append(PageBreak())
        elements.append(Spacer(1, 10))
        elements.append(Paragraph("<b><u>CONSTRUCTION PLAN, TIME AND MOBILIZATION SCHEDULE</u></b>", title_style))
        elements.append(Spacer(1, 8))
        body_texts = [
              "<b><u><i>Work Plan and Methodology</i></u></b>",
              "Work methodology and plan has been set to the sequence of construction requirement, volume of construction and management of construction equipment, labor so as to meet line work progress within the stipulated time and in proper manner. Therefore the construction methodology has been prepared considering the priority and scope of work nature. Detailed bar chart scheduled is submitted in the submitted in the separate sheet showing the sequence and time of each activity of construction work.",
              "<b><u><i>Mobilization Schedule</i></u></b>",
              "After award and signing of contract, a detailed work, program/schedule will be submitted with performance guarantee. A technical team and pre-required construction tool/equipment will be deputed / mobilized at site. Drawings shall be read carefully and quantities of required construction materials will be worked out. A site office and a labor camp will be established within the construction site premises. A supervisor for labor management will be deputed at site. General requirements as per contract shall be fulfilled. During this period the site shall be prepared ready for immediate start of construction work.",
              "<b><u><i>Material Construction and transportation</i></u></b>",
              "Surveying, Designing if required execution of fields works, management of works preparation of bills, quality control, planning, scheduling and Progress review to asset requirements of construction materials equipment’s etc.",
              "<b><u><i>Storage and Handling of Materials</i></u></b>",
              "The construction material will be transported from its source and dumped at site at a secure location selected a storage area built within the vicinity of construction site. During the time of mass construction work, materials will be dumped at site in sufficient quantity, and materials will be continually supplied as per requirement of the work. Sensitive materials like cement and other shall be stored in secure place at site. Transportation of construction materials shall go thoroughly during construction period.",
              "<b><u><i>Equipment and Labor force management work</i></u></b>",
              "Sufficient numbers of equipment will be brought as site and made stand by for connected work. Sufficient numbers of work force like skilled and unskilled labor will be hired. Not all labors will be hired at a time. As per the requirement of the work, labor numbers can be increased or decreased.",
              "<b><u><i>Cash flow</i></u></b>",
              "There should be sufficient cash flow to carry out the work in a smooth way, for this purpose, the available credit resources shall be used. Credit from banks and market credits are main source of cash flow. Similarly billing work of work done quantities will be made time to time after completion of certain amount of work so that smooth run of construction work is maintained and no hindrances is occurred.",
              "<b><u><i>Safety and Quality Assurance</i></u></b>",
              "Special attention will be given for safety and quality of work. First and facilities shall be maintained at site. Safety wear to labors will be made available. Special training will be given to work force to be safe at site during construction. Regarding quality control approved materials will be used. Provision of Materials regular tests required will be done to maintain the quality of the work. only trained technicians will be employed ensure the quality construction works. Manufacturers' certificate of quality confirmation will be obtained from the Manufacturers' suppliers of construction materials and chemicals etc.",
              "<b><u><i>Proposed method of protection of habitant and other structures</i></u></b>",
              "Special care will be taken during construction to prevent any harm to other permanent or temporary structures and habitant. For this, labors and management will be given special instruction and training to prevent such things. Handling and storage of materials will be done properly as not to disturb others.",
              "<b><u><i>Overall Construction</i></u></b>",
              "Work shall be done as to complete within the construction time frame and to the satisfaction of the employer, engineer maintaining standard quality requirement. All the construction work will be done as per the specification, drawing and as per the instruction of Employer and engineer and the best of our knowledge and our workmanship.",
                         ]
        add_paragraphs(elements, body_texts, body_style)
        
        doc.build(elements, onFirstPage=add_mobilization_schedule_stamps, onLaterPages=add_technical_purposal_stamps)
        
        merge_with_letterhead(os.path.join(INPUT_DIR, LETTERHEAD_FILE), temp, final)
        print(f"✅ Mobilization Schedule PDF created: {os.path.basename(final)}")
    except Exception as e:
        print(f"❌ Failed to create Mobilization Schedule PDF: {e}")

# =====================================================
# LANDSCAPE ORG CHART WITH TEXT HEADER
# =====================================================

def draw_landscape_letterhead(canvas, doc):
    width, height = landscape(A4)

    bar_height = 70   # increased height
    top_margin = 10   # space from page top

    # Green header bar
    canvas.setFillColor(colors.darkgreen)
    canvas.rect(
        0,
        height - bar_height - top_margin,
        width,
        bar_height,
        stroke=0,
        fill=1
    )

    # Company Name
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 18)
    canvas.drawCentredString(
        width/2,
        height - 35 - top_margin,
        DEFAULT_ORG_NAME
    )

    # Address
    canvas.setFont("Helvetica", 10)
    canvas.drawCentredString(
        width/2,
        height - 52 - top_margin,
        DEFAULT_ORG_ADDRESS
    )

    # Contact line
    canvas.drawCentredString(
        width/2,
        height - 66 - top_margin,
        f"{DEFAULT_EMAIL} | 9843748114 | www.ecobuilders.com.np"
    )

    # Bottom separator line
    canvas.setStrokeColor(colors.darkgreen)
    canvas.setLineWidth(1.5)
    canvas.line(
        30,
        height - bar_height - top_margin - 5,
        width - 30,
        height - bar_height - top_margin - 5
    )


def create_org_chart_pdf(bid_number, contract_name,
                         org_name, org_address, email_address):

    safe = bid_number.replace("/", "-")
    final = os.path.join(OUTPUT_DIR, f"Organizational_Chart_{safe}.pdf")

    doc = SimpleDocTemplate(final, pagesize=landscape(A4),
                            rightMargin=10, leftMargin=10, topMargin=90, bottomMargin=10)
    
    elements = []
    styles = getSampleStyleSheet()

    elements.append(Spacer(1, 2))
    
    # Title style for organizational chart
    org_title_style = ParagraphStyle(
        'OrgTitleStyle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=12,
        alignment=TA_CENTER,
        spaceAfter=8,
        textColor=colors.darkblue
    )
    
    contract_style = ParagraphStyle(
        'ContractStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        alignment=TA_CENTER,
        spaceAfter=8
    )
    
    # Add contract information
    elements.append(Paragraph("<b>SITE ORGANIZATIONAL CHART</b>", title_style))
    elements.append(Spacer(1, 2))

    # Box style for org chart elements
    box_style = ParagraphStyle(
        'BoxStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=6,
        alignment=TA_CENTER,
        leading=7,
        textColor=colors.black
    )
    
    d = Drawing(11.7*inch, 6*inch)

    main_box_w, main_box_h = 1.5*inch, 0.35*inch
    small_box_w, small_box_h = 1.2*inch, 0.3*inch
    tiny_box_w, tiny_box_h = 1.0*inch, 0.25*inch

    def add_box(drawing, x, y, width, height, text, fill_color):
        drawing.add(Rect(x, y, width, height, fillColor=fill_color,
                     strokeColor=colors.black, strokeWidth=0.7))
        drawing.add(String(x + width/2, y + height/2, text,
                   fontName='Helvetica-Bold', fontSize=6,
                   textAnchor='middle', fillColor=colors.black))

    def add_arrow(drawing, x1, y1, x2, y2):
        drawing.add(Line(x1, y1, x2, y2, strokeColor=colors.black, strokeWidth=0.7))
        angle = math.atan2(y2 - y1, x2 - x1)
        arrow_size = 2.5
        drawing.add(Line(x2, y2,
                     x2 - arrow_size * math.cos(angle - math.pi / 6),
                     y2 - arrow_size * math.sin(angle - math.pi / 6),
                     strokeColor=colors.black, strokeWidth=0.7))
        drawing.add(Line(x2, y2,
                     x2 - arrow_size * math.cos(angle + math.pi / 6),
                     y2 - arrow_size * math.sin(angle + math.pi / 6),
                     strokeColor=colors.black, strokeWidth=0.7))

    # Level 1: Managing Director
    md_x, md_y = 5.5*inch, 5.5*inch
    add_box(d, md_x, md_y, main_box_w, main_box_h, "Managing Director", colors.lightgrey)

    # Level 2: Directors
    dir_positions = [
        (3.0*inch, 4.5*inch),
        (5.5*inch, 4.5*inch),
        (8.0*inch, 4.5*inch)
    ]

    for x, y in dir_positions:
        add_box(d, x, y, main_box_w, main_box_h, "Director", colors.whitesmoke)
        add_arrow(d, md_x + main_box_w/2, md_y, x + main_box_w/2, y + main_box_h)

    # Level 3: Project Manager
    pm_x, pm_y = 5.5*inch, 3.5*inch
    add_box(d, pm_x, pm_y, main_box_w, main_box_h, "Project/Construction Manager", colors.lightgrey)

    for x, y in dir_positions:
        add_arrow(d, x + main_box_w/2, y, pm_x + main_box_w/2, pm_y + main_box_h)

    # Level 4: Sections
    tech_x, tech_y = 0.5*inch, 2.5*inch
    account_x, account_y = 7.0*inch, 2.5*inch
    admin_x, admin_y = 9.0*inch, 2.5*inch

    add_box(d, tech_x, tech_y, main_box_w, main_box_h, "Technical Section", colors.lightgrey)
    add_box(d, account_x, account_y, main_box_w, main_box_h, "Account Section", colors.lightgrey)
    add_box(d, admin_x, admin_y, main_box_w, main_box_h, "Administrative Section", colors.lightgrey)

    add_arrow(d, pm_x + main_box_w/2, pm_y, account_x + main_box_w/2, account_y + main_box_h)
    add_arrow(d, pm_x + main_box_w/4, pm_y, tech_x + main_box_w/2, tech_y + main_box_h)
    add_arrow(d, pm_x + 3*main_box_w/4, pm_y, admin_x + main_box_w/2, admin_y + main_box_h)

    # Level 5: Technical Section Engineers
    tech_engineers = [
        (0.2*inch, 1.8*inch, "Civil Engineer"),
        (1.8*inch, 1.8*inch, "Architectural Engineer"),
        (3.4*inch, 1.8*inch, "Electrical Engineer"),
        (5.0*inch, 1.8*inch, "Structural Engineer")
    ]

    for x, y, title in tech_engineers:
        add_box(d, x, y, small_box_w, small_box_h, title, colors.whitesmoke)
        add_arrow(d, tech_x + main_box_w/2, tech_y, x + small_box_w/2, y + small_box_h)

    # Site Supervisor (under Civil Eng.)
    site_supervisor_x, site_supervisor_y = 0.2*inch, 1.3*inch
    add_box(d, site_supervisor_x, site_supervisor_y, tiny_box_w, tiny_box_h, "Site Supervisor", colors.whitesmoke)
    add_arrow(d, 0.2*inch + small_box_w/2, 1.8*inch, site_supervisor_x + tiny_box_w/2, site_supervisor_y + tiny_box_h)

    # Skilled & Unskilled Labors side-by-side under Site Supervisor
    skilled_x, skilled_y = 0.0*inch, 0.8*inch
    unskilled_x, unskilled_y = 1.2*inch, 0.8*inch  # shifted right by 1 inch

    add_box(d, skilled_x, skilled_y, tiny_box_w, tiny_box_h, "Skilled Labors", colors.whitesmoke)
    add_box(d, unskilled_x, unskilled_y, tiny_box_w, tiny_box_h, "Unskilled Labors", colors.whitesmoke)

    # Arrows from Site Supervisor to both boxes
    add_arrow(d, site_supervisor_x + tiny_box_w/2, site_supervisor_y, skilled_x + tiny_box_w/2, skilled_y + tiny_box_h)
    add_arrow(d, site_supervisor_x + tiny_box_w/2, site_supervisor_y, unskilled_x + tiny_box_w/2, unskilled_y + tiny_box_h)

    # Auto CAD (independent, under Technical Section)
    add_box(d, 1.8*inch, 1.3*inch, tiny_box_w, tiny_box_h, "Auto CAD", colors.whitesmoke)
    add_arrow(d, 1.8*inch + small_box_w/2, 1.8*inch, 1.8*inch + tiny_box_w/2, 1.3*inch + tiny_box_h)

    # Account Section Positions
    account_positions = [
        (6.3*inch, 1.8*inch, "Account Officer"),
        (7.7*inch, 1.8*inch, "Procurement Officer")
    ]

    for x, y, title in account_positions:
        add_box(d, x, y, small_box_w, small_box_h, title, colors.whitesmoke)
        add_arrow(d, account_x + main_box_w/2, account_y, x + small_box_w/2, y + small_box_h)

    # Administrative Section Positions
    admin_positions = [
        (9.2*inch, 1.8*inch, "Admin Officer"),
        (9.2*inch, 1.3*inch, "Store Keeper"),
        (9.2*inch, 0.8*inch, "Office Asst.")
    ]

    for i, (x, y, title) in enumerate(admin_positions):
        box_w = small_box_w if i == 0 else tiny_box_w
        box_h = small_box_h if i == 0 else tiny_box_h
        add_box(d, x, y, box_w, box_h, title, colors.whitesmoke)

        if i == 0:
            add_arrow(d, admin_x + main_box_w/2, admin_y, x + box_w/2, y + box_h)
        else:
            prev_y = admin_positions[i - 1][1]
            add_arrow(d, x + box_w/2, prev_y, x + box_w/2, y + box_h)
    elements.append(d)

    mode = get_landscape_letterhead()
    if mode == "PDF":
    # Build temp first
        temp = os.path.join(OUTPUT_DIR, "temp_org.pdf")
        doc_temp = SimpleDocTemplate(
            temp,
            pagesize=landscape(A4),
            rightMargin=10,
            leftMargin=10,
            topMargin=110,
            bottomMargin=0
            )
        doc_temp.build(elements,onFirstPage=add_org_chart_signature,onLaterPages=add_org_chart_signature)
    # merge external letterhead
        merge_with_letterhead(
            os.path.join(INPUT_DIR, LANDSCAPE_LETTERHEAD_FILE),
            temp,
            final
            )
    else:
    # use auto header
       doc.build(elements,onFirstPage=draw_landscape_with_signature,onLaterPages=draw_landscape_with_signature)
    
    print(f"✅ Created: {os.path.basename(final)}")



# =====================================================
# MAIN
# =====================================================
def main():

    print(f"\n📁 Checking directories...")
    print(f"   -> Please place your images/letterheads in: {os.path.abspath(INPUT_DIR)}")
    print(f"   -> Generated PDFs will be saved to: {os.path.abspath(OUTPUT_DIR)}\n")

    bid_number = input(f"IFB No. [{DEFAULT_BID_NUMBER}]: ") or DEFAULT_BID_NUMBER
    contract_name = input(f"Contract Name [{DEFAULT_CONTRACT_NAME}]: ") or DEFAULT_CONTRACT_NAME
    bid_date = input(f"Date [{DEFAULT_DATE}]: ") or DEFAULT_DATE
    bidder_name = input(f"Bidder Name [{DEFAULT_BIDDER_NAME}]: ") or DEFAULT_BIDDER_NAME
    employer_name = input(f"Employer [{DEFAULT_EMPLOYER_NAME}]: ") or DEFAULT_EMPLOYER_NAME
    employer_address = input(f"Address [{DEFAULT_EMPLOYER_ADDRESS}]: ") or DEFAULT_EMPLOYER_ADDRESS
    org_name = input(f"Organization [{DEFAULT_ORG_NAME}]: ") or DEFAULT_ORG_NAME
    org_address = input(f"Address [{DEFAULT_ORG_ADDRESS}]: ") or DEFAULT_ORG_ADDRESS
    email = input(f"Email [{DEFAULT_EMAIL}]: ") or DEFAULT_EMAIL


    create_technical_bid_pdf(bid_number, contract_name, bid_date,
                             bidder_name, org_name, org_address,
                             email, employer_name, employer_address)

    create_price_bid_pdf(bid_number, contract_name, bid_date,
                         bidder_name, org_name, org_address,
                         email, employer_name, employer_address)

    create_work_methodology_pdf(bid_number, contract_name, bid_date,
                                org_name, org_address, email)
    
    create_org_chart_pdf(bid_number, contract_name,
                         org_name, org_address, email)

    create_poa_with_declaration_pdf(
        bid_number,
        contract_name,
        bid_date,
        employer_name,
        employer_address,
        bidder_name,
        org_name,
        org_address,
        LETTERHEAD_FILE,
        f"POA_and_Declaration_{bid_number.replace('/', '-')}.pdf"
    )
    create_mobilization_schedule_pdf(bid_number, contract_name, email, org_name, org_address)
    print(f"\n📄 ALL PDFs CREATED SUCCESSFULLY AND SAVED TO: {os.path.abspath(OUTPUT_DIR)}\n")


if __name__ == "__main__":
    main()
