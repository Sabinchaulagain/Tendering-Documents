from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.graphics.shapes import Drawing, Rect, Line, String
import math
import os

# --- Constants ---
DEFAULT_BID_NUMBER = "JPM-NCB-14-2081/082"
DEFAULT_CONTRACT_NAME = "Buildings Construction Work of Satyavadi High School, Ja.Na.Pa. 11, Bajhang"
DEFAULT_DATE = "July 21, 2025"
DEFAULT_BIDDER_NAME = "Sabin Chaulagain"
DEFAULT_EMPLOYER_NAME = "Jayaprithivi Municipality, Bajhang"
DEFAULT_EMPLOYER_ADDRESS = "Bajhang, Nepal"
DEFAULT_LEAD_ORG = "Eco Builders & Engineers Pvt. Ltd."
DEFAULT_PARTNER_ORG = "Reshiva Construction Sewa Pvt. Ltd."
DEFAULT_LEAD_SHORT = "Eco"
DEFAULT_PARTNER_SHORT = "Reshiva"
DEFAULT_LEAD_ADDRESS = "Gokarneshwor-06, Kathmandu"
DEFAULT_PARTNER_ADDRESS = "Baneshwor-10, Nepal"
DEFAULT_EMAIL = "ecobuilders12@gmail.com"
DEFAULT_MD_LEAD = "Mr. Sabin Chaulagain"
DEFAULT_MD_PARTNER = "Mr. Sudharshan Chaulagain"
DEFAULT_LEAD_SHARE = 51
DEFAULT_PARTNER_SHARE = 49


styles = getSampleStyleSheet()

jv_name_style = ParagraphStyle(
    name='JVName',
    parent=styles['Normal'],
    fontName='Helvetica-Bold',
    fontSize=18,
    alignment=TA_CENTER,
    textColor=colors.darkblue,
    spaceAfter=4,
)

jv_contact_style = ParagraphStyle(
    name='JVContact',
    parent=styles['Normal'],
    fontName='Helvetica',
    fontSize=10,
    alignment=TA_CENTER,
    textColor=colors.darkgray,
    spaceAfter=2,
)

date_style = ParagraphStyle(
    name='DateRight',
    parent=styles['Normal'],
    fontName='Helvetica',
    fontSize=8,
    alignment=TA_RIGHT,
    spaceAfter=8,
)

contract_style = ParagraphStyle(
    name='RightContract',
    parent=styles['Normal'],
    fontName='Helvetica-Bold',
    fontSize=8,
    alignment=TA_RIGHT,
    spaceAfter=2,
)

body_style = ParagraphStyle(
    name='Body',
    parent=styles['Normal'],
    fontName='Helvetica',
    fontSize=8,
    alignment=TA_LEFT,
    leading=10,
    spaceAfter=2,
)

title_center_large = ParagraphStyle(
    name='TitleCenterLarge',
    parent=styles['Heading1'],
    fontName='Helvetica-Bold',
    fontSize=14,
    alignment=TA_CENTER,
    spaceAfter=6,
)

section_title_big = ParagraphStyle(
    name='SectionTitleBig',
    parent=styles['Heading1'],
    fontName='Helvetica-Bold',
    fontSize=14,
    alignment=TA_CENTER,
    spaceAfter=10,
)

justify_style = ParagraphStyle(
    name='Justify',
    alignment=TA_LEFT,
    fontName='Helvetica',
    fontSize=8,
    leading=10,
    spaceAfter=2,
)

center_style = ParagraphStyle(
    name='CenterTable',
    alignment=TA_CENTER,
    fontName='Helvetica',
    fontSize=8,
    leading=10,
    spaceAfter=2,
)

def add_mobilization_schedule_stamps(canvas, doc):
    canvas.saveState()
    try:
        img1 = ImageReader("lead_stamp.png")
        canvas.drawImage(img1, 
                        x=6*inch, 
                        y=6*inch, 
                        width=1.5*inch, 
                        height=0.75*inch, 
                        mask='auto',
                        preserveAspectRatio=True)
        
        img2 = ImageReader("lead_signature.png")
        canvas.drawImage(img2,
                        x=4.5*inch,
                        y=6*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
      
        img4 = ImageReader("partner_stamp.png")
        canvas.drawImage(img4,
                        x=5.2*inch,
                        y=6.1*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
    except Exception as e:
        # Fallback text if images not found
        canvas.setFont("Helvetica-Bold", 10)
        canvas.setFillColorRGB(1, 0, 0)
        
        print(f"Warning: Could not load stamp images - using text placeholders. Error: {e}")
    finally:
        canvas.restoreState()
        
def add_technical_purposal_stamps(canvas, doc):
    canvas.saveState()
    try:
        img1 = ImageReader("lead_stamp.png")
        canvas.drawImage(img1, 
                        x=6*inch, 
                        y=0.5*inch, 
                        width=1.5*inch, 
                        height=0.75*inch, 
                        mask='auto',
                        preserveAspectRatio=True)
        
        img2 = ImageReader("lead_signature.png")
        canvas.drawImage(img2,
                        x=4.5*inch,
                        y=0.5*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
      
        img4 = ImageReader("partner_stamp.png")
        canvas.drawImage(img4,
                        x=5.2*inch,
                        y=0.6*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
    except Exception as e:
        # Fallback text if images not found
        canvas.setFont("Helvetica-Bold", 10)
        canvas.setFillColorRGB(1, 0, 0)
        
        print(f"Warning: Could not load stamp images - using text placeholders. Error: {e}")
    finally:
        canvas.restoreState()
        
def add_jv_agreement_stamps(canvas, doc):
    canvas.saveState()
    try:
        img1 = ImageReader("lead_stamp.png")
        canvas.drawImage(img1, 
                        x=0.4*inch, 
                        y=3*inch, 
                        width=1.5*inch, 
                        height=0.75*inch, 
                        mask='auto',
                        preserveAspectRatio=True)
    
        img2 = ImageReader("lead_signature.png")
        canvas.drawImage(img2,
                        x=1.1*inch,
                        y=3*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
        img3 = ImageReader("partner_stamp.png")
        canvas.drawImage(img3,
                        x=4.4*inch,
                        y=3*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
        img4 = ImageReader("partner_signature.png")
        canvas.drawImage(img4,
                        x=5*inch,
                        y=3.1*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
    except Exception as e:
        canvas.setFont("Helvetica-Bold", 10)
        canvas.setFillColorRGB(1, 0, 0)
        
        print(f"Warning: Could not load stamp images - using text placeholders. Error: {e}")
    finally:
        canvas.restoreState()
        
        
def add_poa_stamps(canvas, doc):
    canvas.saveState()
    try:
        img1 = ImageReader("lead_stamp.png")
        canvas.drawImage(img1, 
                        x=0.4*inch, 
                        y=3.5*inch, 
                        width=1.5*inch, 
                        height=0.75*inch, 
                        mask='auto',
                        preserveAspectRatio=True)
    
        img2 = ImageReader("lead_signature.png")
        canvas.drawImage(img2,
                        x=1.1*inch,
                        y=3.5*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
        img3 = ImageReader("partner_stamp.png")
        canvas.drawImage(img3,
                        x=4.4*inch,
                        y=3.5*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
        img4 = ImageReader("partner_signature.png")
        canvas.drawImage(img4,
                        x=5*inch,
                        y=3.6*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
        img5 = ImageReader("lead_signature.png")
        canvas.drawImage(img5,
                        x=0.5*inch,
                        y=4.75*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)    
    except Exception as e:
        # Fallback text if images not found
        canvas.setFont("Helvetica-Bold", 10)
        canvas.setFillColorRGB(1, 0, 0)
        
        print(f"Warning: Could not load stamp images - using text placeholders. Error: {e}")
    finally:
        canvas.restoreState()
        
def add_bid_letter_stamp(canvas, doc):
    canvas.saveState()
    try:
        img1 = ImageReader("lead_stamp.png")
        canvas.drawImage(img1, 
                        x=1.3*inch, 
                        y=1.5*inch, 
                        width=1.5*inch, 
                        height=0.75*inch, 
                        mask='auto',
                        preserveAspectRatio=True)
        
        img2 = ImageReader("lead_signature.png")
        canvas.drawImage(img2,
                        x=0.5*inch,
                        y=1.5*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
      
        img4 = ImageReader("partner_stamp.png")
        canvas.drawImage(img4,
                        x=2.1*inch,
                        y=1.5*inch,
                        width=1.5*inch,
                        height=0.75*inch,
                        mask='auto',
                        preserveAspectRatio=True)
        
    except Exception as e:
        # Fallback text if images not found
        canvas.setFont("Helvetica-Bold", 10)
        canvas.setFillColorRGB(1, 0, 0)
        
        print(f"Warning: Could not load stamp images - using text placeholders. Error: {e}")
    finally:
        canvas.restoreState()
       

# --- Helper Functions ---
def build_letterhead(doc_width, jv_name, jv_address, email_address):
    header_data = [
        [Paragraph(f"<b>{jv_name.upper()}</b>", jv_name_style)],
        [Paragraph(jv_address, jv_contact_style)],
        [Paragraph(f"E-mail: <b><u>{email_address}</u></b>", jv_contact_style)],
        [HRFlowable(width="100%", thickness=1.5, color=colors.darkblue)]
    ]
    table = Table(header_data, colWidths=[doc_width])
    table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
    ]))
    return table

def add_paragraphs(elements, texts, style, spacing=8):
    """Helper to add multiple paragraphs with consistent spacing."""
    for text in texts:
        elements.append(Paragraph(text, style))
        elements.append(Spacer(1, spacing))

def validate_share(lead_share, partner_share):
    """Ensure the shares sum to 100%."""
    total = lead_share + partner_share
    if total != 100:
        raise ValueError(f"Shares must sum to 100% (current total: {total}%)")
        
        
        
def create_mobilization_schedule_pdf(bid_number, contract_name, jv_name, jv_address, email_address):
    safe_bid_number = bid_number.replace("/", "-")
    filename = f"Mobilization_Schedule_{safe_bid_number}.pdf"
    
    try:
        doc = SimpleDocTemplate(filename, pagesize=A4,
                                rightMargin=30, leftMargin=30,
                                topMargin=30, bottomMargin=18)
        elements = []

        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
        elements.append(Spacer(1, 15))
        elements.append(Paragraph(f"Name of Project: <b>{contract_name}</b>", justify_style))
        elements.append(Paragraph(f"Contract Identification No.: <b>{bid_number}</b>", justify_style))
        elements.append(Paragraph(f"Proposed By: <b>{jv_name}</b>", justify_style))
        elements.append(Paragraph("MOBILIZATION SCHEDULE", title_center_large))
        elements.append(Spacer(1, 10))
                
        # Table data
        table_data = [
            ["", "", "1st week", "2nd week", "3rd week", "4th week"],
            ["S.N.", "Description of Work", "", "", "", ""],
            ["1", "Technical Supervision of site", "", "", "", ""],
            ["2", "Mobilization of Temporary Construction materials", "", "", "", ""],
            ["3", "Mobilization of labor / safety equipment and tools", "", "", "", ""],
            ["4", "Construction of temporary warehouse and labor camp/Temporary toilet", "", "", "", ""],
            ["5", "Mobilization of tools and equipment/material", "", "", "", ""]
        ]
        
        # Table style
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
        
        # Create table with adjusted column widths
        table = Table(table_data, colWidths=[0.4*inch, 4.1*inch, 0.7*inch, 0.7*inch, 0.7*inch, 0.7*inch])
        table.setStyle(table_style)
        
        # Add table to elements
        elements.append(table)
        
        elements.append(PageBreak())
        
        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph("<b><u>CONSTRUCTION PLAN, TIME AND MOBILIZATION SCHEDULE</u></b>", title_center_large))
        elements.append(Spacer(1, 8))

        body_texts = [
              "<b><u><i>Work Plan and Methodology</i></u></b>",
              "Work methodology and plan has been set to the sequence of construction requirement, volume of construction and management of construction equipment, labor so as to meet line work progress within the stipulated time and in proper manner. Therefore the construction methodology has been prepared considering the priority and scope of work nature. Detailed bar chart scheduled is submitted in the submitted in the separate sheet showing the sequence and time of each activity of construction work.",
              "<b><u><i>Mobilization Schedule</i></u></b>",
              "After award and signing of contract, a detailed work, program/schedule will be submitted with performance guarantee. A technical team and pre-required construction tool/equipment will be deputed / mobilized at site. Drawings shall be read carefully and quantities of required construction materials will be worked out. A site office and a labor camp will be established within the construction site premises. A supervisor for labor management will be deputed at site. General requirements as per contract shall be fulfilled. During this period the site shall be prepared ready for immediate start of construction work.",
              "<b><u><i>Material Construction and transportation</i></u></b>",
              "Surveying, Designing if required execution of fields works, management of works preparation of bills, quality control, planning, scheduling and Progress review to asset requirements of construction materials equipment’s etc.",
              "<b><u><i>Material Construction and transportation</i></u></b>",
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
        
        # Build document
        doc.build(elements, onFirstPage=add_mobilization_schedule_stamps, onLaterPages=add_technical_purposal_stamps)
        print(f"✅ Mobilization Schedule PDF created: {filename}")
    except Exception as e:
        print(f"❌ Failed to create Mobilization Schedule PDF: {e}")
        
        
def create_work_methodology_pdf(bid_number, contract_name, bid_date, jv_name, jv_address, email_address, employer_name, employer_address):
    
    body_style = ParagraphStyle(
        name='Body',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=12,
        alignment=TA_LEFT,
        leading=14,
        spaceAfter=8,
    )
    
    title_center_large = ParagraphStyle(
        name='TitleCenterLarge',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=8,
    )

    
    safe_bid_number = bid_number.replace("/", "-")
    filename = f"Work_methodology_{safe_bid_number}.pdf"
    
    try:
        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=40)
        elements = []

        elements.append(Paragraph("PROPOSED WORK METHODOLOGY", title_center_large))
        elements.append(Paragraph(f"<u>Name of the Applicant : </u><b>{jv_name}</b>", title_center_large))
        elements.append(Spacer(1, 10))

        body_texts = [
            "<b>A. <u><i>Preliminary Site Organization Chart :</i></u></b>",
            "A Preliminary Site Organization Chart Showing the organization set up the considering sufficient in all respect to execute the work as per the specification within the condition of contract is attached herewith. All the personnel proposed for this project will be worked for whole contract period. A brief description of the relation of the staff within the site organization and its relation and business with head office is also detailed here under.",
            "<b>B. <u><i>Narrative Description of site organization Chart :</i></u></b>",
            "The Project Manager is the responsible for successful implementation smooth execution and timely completion of whole work. He will receive information from head office, give suggestion, progress report from administration, technical and account section as well as guidelines, special instruction, suggestion and take timely action by issuing notice, instruction making on the spot-study in a bid achieve timely execution maintaining. cordial relation and good co-operation, However, the company also look after the project as per requirement and visit time to time. The whole site organization under the project manager will be divided informally in two section i.e. Technical, Administrative and Financial section.",
            "<b>(i) <u><i>Technical Section :</i></u></b>",
            "This section shall comprise at least one Project Manager, one Experienced Civil Engineer, one Design Engineer, one Auto CAD operator, sufficient no’s of construction supervisor, mechanic operators etc. This section will sturdy drawing and specifications, take measurement of works, prepare running bills and submits the same to employer for payment on receipt of instructions from employer's representative and consultant regarding changes on include such changes in works. it will also monitor progress and take necessary action in time to keep pace with targeted schedule. This section will also look after machine requirement, its maintenance and running through mechanics and operators its job is also to prepare work schedule manpower schedule materials schedule, design drawing, survey, laboratory testing etc. However it can be more as the work is going to take in company.",
            "<b>(ii) <u><i>Administrative/Account Section :</i></u></b>",
            "This section will deal correspondences maintain records of staff and workmen, appoint personal, look after insurance, compensation, safety measures, attendance, records of leave, afford and maintain transportation facilities, housing, medical and sanitation facilities. it will also increase or decrease substitute staffs, workmen depending on manpower schedule keeping close touch with head office normally all information area received at site-office through administrative section and also this section is though not directly connected to construction-activities, indispensable for construction management's. As the perfect is concerned with directly to the community, this section will comprise 1 Administrative cum procurement manager, one accountant one store keeper, one assistant, guards and peons etc. This section will also provide manpower-material and finance-schedule consultation with technical schedule to the project manager.",
        ]
        add_paragraphs(elements, body_texts, body_style)
        elements.append(PageBreak())

        elements.append(Paragraph("PROPOSED WORK METHODOLOGY", title_center_large))
        elements.append(Paragraph(f"<u>Name of the Applicant : </u><b>{jv_name}</b>", title_center_large))
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
        add_paragraphs(elements, body_texts, body_style)
        elements.append(PageBreak())
        
        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
        elements.append(Spacer(1, 15))
        elements.append(Paragraph("<b>CONSTRUCTION PLAN AND METHEDOLOGY</b>", title_center_large))
        elements.append(Spacer(1, 10))

        body_texts = [
            f"<b>{contract_name}, Contract Identification No.:{bid_number}</b>",
            "The contractor upon signing of contract agreement will plan and stm1 to mobilize the necessary Equipment, manpower, and other logistics' within 15 days after signing of the contract. The Insurance cover will be provided from the start date to the end of Defects Liability period.",
            "Detail work program will be prepared and submitted to the Engineer for his approval after the Contract agreement is made.",
            "A site office will be established at a suitable place within the contract location. The site office, Labor camp as well as working area will be kept in a hygienic condition. Provision of toilets for Labor and employees will be made to avoid public nuisance and pollution of watercourses and Air. Depending up on the source of water, adequate drinking water and other water will be Supplied for the use of own staff and labor.",
            "The work shall be carried out in accordance with the relevant quality standards, test procedures or codes of practice as specific in the Technical Specification. Similarly, the method of working To be adopted will be such as to enable the satisfactory completion of the works before scheduled Date completion.",
            "The sequence of construction activities will be so managed that the double handling of materials Will be avoided.",
            "The source of construction material will get approved from the Engineer. Test sample of the Material and shop drawings if any require will be submitted early to allow sufficient time to the Engineer for review and approval.",
            "Safety and security of life and property is one of the very impol1ant part of the contract. For this/Fencing, guards, warning signs, safety helmet, adequate light for night work (if required and Permitted) etc. will be provided as required and to the satisfaction of the Engineer. Upon completion of the works the site will be cleaned and all temporary ot1ices, sanitary units, Stores, workshop, plant, tools, rubbish etc. will be removed from the side.",
        ]
        add_paragraphs(elements, body_texts, body_style)
        elements.append(PageBreak())
        
        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
        elements.append(Spacer(1, 15))
        elements.append(Paragraph("<b>WORK METHOD STATEMENT</b>", title_center_large))
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
        add_paragraphs(elements, body_texts, body_style)
        elements.append(Paragraph("<b><u>Relations Description of relationship between Head office & Site Office Management</u></b>", title_center_large))
        elements.append(Paragraph("The full authority will be given to the Contract Manager. The head office, as requested by the Contract Manager will furnish financial management. Site office will inform all its major Activities including progress of the work to head office from time to tine interval. The head office will solve any major dispute which cannot be solved by the Contract Manager.", body_style))        
        elements.append(PageBreak())

        doc.build(elements, onFirstPage=add_technical_purposal_stamps, onLaterPages=add_technical_purposal_stamps)
        print(f"✅ Work Methodology PDF created: {filename}")
    except Exception as e:
        print(f"❌ Failed to create Work Methodology PDF: {e}")



# --- PDF Generation Functions ---
def create_technical_bid_pdf(bid_number, contract_name, bid_date, bidder_name, jv_name, jv_address, email_address, employer_name, employer_address):
    safe_bid_number = bid_number.replace("/", "-")
    filename = f"Letter_of_Technical_Bid_{safe_bid_number}.pdf"
    
    try:
        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=40)
        elements = []

        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
        elements.append(Spacer(1, 15))
        elements.append(Paragraph(f"Date: {bid_date}", date_style))
        elements.append(Paragraph("LETTER OF TECHNICAL BID", title_center_large))
        elements.append(Paragraph(f"Name of the contract: <b>{contract_name}</b>", contract_style))
        elements.append(Paragraph(f"Invitation for Bid No.: <b>{bid_number}</b>", contract_style))
        elements.append(Spacer(1, 15))

        body_texts = [
            f"To,<br/>{employer_name},<br/>{employer_address}",
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
            "j) We agree to permit the Employer/DP or its representative to inspect our accounts and records and other documents relating to the bid submission and to have them audited by author appointed by the Employer.",
            "k) If our Bid is accepted we commit to mobilizing key equipment and personnel in accordance with the requirements set forth in section III (Evaluation and Qualification Criteria) and our technical proposal, or as otherwise agreed with the employer.",
            "l) We declare that we have not running contracts more than five (5) in accordance with ITB 4.9",
            f"Name: <b>{bidder_name}</b><br/>In the Capacity of Attorney Person<br/><br/><br/><br/><br/><br/><br/><br/>Signed...<br/>Duly authorized to sign the Bid for and on behalf of <b>{jv_name}</b><br/>Date: <b>{bid_date}</b>"
        ]
        add_paragraphs(elements, body_texts, body_style)

        doc.build(elements, onFirstPage=add_bid_letter_stamp)
        print(f"✅ Technical Bid PDF created: {filename}")
    except Exception as e:
        print(f"❌ Failed to create Technical Bid PDF: {e}")

def create_price_bid_pdf(bid_number, contract_name, bid_date, bidder_name, jv_name, jv_address, email_address, employer_name, employer_address):
    safe_bid_number = bid_number.replace("/", "-")
    filename = f"Letter_of_Price_Bid_{safe_bid_number}.pdf"
    
    try:
        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=40)
        elements = []

        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
        elements.append(Spacer(1, 15))
        elements.append(Paragraph(f"Date: {bid_date}", date_style))
        elements.append(Paragraph("LETTER OF PRICE BID", title_center_large))
        elements.append(Paragraph(f"Name of the contract: <b>{contract_name}</b>", contract_style))
        elements.append(Paragraph(f"Invitation for Bid No.: <b>{bid_number}</b>", contract_style))
        elements.append(Spacer(1, 15))

        body_texts = [
            f"To,<br/>{employer_name},<br/>{employer_address}",
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
            Paragraph(f"<b>Name of recipient</b><br/>................<br/>................", center_style),
            Paragraph(f"<b>Address</b><br/>.................<br/>.................", center_style),
            Paragraph(f"<b>Reason</b><br/>..................<br/>..................", center_style),
            Paragraph(f"<b>Amount</b><br/>..................<br/>..................", center_style),
        ]]
        elements.append(Table(price_table, colWidths=[1.7 * inch, 1.7 * inch, 1.7 * inch, 1.7 * inch]))
        elements.append(Spacer(1, 12))

        footer_texts = [        
            "h) We understand that this bid, together with your written acceptance thereof included in your notification of award, shall constitute a binding contract between us, until a formal contract is prepared and executed;",
            "i) We understand that you are not bound to accept the lowest evaluated bid or any other bid that you may receive and",
            "j) We declare that we are solely responsible for the authenticity of the documents submitted by us.",
            "k) We agree to permit the Employer/DP or its representative to inspect our accounts and records and other documents relating to the bid submission and to have them audited by author appointed by the Employer.",
            f"Name: <b>{bidder_name}</b><br/>In the Capacity of Attorney Person<br/><br/><br/><br/><br/><br/><br/><br/>Signed...<br/>Duly authorized to sign the Bid for and on behalf of <b>{jv_name}</b><br/>Date: <b>{bid_date}</b>"
        ]
        add_paragraphs(elements, footer_texts, body_style)

        doc.build(elements, onFirstPage=add_bid_letter_stamp)
        print(f"✅ Price Bid PDF created: {filename}")
    except Exception as e:
        print(f"❌ Failed to create Price Bid PDF: {e}")

def create_jv_agreement_pdf(bid_number, contract_name, bid_date, employer_name, employer_address,
                          lead_org, partner_org, lead_short, partner_short, lead_address, partner_address,
                          email_address, md1, md2, lead_share, partner_share, bidder_name):
    safe_bid_number = bid_number.replace("/", "-")
    filename = f"JV_Agreement_and_POA_{safe_bid_number}.pdf"
    
    try:
        validate_share(lead_share, partner_share)  # Validate share totals
        jv_name = f"{lead_short}-{partner_short} J.V."
        jv_address = lead_address

        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=40)
        elements = []
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("JOINT VENTURE AGREEMENT", section_title_big))
        elements.append(HRFlowable(width="100%", thickness=1, color=colors.black))
        elements.append(Spacer(1, 8))

        agreement_texts = [
            "<u>This agreement of Joint Venture into between:</u>",
            f"M/S <b>{lead_org}</b> shortly known as <b>{lead_short}</b> a Company registered in the government of Nepal with with department of industries having its registered office {lead_address} referred as First Partner or Lead Partner.",
            f"M/S <b>{partner_org}</b> shortly known as <b>{partner_short}</b> a Company registered in the government of Nepal with with department of industries having its registered office {partner_address} referred as Second Partner.",
            f"Whereas <b>{employer_name},{employer_address}</b>. Invite the Bid for the <b>{contract_name}</b>, IFB No.: <b>{bid_number}</b>. Experienced civil Contractors having experiences in Civil Construction field and are desirous and have agreed to apply jointly qualification, bidding, construction, execution and maintenance of the above mentioned work.",
            "All parties agree it upon follows:",
            f"The name of Joint Venture will be <b>{jv_name.upper()}</b>",
            f"2. The address of Joint Venture will be <b>{jv_address}</b>",
            f"3. The Lead Partner of this Joint venture is <b>{lead_org}</b>",
            "4. All the partners of the joint venture shall be jointly and severally liable for the execution of the contract. All the parties have agreed that a board consisting of representative from each party will be formed which will be responsible for overall management to the job order to fulfill all contractual obligations to complete the job within the stipulated period.",
            f"5. All the parties have agreed that the contribution of each party for this contract will be as below & the return from this contract and the loss will also be shared as per each party's contribution.<br></br><b>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp (a) {lead_org}&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp - {lead_share}% of total works,<br></br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp (b) {partner_org}&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp - {partner_share}% of total works</b>",
            "6. All the expenses involved in execution of contract shall be borne by each party in proportion to their participation ratio as explained on clause No. 5.",
            "7. Matters not stipulated in this agreement shall be decided between the parties mutually from time to time. Matters provided under this agreement or any of its terms and conditions may be amended.",
            f"8. All the companies agreed on the Terms and Condition Mentioned above and signed the agreement today on <b>{bid_date}</b>.",
            f"This agreement empowers Mr <b>{md1}, MD of {lead_org}</b> to sign all the Documents submits bids, receive instruction, negotiable and deal with employer on behalf of the Joint Venture."
            "Seal and Signature.<br/><br/<br/><br/><br/><br/><br/><br/><br/><br/>"
        ]
        add_paragraphs(elements, agreement_texts, justify_style, spacing=10)

        sigs = [[
            Paragraph(f"<b>{md1}</b><br/>MD<br/>{lead_org}", justify_style),
            Paragraph(f"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>{md2}</b><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;MD<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{partner_org}", justify_style)
        ]]
        elements.append(Table(sigs, colWidths=[3.5 * inch, 3.5 * inch]))
        elements.append(Spacer(1, 12))
        elements.append(PageBreak())

        # --- Power of Attorney ---
        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("POWER OF ATTORNEY", section_title_big))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"Date: {bid_date}", date_style))

        poa_texts = [
            f"<b>To,<br/>{employer_name},<br/>{employer_address}.</b>",
            f"&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<b><u>Subject: Power of Attorney</u></b> ",
            f"For: <b>{contract_name}</b> and Invitation for Bids No.: <b>{bid_number}</b>",
            f"Dear sir,<br/>Known all men by these presents that we the undersigned, All Authorized JV Partners lawfully authorized to represent and act on behalf of the said form under the company Act do hereby authorized <b>{bidder_name}, authorized representative of JV</b> whose specimen signature appears as given below to run all business activities signing joint venture/preparation/signing/providing/Qualification and bid, withdrawal and modification of bid, negotiable with the Employer, execute the contract and conduct all necessary dents/ agreements with all project, run all banking activities, to authorize any other person to represent on behalf of his authorization within Nepal and abroad for Contract mentioned above.",
            f"This undersigned shall acknowledge the Legal effects of the signature of the said attorney holder after the signing and sealing of power of attorney.",
            f"Specimen Signature of:<br/><br/><br/><br/><br/><br/><br/><b>(Mr. {bidder_name})</b><br/>Authorized Representative<br/>Seal and Signature:<br/><br/><br/><br/><br/><br/>"
        ]
        add_paragraphs(elements, poa_texts, justify_style, spacing=12)
        elements.append(Table(sigs, colWidths=[3.5 * inch, 3.5 * inch]))

        doc.build(elements, onFirstPage=add_jv_agreement_stamps, onLaterPages=add_poa_stamps)
        print(f"✅ JV Agreement & POA PDF created: {filename}")
    except ValueError as e:
        print(f"❌ Validation Error: {e}")
    except Exception as e:
        print(f"❌ Failed to create JV Agreement PDF: {e}")

# --- Organizational Chart Function ---
def create_org_chart_pdf(bid_number, contract_name, bid_date, jv_name, jv_address, email_address):
    safe_bid_number = bid_number.replace("/", "-")
    filename = f"Organizational_Chart_{safe_bid_number}.pdf"
    
    try:
        doc = SimpleDocTemplate(filename, pagesize=landscape(A4),
                                rightMargin=10, leftMargin=10, topMargin=10, bottomMargin=10)
        
        elements = []
        styles = getSampleStyleSheet()

        # Add letterhead
        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
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
        elements.append(Paragraph("SITE ORGANIZATIONAL CHART", org_title_style))
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
            (1.8*inch, 1.8*inch, "Electrical Engineer"),
            (3.4*inch, 1.8*inch, "Architecture Engineer"),
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
        doc.build(elements, onFirstPage=add_technical_purposal_stamps)
        print(f"✅ Organizational chart created: {filename}")
    except Exception as e:
        print(f"❌ Failed to create Organizational Chart PDF: {e}")

# --- Main Function ---
def main():
    print("\n📝 PDF Generator for Technical Bid, Price Bid, JV Agreement, and Organizational Chart\n")
    
    # Input Collection with Defaults
    bid_number = input(f"Enter Invitation for Bid No. [{DEFAULT_BID_NUMBER}]: ") or DEFAULT_BID_NUMBER
    contract_name = input(f"Enter Name of the Contract [{DEFAULT_CONTRACT_NAME}]: ") or DEFAULT_CONTRACT_NAME
    bid_date = input(f"Enter Date (e.g. July 17, 2025) [{DEFAULT_DATE}]: ") or DEFAULT_DATE
    bidder_name = input(f"Enter Bidder's Full Name [{DEFAULT_BIDDER_NAME}]: ") or DEFAULT_BIDDER_NAME
    employer_name = input(f"Enter Employer Name [{DEFAULT_EMPLOYER_NAME}]: ") or DEFAULT_EMPLOYER_NAME
    employer_address = input(f"Enter Employer Address [{DEFAULT_EMPLOYER_ADDRESS}]: ") or DEFAULT_EMPLOYER_ADDRESS

    lead_org = input(f"Enter Lead Organization Name [{DEFAULT_LEAD_ORG}]: ") or DEFAULT_LEAD_ORG
    partner_org = input(f"Enter Partner Organization Name [{DEFAULT_PARTNER_ORG}]: ") or DEFAULT_PARTNER_ORG
    lead_short = input(f"Short name for Lead Org (e.g., Eco) [{DEFAULT_LEAD_SHORT}]: ") or DEFAULT_LEAD_SHORT
    partner_short = input(f"Short name for Partner Org (e.g., Reshiva) [{DEFAULT_PARTNER_SHORT}]: ") or DEFAULT_PARTNER_SHORT
    lead_address = input(f"Enter Lead Organization Address [{DEFAULT_LEAD_ADDRESS}]: ") or DEFAULT_LEAD_ADDRESS
    partner_address = input(f"Enter Partner Organization Address [{DEFAULT_PARTNER_ADDRESS}]: ") or DEFAULT_PARTNER_ADDRESS
    email_address = input(f"Enter JV Email Address [{DEFAULT_EMAIL}]: ") or DEFAULT_EMAIL
    md1 = input(f"Name of MD from Lead Org [{DEFAULT_MD_LEAD}]: ") or DEFAULT_MD_LEAD
    md2 = input(f"Name of MD from Partner Org [{DEFAULT_MD_PARTNER}]: ") or DEFAULT_MD_PARTNER

    # Share Input with Validation Loop
    while True:
        try:
            lead_share = int(input(f"Lead Org Share % (e.g., 51) [{DEFAULT_LEAD_SHARE}]: ") or DEFAULT_LEAD_SHARE)
            partner_share = int(input(f"Partner Org Share % (e.g., 49) [{DEFAULT_PARTNER_SHARE}]: ") or DEFAULT_PARTNER_SHARE)
            validate_share(lead_share, partner_share)
            break
        except ValueError as e:
            print(f"⚠️ Error: {e}. Please try again.\n")

    jv_name = f"{lead_short}-{partner_short} J.V."
    jv_address = lead_address

    # Generate Bid Documents
    create_technical_bid_pdf(bid_number, contract_name, bid_date, bidder_name, jv_name, jv_address, email_address, employer_name, employer_address)
    create_price_bid_pdf(bid_number, contract_name, bid_date, bidder_name, jv_name, jv_address, email_address, employer_name, employer_address)
    create_jv_agreement_pdf(bid_number, contract_name, bid_date, employer_name, employer_address, lead_org, partner_org, lead_short, partner_short, lead_address, partner_address, email_address, md1, md2, lead_share, partner_share, bidder_name)
    create_work_methodology_pdf(bid_number, contract_name, bid_date, jv_name, jv_address, email_address, employer_name, employer_address)    
    create_mobilization_schedule_pdf(bid_number, contract_name, jv_name, jv_address, email_address)
    create_org_chart_pdf(
        bid_number=bid_number,
        contract_name=contract_name,
        bid_date=bid_date,
        jv_name=jv_name,
        jv_address=jv_address,
        email_address=email_address
    )
    
    print(f"\n📄 All PDFs saved to: {os.getcwd()}")

if __name__ == "__main__":
    main()
