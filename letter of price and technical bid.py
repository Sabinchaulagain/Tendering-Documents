from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

def create_wrapped_price_bid_letter(filename, bid_number, contract_name, bid_date, bidder_name, employer_name, employer_address):
    doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40,topMargin=40, bottomMargin=30)

    styles = getSampleStyleSheet()
    elements = []

    # Styles
    header_style = ParagraphStyle(name='Header', parent=styles['Heading2'],
                                  fontName='Helvetica-Bold', fontSize=14,
                                  alignment=TA_CENTER, spaceAfter=6)

    date_style = ParagraphStyle(name='DateRight', parent=styles['Normal'],
                                fontName='Helvetica', fontSize=12,
                                alignment=TA_RIGHT, spaceAfter=12)

    contract_style = ParagraphStyle(name='RightContract', parent=styles['Normal'],
                                    fontName='Helvetica', fontSize=12,
                                    alignment=TA_RIGHT, spaceAfter=10)

    body_style = ParagraphStyle(name='Body', parent=styles['Normal'],
                                fontName='Helvetica', fontSize=12,
                                alignment=TA_LEFT, spaceAfter=10)

    # Header
    elements.append(Paragraph("Eco - Reshiva J.V", header_style))
    elements.append(Paragraph("Gokarneshwor-06, Kathmandu", header_style))
    elements.append(Paragraph("E – Mail: ecobuilders12@gmail.com", header_style))

    # Date
    elements.append(Paragraph(f"Date: {bid_date}", date_style))

    # Title
    elements.append(Paragraph("<b>Letter of Price Bid</b>", header_style))
    elements.append(Spacer(1, 12))

    # Right-aligned contract info (inputs bolded)
    elements.append(Paragraph(f"Name of the contract: <b>{contract_name}</b>", contract_style))
    elements.append(Paragraph(f"Invitation for Bid No.: <b>{bid_number}</b>", contract_style))

    # Body
    body_paragraphs = [
        f"To,<br/>{employer_name},<br/>{employer_address}",
        "We, the undersigned, declare that:",
        "a) We have examined and have no reservation to the Bidding Documents, including Addenda issued in accordance with instructions to Bidder (ITB) Clause 8;",
        f"b) We offer to execute in conformity with the Bidding Documents the Following Works: <b>{contract_name}</b> and Invitation for Bid No.: <b>{bid_number}</b>",
        "c) The total price of our Bid, excluding any discount offered in item (d) below is AS PER BOQ or when left blank is the Bid Price indicated in the Bill of Quantities.",
        "d) The discount offered and the methodologies for their application are: None",
        "e) Our Bid shall be valid for a period of 150 days from the date fixed for the bid submission deadline in accordance with the bidding document and it shall remain binding upon us and may be accepted at any time before the expiration of that Period.",
        "f) If our bid is accepted, we commit to obtain a performance security in accordance with the Bidding Document;",
        "g) We have paid, or will pay the following commissions, gratuities, or fees with respect to the bidding process or execution of the Contract:<br/><br/>Name of recipient  Address  Reason  Amount<br/>………………………….  ………………  ………………….  ………………..<br/>………………………….  ………………  ………………….  ………………..",
        "h) We understand that this bid, together with your written acceptance thereof included in your notification of award, shall constitute a binding contract between us, until a formal contract is prepared and executed;",
        "i) We understand that you are not bound to accept the lowest evaluated bid or any other bid that you may receive;",
        "j) We declare that we are solely responsible for the authenticity of the documents submitted by us.",
        "k) We agree to permit the Employer/DP or its representative to inspect our accounts and records and other documents relating to the bid submission and to have them audited by author appointed by the Employer.",
        f"Name: <b>{bidder_name}</b><br/>In the Capacity of Attorney Person.<br/>Duly authorized to sign the Bid for and on behalf of Eco - Reshiva J.V.<br/>Date: {bid_date}"
    ]
    for para in body_paragraphs:
        elements.append(Paragraph(para, body_style))

    doc.build(elements)


def create_wrapped_technical_bid_letter(filename, bid_number, contract_name, bid_date, bidder_name, employer_name, employer_address):
    doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=30)
    
    styles = getSampleStyleSheet()
    elements = []

    # Styles
    header_style = ParagraphStyle(name='Header', parent=styles['Heading2'],
                                  fontName='Helvetica-Bold', fontSize=14,
                                  alignment=TA_CENTER, spaceAfter=6)

    date_style = ParagraphStyle(name='DateRight', parent=styles['Normal'],
                                fontName='Helvetica', fontSize=12,
                                alignment=TA_RIGHT, spaceAfter=12)

    contract_style = ParagraphStyle(name='RightContract', parent=styles['Normal'],
                                    fontName='Helvetica', fontSize=12,
                                    alignment=TA_RIGHT, spaceAfter=10)

    body_style = ParagraphStyle(name='Body', parent=styles['Normal'],
                                fontName='Helvetica', fontSize=12,
                                alignment=TA_LEFT, spaceAfter=10)

    # Header
    elements.append(Paragraph("Eco - Reshiva J.V", header_style))
    elements.append(Paragraph("Gokarneshwor-06, Kathmandu", header_style))
    elements.append(Paragraph("E – Mail: ecobuilders12@gmail.com", header_style))

    # Date
    elements.append(Paragraph(f"Date: {bid_date}", date_style))

    # Title
    elements.append(Paragraph("<b>Letter of Technical Bid</b>", header_style))
    elements.append(Spacer(1, 12))

    # Right-aligned contract info
    elements.append(Paragraph(f"Name of the contract: <b>{contract_name}</b>", contract_style))
    elements.append(Paragraph(f"Invitation for Bid No.: <b>{bid_number}</b>", contract_style))

    # Body content
    body_paragraphs = [
        f"To,<br/>{employer_name},<br/>{employer_address}",
        "We, the undersigned, declare that:",
        "a) We have examined and have no reservation to the Bidding Documents, including Addenda issued in accordance with instructions to Bidder (ITB) Clause 8;",
        f"b) We offer to execute in conformity with the Bidding Documents the Following Works:<br/><br/><b>{contract_name}</b> and Invitation for Bid No.: <b>{bid_number}</b>",
        "c) Our Bid consisting of the Technical and the Price Bid shall be valid for a period of 150 days from the date fixed for the bid submission deadline in accordance with the bidding document and it shall remain binding upon us and may be accepted at any time before the expiration of that Period.",
        "d) Our firm, including any subcontractors or suppliers for any part of the Contract, have nationalities from eligible countries in accordance with ITB 4.2 and meet the requirement of ITB 3.5 & 3.6;",
        "e) We are not participating as a Bidder or as a subcontractor in more than one bid in the bidding process in accordance with ITB 4.3(e), other than alternative offers submitted in accordance with ITB 13;",
        "f) Our firm, its affiliates or subsidiaries, including any Subcontractor or suppliers for any part of the contract has not been declared ineligible by DP, under the Employer's country laws or official regulations or by an act of compliance with a decision of the United Nations Security Council;",
        "g) We are not a government owned entity / we are a government owned entity but meet the Requirements of ITB 4.5;",
        "h) We declare that we, including any subcontractor or suppliers for any part of the contract, do not have any conflict of interest in accordance with ITB 4.3 and we have not been punished for an offense relating to the concerned profession or business.",
        "i) We declare that we are solely responsible for the authenticity of the documents submitted by us.",
        "j) We agree to permit the Employer/DP or its representative to inspect our accounts and records and other documents relating to the bid submission and to have them audited by an authority appointed by the Employer.",
        "k) If our Bid is accepted, we commit to mobilizing key equipment and personnel in accordance with the requirements set forth in section III (Evaluation and Qualification Criteria) and our technical proposal, or as otherwise agreed with the employer.",
        "l) We declare that we are not running more than five (5) contracts in accordance with ITB 4.9.",
        f"Name: <b>{bidder_name}</b><br/><br/>In the Capacity of Attorney Person.<br/>Duly authorized to sign the Bid for and on behalf of Eco - Reshiva J.V.<br/>Date: {bid_date}"
    ]

    for para in body_paragraphs:
        elements.append(Paragraph(para, body_style))

    doc.build(elements)
    


bid_number_input = input("Enter the Invitation for Bid No.: ")
contract_name_input = input("Enter the Name of the Contract: ")
date_input = input("Enter the Date: ")
bidder_name_input = input("Enter the Full Name of Bidder: ")
employer_name_input = input("Enter the Employer Name: ")
employer_address_input = input("Enter the Employer Address: ")



create_wrapped_technical_bid_letter(
    "letter of technical bid.pdf",
    bid_number_input,
    contract_name_input,
    date_input,
    bidder_name_input,
    employer_name_input,
    employer_address_input
)

create_wrapped_price_bid_letter(
    "letter of price bid.pdf",
    bid_number_input,
    contract_name_input,
    date_input,
    bidder_name_input,
    employer_name_input,
    employer_address_input
)


