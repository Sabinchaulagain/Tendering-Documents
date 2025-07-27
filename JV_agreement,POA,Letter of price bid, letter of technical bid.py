    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    import os
    
    styles = getSampleStyleSheet()

    jv_name_style = ParagraphStyle(
        name='JVName',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=18,  # bigger font size
        alignment=TA_CENTER,
        textColor=colors.darkblue,
        spaceAfter=4,
    )
    
    jv_contact_style = ParagraphStyle(
        name='JVContact',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,  # bigger than before
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
    

    def build_letterhead(doc_width, jv_name, jv_address, email_address):
        header_data = [
            [Paragraph(f"<b>{jv_name.upper()}</b>", jv_name_style)],
            [Spacer(1, 6)],  # <-- Added spacing between name and address
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
    
    
    def create_technical_bid_pdf(bid_number, contract_name, bid_date, bidder_name, jv_name, jv_address, email_address, employer_name, employer_address):
        safe_bid_number = bid_number.replace("/", "-")
        filename = f"Letter_of_Technical_Bid_{safe_bid_number}.pdf"
        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=40)
        elements = []
    
        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
        elements.append(Spacer(1, 15))
        elements.append(Paragraph(f"Date: {bid_date}", date_style))
        elements.append(Paragraph("LETTER OF TECHNICAL BID", title_center_large))
        elements.append(Paragraph(f"Name of the contract: <b>{contract_name}</b>", contract_style))
        elements.append(Paragraph(f"Invitation for Bid No.: <b>{bid_number}</b>", contract_style))
        elements.append(Spacer(1, 15))
    
        body = [
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
        for p in body:
            elements.append(Paragraph(p, body_style))
            elements.append(Spacer(1, 8))
    
        doc.build(elements)
        print(f"✅ Technical Bid PDF created: {filename}")
    
    def create_price_bid_pdf(bid_number, contract_name, bid_date, bidder_name, jv_name, jv_address, email_address, employer_name, employer_address):
        safe_bid_number = bid_number.replace("/", "-")
        filename = f"Letter_of_Price_Bid_{safe_bid_number}.pdf"
        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=40)
        elements = []
    
        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
        elements.append(Spacer(1, 15))
    
        elements.append(Paragraph(f"Date: {bid_date}", date_style))
        elements.append(Paragraph("LETTER OF PRICE BID", title_center_large))
        elements.append(Paragraph(f"Name of the contract: <b>{contract_name}</b>", contract_style))
        elements.append(Paragraph(f"Invitation for Bid No.: <b>{bid_number}</b>", contract_style))
        elements.append(Spacer(1, 15))
    
        body = [
            f"To,<br/>{employer_name},<br/>{employer_address}",
            "We, the undersigned, declare that:",
            "a)  We have examined and have no reservation to the Bidding Documents, including Addenda issued in according with instructions to Bidder (ITB) Clause 8;",
            f"b) We offer to execute in conformity with the Bidding Documents the Following Works:<br/><b>{contract_name}</b> and Invitation for Bid No.: <b>{bid_number}</b>",
            "c) The total price of our Bid, excluding any discount offered in item (d) below is <b>AS PER BOQ</b> or when left blank is the Bid Price indicated in the Bill of Quantities.",
            "d) The discount offered and the methodologies for their application are: <b> None</b>",
            "e) Our Bid shall be valid for a period of <b>150</b> days from the date fixed for the bid submission deadline in according with the bidding document and it shall remain binding upon us and may be accepted at any time before the expiration of that Period.",
            "f) If our bid is accepted, we commit to obtain a performance security in accordance with the Bidding Document:",
            "g) We have paid, or will pay the following commissions, gratuities, or fees with respect to the bidding process or execution of the Contract:",
        ]
        for p in body:
            elements.append(Paragraph(p, body_style))
            elements.append(Spacer(1, 8))
    
        sigs = [[Paragraph(f"<b>Name</b><br/>................", center_style),
                Paragraph(f"<b>Address</b><br/>.................", center_style),
                Paragraph(f"<b>Reason</b><br/>..................", center_style),
                Paragraph(f"<b>Amount</b><br/>..................", center_style),]]
        elements.append(Table(sigs, colWidths=[1.7 * inch, 1.7 * inch, 1.7 * inch, 1.7 * inch]))
        elements.append(Spacer(1, 12))
        body = [        
            "h) We understand that this bid, together with your written acceptance thereof included in your notification of award, shall constitute a binding contract between us, until a formal contract is prepared and executed;",
            "i) We understand that you are not bound to accept the lowest evaluated bid or any other bid that you may receive and",
            "j) We declare that we are solely responsible for the authenticity of the documents submitted by us.",
            "k) We agree to permit the Employer/DP or its representative to inspect our accounts and records and other documents relating to the bid submission and to have them audited by author appointed by the Employer.",
            f"Name: <b>{bidder_name}</b><br/>In the Capacity of Attorney Person<br/><br/><br/><br/><br/><br/><br/><br/>Signed...<br/>Duly authorized to sign the Bid for and on behalf of <b>{jv_name}</b><br/>Date: <b>{bid_date}</b>"
        ]
        for p in body:
            elements.append(Paragraph(p, body_style))
            elements.append(Spacer(1, 8))
    
        doc.build(elements)
        print(f"✅ Price Bid PDF created: {filename}")
    
    def create_jv_agreement_pdf(bid_number, contract_name, bid_date, employer_name, employer_address,
                                lead_org, partner_org, lead_short, partner_short, lead_address, partner_address,
                                email_address, md1, md2, lead_share, partner_share, bidder_name):
        jv_name = f"{lead_short}-{partner_short} J.V."
        jv_address = lead_address
        safe_bid_number = bid_number.replace("/", "-")
        filename = f"JV_Agreement_and_POA_{safe_bid_number}.pdf"
    
        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=40)
        elements = []


        elements.append(Paragraph("JOINT VENTURE AGREEMENT", section_title_big))
        elements.append(HRFlowable(width="100%", thickness=1, color=colors.black))
        elements.append(Spacer(1, 8))
    
        agreement_paragraphs = [
            "<u>This agreement of Joint Venture into between:</u>",
            f"M/S <b>{lead_org}</b> shortly known as <b>{lead_short}</b> a Company registered in the government of Nepal with with department of industries having its registered office {lead_address}  referred as First Partner or Lead Partner.",
            f"M/S <b>{partner_org}</b> shortly known as <b>{partner_short}</b> a Company registered in the government of Nepal with with department of industries having its registered office {partner_address} referred as Second Partner.",
            f"Whereas <b>{employer_name},{employer_address}</b>. Invite the Bid for the <b>{contract_name}</b>, IFB No.: <b>{bid_number}</b>. Experienced civil Contractors having experiences in Civil Constructionfield and are desirous and have agreed to apply jointly qualification, bidding, construction, execution and maintenance of the above mentioned work.",
            f"All parties agree it upon follows:",
            f"The name of Joint Venture will be <b>{jv_name.upper()}</b>",
            f"2. The address of Joint Venture will be <b>{jv_address}</b>",
            f"3. The Lead Partner of this Joint venture is <b>{lead_org}</b>",
            f"4. All the partners of the joint venture shall be jointly and severally liable for the execution of the contract. All the parties have agreed that a board consisting of representative from each party will be formed which will be responsible for overall management to the job order to fulfill all contractual obligations to complete the job within the stipulated period.",
            f"5. All the parties have agreed that the contribution of each party for this contract will be as below & the return from this contract and the loss will also be shared as per each party's contribution.<br></br><b>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp (a) {lead_org}&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp - {lead_share}% of total works,<br></br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp (b) {partner_org}&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp - {partner_share}% of total works</b>",
            "6. All the expenses involved in execution of contract shall be borne by each party in proportion to their participation ratio as explained on clause No. 5.",
            "7. Matters not stipulated in this agreement shall be decided between the parties mutually from time to time. Matters provided under this agreement or any of its terms and conditions may be amended.",
            f"8. All the companies agreed on the Terms and Condition Mentioned above and signed the agreement today on <b>{bid_date}</b>.",
            f"This agreement empowers Mr <b>{md1}, MD of {lead_org}</b> to sign all the Documents submits bids, receive instruction, negotiable and deal with employer on behalf of the Joint Venture."
            "Seal and Signature.<br/><br/<br/><br/><br/><br/><br/><br/><br/><br/>"
        ]
        for p in agreement_paragraphs:
            elements.append(Paragraph(p, justify_style))
            elements.append(Spacer(1, 10))
    
        sigs = [[Paragraph(f"<b>{md1}</b><br/>MD<br/>{lead_org}", justify_style),
                 Paragraph(f"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>{md2}</b><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;MD<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{partner_org}", justify_style)]]
        elements.append(Table(sigs, colWidths=[3.5 * inch, 3.5 * inch]))
        elements.append(Spacer(1, 12))
        elements.append(PageBreak())

        elements.append(build_letterhead(doc.width, jv_name, jv_address, email_address))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("POWER OF ATTORNEY", section_title_big))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"Date: {bid_date}", date_style))
        poa = [
            f"<b>To,<br/>{employer_name},<br/>{employer_address}.</b>",
            f"&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<b><u>Subject: Power of Attorney</u></b> ",
            f"For: <b>{contract_name}</b> and Invitation for Bids No.: <b>{bid_number}</b>",
            f"Dear sir,<br/>Known all men by these presents that we the undersigned, All Authorized JV Partners lawfully authorized to represent and act on behalf of the said form under the company Act do hereby authorized <b>{bidder_name}, authorized representative of JV</b> whose specimen signature appears as given below to run all business activities signing joint venture/preparation/signing/providing/Qualification and bid, withdrawal and modification of bid, negotiable with the Employer, execute the contract and conduct all necessary dents/ agreements with all project, run all banking activities, to authorize any other person to represent on behalf of his authorization within Nepal and abroad for Contract mentioned above.",
            f"This undersigned shall acknowledge the Legal effects of the signature of the said attorney holder after the signing and sealing of power of attorney.",
            f"Specimen Signature of:<br/><br/><br/><br/><br/><br/><br/><b>(Mr. {bidder_name})</b><br/>Authorized Representative<br/>Seal and Signature:<br/><br/><br/><br/><br/><br/>"
        ]
        for p in poa:
            elements.append(Paragraph(p, justify_style))
            elements.append(Spacer(1, 12))
        elements.append(Table(sigs, colWidths=[3.5 * inch, 3.5 * inch]))
    
        doc.build(elements)
        print(f"✅ JV Agreement & POA PDF created: {filename}")
    
    def main():
        bid_number = input("Enter Invitation for Bid No.: ") or "JPM-NCB-14-2081/082"
        contract_name = input("Enter Name of the Contract: ") or "Buildings Construction Work of satyavadi High School, Ja.Na.Pa. 11, Bajhang"
        bid_date = input("Enter Date (e.g. July 17, 2025): ") or "July 21, 2025"
        bidder_name = input("Enter Bidder's Full Name: ") or "Sabin Chaulagain"
        employer_name = input("Enter Employer Name: ") or "Jayaprithivi Municipality, Bajhang"
        employer_address = input("Enter Employer Address: ") or "Bajhang, Nepal"
    
        lead_org = input("Enter Lead Organization Name: ") or "Eco Builders & Engineers Pvt. Ltd."
        partner_org = input("Enter Partner Organization Name: ") or "Reshiva Construction Sewa Pvt. Ltd."
        lead_short = input("Short name for Lead Org (e.g., Eco): ") or "Eco"
        partner_short = input("Short name for Partner Org (e.g., Reshiva): ") or "Reshiva"
        lead_address = input("Enter Lead Organization Address: ") or "Gokarneshwor-06, Kathmandu"
        partner_address = input("Enter Partner Organization Address: ") or "Baneshwor-10, Nepal"
        email_address = input("Enter JV Email Address: ") or "ecobuilders12@email.com"
        md1 = input("Name of MD from Lead Org: ") or "Mr. Sabin Chaulagain"
        md2 = input("Name of MD from Partner Org: ") or "Mr. Sudharshan Chaulagain"
        lead_share = input("Lead Org Share % (e.g., 51): ") or "51"
        partner_share = input("Partner Org Share % (e.g., 49): ") or "49"
    
        jv_name = f"{lead_short}-{partner_short} J.V."
        jv_address = lead_address
    
        create_technical_bid_pdf(bid_number, contract_name, bid_date, bidder_name, jv_name, jv_address, email_address, employer_name, employer_address)
        create_price_bid_pdf(bid_number, contract_name, bid_date, bidder_name, jv_name, jv_address, email_address, employer_name, employer_address)
        create_jv_agreement_pdf(bid_number, contract_name, bid_date, employer_name, employer_address,
                                lead_org, partner_org, lead_short, partner_short,
                                lead_address, partner_address, email_address, md1, md2,
                                lead_share, partner_share, bidder_name)
    
        print(f"\n📄 All PDFs saved to: {os.getcwd()}")
    
    if __name__ == "__main__":
        main()
