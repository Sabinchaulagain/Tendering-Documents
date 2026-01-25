import os
import glob
import re
import pdfplumber
import pikepdf
def get_target_pdf():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return next((f for f in glob.glob(os.path.join(script_dir, "*.pdf")) 
                 if not os.path.basename(f).startswith("Extracted_")), None)
def extract_bid_data(pdf_path):
    results = {
        "Project Name": "Not Found",
        "Employer Address": "Not Found",
        "Bid No": "Not Found",
        "Bid Security Amount": "Not Found",
        "Avg. Annual Turnover": "Not Found",
        "Required Bid Capacity": "Not Found",
        "Cash-flow Requirement": "Not Found"
    }  
    pgs_to_extract = {"Personnel": set(), "Equipment": set(), "Experience": set()}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text: continue
            clean = " ".join(text.split())

            if results["Bid No"] == "Not Found":
                bid_m = re.search(r"(?:Bid No|NCB No|Contract Id No)\.?:?\s*([A-Z0-9/_\-]{10,50})", clean, re.IGNORECASE)
                if bid_m:
                    results["Bid No"] = bid_m.group(1).strip()

            if results["Project Name"] == "Not Found" and i < 5:
                proj_m = re.search(r"(?:PROCUREMENT OF|for)\s+((?:Construction|Building).*?)(?=National|Issued|Contract|Invitation|1 Page)", clean, re.IGNORECASE)
                if proj_m:
                    results["Project Name"] = proj_m.group(1).strip(", ")

            if results["Employer Address"] == "Not Found":
                addr_m = re.search(r"\[(Ministry of Education.*?Bhaktapur)\]|Employer's address is[:\s]+(.*?)(?=\.|\sTel|Telephone)", clean, re.IGNORECASE)
                if addr_m:
                    results["Employer Address"] = (addr_m.group(1) or addr_m.group(2)).strip()

            fin_map = {
                "Bid Security Amount": r"minimum of NRs\.?\s*([\d,.]+)\s*(million)?",
                "Avg. Annual Turnover": r"Annual Construction Turnover.*?NRs\.?\s*([\d,.]+)\s*(million)?",
                "Required Bid Capacity": r"bidding capacity.*?NRs\.?\s*([\d,.]+)\s*(million)?",
                "Cash-flow Requirement": r"cash-flow requirement.*?NRs\.?\s*([\d,.]+)\s*(million)?"
            }
            for key, pattern in fin_map.items():
                if results[key] == "Not Found":
                    m = re.search(pattern, clean, re.IGNORECASE)
                    if m:
                        suffix = f" {m.group(2)}" if m.group(2) else ""
                        results[key] = f"NRs. {m.group(1).rstrip('.')}{suffix}"

            if "Personnel Requirements" in clean and "Position" in clean: pgs_to_extract["Personnel"].add(i)
            if "Equipment Requirements" in clean and "Characteristics" in clean: pgs_to_extract["Equipment"].add(i)
            if "Form EXP" in clean or ("Experience" in clean and ("2.4.1" in clean or "2.4.2" in clean)): 
                pgs_to_extract["Experience"].add(i)
    return results, pgs_to_extract
target = get_target_pdf()
if target:
    data, pages = extract_bid_data(target)
    
    folder_name = "".join(x for x in data["Project Name"][:30] if x.isalnum() or x==' ')
    if not os.path.exists(folder_name): os.makedirs(folder_name)
    
    with open(os.path.join(folder_name, "Bid_Summary.txt"), "w", encoding="utf-8") as f:
        for k, v in data.items(): f.write(f"{k}: {v}\n")
    
    with pikepdf.open(target) as old_pdf:
        for section, indices in pages.items():
            if indices:
                new_pdf = pikepdf.Pdf.new()
                for idx in sorted(list(indices)): new_pdf.pages.append(old_pdf.pages[idx])
                new_pdf.save(os.path.join(folder_name, f"Extracted_{section}.pdf"))
                
    print(f"✅ Finished! Everything saved in folder: '{folder_name}'")
else:
    print("⚠️ No PDF found.")
