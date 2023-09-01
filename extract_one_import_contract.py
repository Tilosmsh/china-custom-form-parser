import fitz  # PyMuPDF
import re
import pandas as pd

def extract_info_from_pdf(pdf_path, num_pages):
    pdf_document = fitz.open(pdf_path)
    
    customs_number = ""
    contract_number = ""
    origin_country = ""
    import_date = ""
    currency = ""
    product_names = []
    total_amount = 0.0
    
    for page_num in range(1, num_pages + 1):
        current_page = pdf_document[page_num - 1]
        page_text = current_page.get_text("text")
        print(page_text)
        # Extract fixed information
        if page_num == 1:
            customs_number = re.search(r"海关编号：(\d+)", page_text).group(1)
            contract_number = re.search(r"合同协议号\n(.+)", page_text).group(1)
            origin_country = re.search(r"启运国（地区）\((\d+)\)\n(.+)", page_text).group(2)
            import_date = re.search(r"进口日期\n(\d+)", page_text).group(1)
        
        # Extract variable information
        product_matches = re.findall(r"\d+\n\d+\n([^(]+)\n\d+\w+\n\d+\n.+\n.+\n\d+\.\d+\n(\d+\.\d+)\n([A-Z]+)\n[^\n]+", page_text)
        for match in product_matches:
            product_names.append(match[0])
            total_amount += float(match[1])
            currency = match[2]
    
    pdf_document.close()
    
    # Return extracted information
    return {
        "CustomsNumber": customs_number,
        "ContractNumber": contract_number,
        "OriginCountry": origin_country,
        "ImportDate": import_date,
        "Currency": currency,
        "ProductNames": '，'.join(product_names[:2]),  # First two product names
        "TotalAmount": total_amount
    }





# # Example usage
# pdf_path = "./ImportPrint.pdf"
# num_pages = 4  # Number of pages to process

# info = extract_info_from_pdf(pdf_path, num_pages)
# print(info)
