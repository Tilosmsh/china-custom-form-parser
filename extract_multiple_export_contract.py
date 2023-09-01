import fitz  # PyMuPDF
import re
import pandas as pd

def extract_info_from_contract(pages):    
    customs_number = ""
    contract_number = ""
    origin_country = ""
    import_date = ""
    currency = ""
    product_names = []
    total_amount = 0.0
    
    for page_text in pages:

        # Extract fixed information
        customs_number = re.search(r"海关编号：(\d+)", page_text).group(1)
        contract_number = re.search(r"合同协议号\n(.+)", page_text).group(1)
        origin_country = re.search(r"运抵国（地区）\((\d+)\)\n(.+)", page_text).group(2)
        import_date = re.search(r"出口日期\n(\d+)", page_text).group(1)
        
        # Extract variable information
        product_matches = re.findall(r"\d+\n\d+\n([\D]+?)\n\d+\D+\n\d+\n.+\n.+\n^[1-9]\d*(\.\d+)?$\n(^[1-9]\d*(\.\d+)?$)\n([A-Z]+)\n[^\n]+", page_text, re.M)
        # print(product_matches)
        for match in product_matches:
            product_names.append(match[0].replace("\n", ""))
            total_amount += float(match[2])
            currency = match[4]
    
    
    # Return extracted information
    return {
        "关单编号": customs_number,
        "合同号": contract_number,
        "品名": '，'.join(product_names[:2]),  # First two product names
        "运抵国": origin_country,
        "出口日期": import_date,
        "币种": currency,
        "金额": total_amount
    }

def extract_contracts_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)
    num_pages = len(pdf_document)
    
    contracts = []
    current_contract = []
    current_contract_pages = 1
    
    for page_num in range(num_pages):
        current_page = pdf_document[page_num]
        page_text = current_page.get_text("text")
        # Identify new contracts based on "Page x of y" pattern
        page_pattern = re.findall(r"Page (\d+) of (\d+)", page_text)

        print(page_text)
        # page_pattern_example = [('1', '4')]
        if int(page_pattern[0][1]) == current_contract_pages:
            current_contract.append(page_text)
            contracts.append(current_contract)
            current_contract = []
            current_contract_pages = 1
        else:
            current_contract.append(page_text)
            current_contract_pages += 1
    
    # if current_contract_pages > 0:
    #     contracts.append(current_contract)
    
    # Process information for each contract
    contract_info_list = []
    for contract_pages in contracts:
        contract_info = extract_info_from_contract(contract_pages)
        print(contract_info)
        contract_info_list.append(contract_info)
    
    pdf_document.close()
    return contract_info_list

def save_to_excel(data, output_path):
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False)

def main(pdf_path, output_excel_path):
    contract_info_list = extract_contracts_from_pdf(pdf_path)
    save_to_excel(contract_info_list, output_excel_path)
    print(f"Excel file saved as {output_excel_path}")

if __name__ == "__main__":
    input_pdf_path = "ExportPrint.pdf"
    output_excel_path = "ExportPrint.xlsx"
    
    main(input_pdf_path, output_excel_path)

# # Example usage
# pdf_path = "./ImportPrint.pdf"

# contract_info_list = extract_contracts_from_pdf(pdf_path)
# for idx, contract_info in enumerate(contract_info_list, start=1):
#     print(f"Contract {idx} Info:")
#     print(contract_info)
#     print("-" * 30)
