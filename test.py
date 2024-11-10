import xml.etree.ElementTree as ET
import openpyxl


def parse_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    result = []

    for voucher in root.findall('.//VOUCHER'):
        voucher_data = {}
        voucher_data['Date'] = voucher.find('DATE').text if voucher.find('DATE') is not None else 'NA'
        voucher_data['Transaction Type'] = 'Parent'  # Adjust as needed for Parent/Child/Other logic
        voucher_data['Vch No.'] = voucher.find('VOUCHERNUMBER').text if voucher.find(
            'VOUCHERNUMBER') is not None else 'NA'
        voucher_data['Ref No.'] = 'NA' 
        voucher_data['Ref Type'] = 'NA'  
        voucher_data['Ref Date'] = 'NA'  
        voucher_data['Debtor'] = voucher.find('PARTYLEDGERNAME').text if voucher.find(
            'PARTYLEDGERNAME') is not None else 'NA'
        voucher_data['Ref Amount'] = 'NA' 
        voucher_data['Amount'] = voucher.find('AMOUNT').text if voucher.find('AMOUNT') is not None else 'NA'
        voucher_data['Particulars'] = voucher.find('NARRATION').text if voucher.find('NARRATION') is not None else 'NA'
        voucher_data['Vch Type'] = voucher.find('VCHTYPE').text if voucher.find('VCHTYPE') is not None else 'NA'
        voucher_data['Amount Verified'] = 'Yes'  

        # Debugging: print to check each voucher data
        print(voucher_data)

        result.append(voucher_data)

    return result


def write_to_excel(data, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    
    sheet.append(
        ["Date", "Transaction Type", "Vch No.", "Ref No.", "Ref Type", "Ref Date", "Debtor", "Ref Amount", "Amount",
         "Particulars", "Vch Type", "Amount Verified"])

    
    for row in data:
        sheet.append([row["Date"], row["Transaction Type"], row["Vch No."], row["Ref No."], row["Ref Type"],
                      row["Ref Date"], row["Debtor"], row["Ref Amount"], row["Amount"], row["Particulars"],
                      row["Vch Type"], row["Amount Verified"]])

    workbook.save(output_file)


if __name__ == "__main__":
    xml_file = 'input.xml'  
    output_file = 'Response_File.xlsx'

    parsed_data = parse_xml(xml_file)
    write_to_excel(parsed_data, output_file)
    print(f"Excel file '{output_file}' generated successfully.")
