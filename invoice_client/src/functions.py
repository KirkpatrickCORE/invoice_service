import openpyxl

# Fucntion that inputs the data into the Excel template and uses
# the function excel_to_pdf to generate the invoice PDF:
def populate_excel_template(data, template_path, output_path):

    # Featch and open the Excel workbook & specify the appropriate
    # worksheet:
    wb = openpyxl.load_workbook(template_path)
    ws_invoice_p1 = wb['invoice_p1']

    # Set the JSON data:
    df = data

    # Input the data into the workbook invoice_p1:
    ws_invoice_p1['C6'] = df['billing_number']
    ws_invoice_p1['C7'] = df['invoice_number']
    ws_invoice_p1['K6'] = df['invoice_date']
    ws_invoice_p1['B10'] = df['client']
    ws_invoice_p1['B11'] = df['address_street']
    ws_invoice_p1['B12'] = df['address_line3']
    ws_invoice_p1['C16'] = df['customer_number']
    ws_invoice_p1['C17'] = df['prime_contract_number']
    ws_invoice_p1['C18'] = df['subcontractor_number']
    ws_invoice_p1['C20'] = df['project_number']
    ws_invoice_p1['C21'] = df['project_name']
    ws_invoice_p1['C22'] = df['project_pop']
    ws_invoice_p1['C23'] = df['terms']
    ws_invoice_p1['C24'] = df['effective_pay_date']
    ws_invoice_p1['C25'] = df['due_date']
    ws_invoice_p1['A28'] = df['job']
    ws_invoice_p1['A29'] = df['employee_name']
    ws_invoice_p1['I20'] = df['rate'] * df['hours']
    ws_invoice_p1['I19'] = (df['rate'] * df['hours']) / 83520
    ws_invoice_p1['H29'] = df['rate']
    ws_invoice_p1['I24'] = df['billing_period_from']
    ws_invoice_p1['I25'] = df['billing_period_to']
    ws_invoice_p1['I17'] = df['fee']

    wb.save(output_path)

