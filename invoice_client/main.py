from src.functions import populate_excel_template
from src.test_data import contract_data

def main():
    populate_excel_template(data = contract_data,
                            template_path = '/Users/micahkirkpatrick/kirkpatrickcore/kcore_company/finance/revenue/invoice/invoice_template/kcore_invoice_template.xlsx',
                            output_path = '/Users/micahkirkpatrick/kirkpatrickcore/kcore_company/finance/revenue/invoice/invoice_template/test_invoce.xlsx',
                            )


if __name__ == '__main__':
    main()
