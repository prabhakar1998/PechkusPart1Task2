import xlwt
import inflect
from GetOrderDetails import GetOrderDetails
from CreateInvoice import CreateInvoice

inflect_object = inflect.engine()
INPUT_FILE_NAMES = ["orders.xlsx", "GST.xlsx", "state.xlsx"]

orders = GetOrderDetails(INPUT_FILE_NAMES)
order_deatils = orders.resources["orders"]
invoice_numbers = order_deatils.keys()
invoice_count = len(invoice_numbers)
invoice_index = 0
workbook = xlwt.Workbook(encoding='ascii', style_compression=2)
xlwt.add_palette_colour("gray_ega", 0x21)
workbook.set_colour_RGB(0x21, 235, 235, 224)
while invoice_index < invoice_count:
    order = order_deatils[invoice_numbers[invoice_index]][invoice_numbers[invoice_index]]
    state_abbrevation = order['Shipping Province']
    order["StateName"] = orders.resources["state"][state_abbrevation]
    order["StateName"] = orders.resources["state"][state_abbrevation]
    inv1 = CreateInvoice(inflect_object, 0, workbook, order, "Sheeet%d" % ((invoice_index + 2)/2))
    inv1.create_sheet()
    inv1.paint_template()
    inv1.insert_invoice_details()
    invoice_index += 1
    if invoice_index >= invoice_count:
        break
    # second invoice
    order = order_deatils[invoice_numbers[invoice_index]][invoice_numbers[invoice_index]]
    state_abbrevation = order['Shipping Province']
    order["StateName"] = orders.resources["state"][state_abbrevation]
    order["StateName"] = orders.resources["state"][state_abbrevation]
    inv2 = CreateInvoice(inflect_object, 12, workbook, order, "Sheeet%d" % ((invoice_index + 1)/2))
    inv2.worksheet = inv1.worksheet
    inv2.paint_template()
    inv2.insert_invoice_details()
    invoice_index += 1
workbook.save("invoice.xlsx")
