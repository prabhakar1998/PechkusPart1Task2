import xlwt


class CreateInvoice():
    """This class is to create the invoice excel sheets."""
    def __init__(self, inflect_object, start_left_cell_column, workbook, order_details, sheetName='Invoice1'):
        self.start_column = start_left_cell_column
        self.workbook = workbook
        self.inflect_object = inflect_object
        # creating the table style scheme
        """
        It will set the styles for workbook.

        style1_right, style1_center, style1_left has background color LIGHT_GRAY
        with alignment of right, center and left.
        style2_right, style2_center, style2_left has background color DARK_GRAY
        fwith alignment of right, center and left.
        """
        xlwt.add_palette_colour("gray_ega", 0x21)
        self.workbook.set_colour_RGB(0x21, 235, 235, 224)
        self.sheetName = sheetName
        self.order_details = order_details
        self.order_items_count = len(order_details["item_details"])
        self.decleration = "We declare that this invoice shows the actual " \
                           "price of the goods described and all that "\
                           "perticulars are true and correct."
        self.sheetName = sheetName
        self.styleHeading = xlwt.easyxf("pattern: pattern solid, fore_color gray_ega;"
                                        "font: height 280, bold True, color black;"
                                        " align: horiz center")
        self.styleSubHeading = xlwt.easyxf("pattern: pattern solid, fore_color gray_ega;"
                                           " font: height 220, color black;"
                                           "align: horiz center")
        self.style1_right = xlwt.easyxf("pattern: pattern solid, fore_colour gray_ega;"
                                        "align: horiz right")
        self.style1_center = xlwt.easyxf("pattern: pattern solid, fore_colour gray_ega;"
                                         "align: horiz center")
        self.style1_left = xlwt.easyxf("pattern: pattern solid, fore_colour gray_ega;"
                                       "align: horiz left")
        self.style2_right = xlwt.easyxf("pattern: pattern solid, fore_colour "
                                        "gray25; align: horiz right")
        self.style2_center = xlwt.easyxf("pattern: pattern solid, fore_colour "
                                         "gray25; align: horiz center")
        self.style2_left = xlwt.easyxf("pattern: pattern solid, fore_colour "
                                       "gray25; align: horiz left")
        self.style_table = xlwt.easyxf("align: horiz center; borders: top_color black,"
                                       "bottom_color black, right_color black, "
                                       "left_color black, left thin, right thin, top "
                                       "thin, bottom thin;")
        self.style_table_initial = xlwt.easyxf("pattern: pattern solid, fore_colour white;"
                                               " align: horiz center; borders: top_color white,"
                                               "bottom_color white, right_color black, "
                                               "left_color black,left thin, right thin, "
                                               " top no_line, bottom no_line;")
        self.style_bottom_invoice = xlwt.easyxf("align: horiz center; borders:"
                                                "top_color black, bottom_color black, "
                                                "right_color black, left_color black,"
                                                " left thin, right thin, top thin, bottom thin;")

    def create_sheet(self):
        """Create a sheet in the workwbook."""
        self.worksheet = self.workbook.add_sheet(self.sheetName)
        # It enables cell overriding permissions.
        self.worksheet._cell_overwrite_ok = True

    def paint_template(self):
        """colouring the sheet and creating empty bill table."""
        for row in range(20):
            for col in range(self.start_column, self.start_column + 11):
                self.worksheet.write(row, col, "",
                                     xlwt.easyxf('pattern: pattern solid, fore_colour gray_ega;'))

        for col in range(self.start_column, self.start_column + 11):
            self.worksheet.write(4, col, "",
                                 xlwt.easyxf('pattern: pattern solid, fore_colour gray25;'))

        for row in range(12, 19):
            for col in range(self.start_column + 1, self.start_column + 10):
                self.worksheet.write(row, col, "", xlwt.easyxf('pattern: pattern solid, fore_colour white;'))

        # Colouring the bottom of invoice with white colour
        for row in range(21, self.order_items_count + 39):
            for col in range(self.start_column, self.start_column + 11):
                self.worksheet.write(row, col, "",
                                     xlwt.easyxf('pattern: pattern solid, fore_colour white;'))

        # Creating the bill table
        for row in range(22, self.order_items_count + 29):
            self.worksheet.write(row, self.start_column + 0, "",
                                 self.style_table_initial)
            for col in range(self.start_column + 5, self.start_column + 8):
                self.worksheet.write(row, col, "", self.style_table_initial)
            self.worksheet.write_merge(row, row, self.start_column + 9,
                                       self.start_column + 10, "",
                                       self.style_table_initial)

    def insert_invoice_details(self):
        """
        After the template is created this function inserts invoice details in the template.

        This takes no input and retuns nothing.
        """
        self.worksheet.write_merge(1, 1, self.start_column + 1,
                                   self.start_column + 9,
                                   "ENCELADUS INTERNET PRIVATE LIMITED",
                                   self.styleHeading)
        self.worksheet.write_merge(2, 2, self.start_column + 1,
                                   self.start_column + 9,
                                   "Address of Enceladus Internet Private Limited",
                                   self.styleSubHeading)

        self.worksheet.write_merge(4, 4, self.start_column + 1,
                                   self.start_column + 4,
                                   "GSTIN/UIN- GSTINXXAAWWKK00Z", self.style2_left)
        self.worksheet.write_merge(4, 4, self.start_column + 7,
                                   self.start_column + 9,
                                   "Email- hello@pechkus.co", self.style2_right)

        self.worksheet.write_merge(6, 6, self.start_column + 1,
                                   self.start_column + 2,
                                   "Invoice Number", self.style1_left)
        self.worksheet.write_merge(6, 6, self.start_column + 3,
                                   self.start_column + 7,
                                   "Mode/Terms of Payment", self.style1_center)
        self.worksheet.write_merge(6, 6, self.start_column + 8,
                                   self.start_column + 9,
                                   "Dated", self.style1_right)

        self.worksheet.write_merge(7, 7, self.start_column + 1,
                                   self.start_column + 2,
                                   self.order_details["Name"], self.style1_left)
        if "COD" in self.order_details["Payment Method"]:
            self.worksheet.write_merge(7, 7, self.start_column + 3,
                                       self.start_column + 7, "COD",
                                       self.style1_center)
        else:
            self.worksheet.write_merge(7, 7, self.start_column + 3,
                                       self.start_column + 7,
                                       "PPD - RAZORPAY", self.style1_center)
        self.worksheet.write_merge(7, 7, self.start_column + 8,
                                   self.start_column + 9,
                                   self.order_details["Created at"].split()[0],
                                   self.style1_right)
        self.worksheet.write_merge(9, 9, self.start_column + 1,
                                   self.start_column + 2,
                                   "Buyer's Order Number", self.style1_left)
        self.worksheet.write_merge(9, 9, self.start_column + 8,
                                   self.start_column + 9,
                                   "Despatched Through", self.style1_right)
        self.worksheet.write_merge(10, 10, self.start_column + 1,
                                   self.start_column + 2,
                                   self.order_details["Name"], self.style1_left)
        self.worksheet.write_merge(10, 10, self.start_column + 8, self.start_column + 9, "ECOM EXPRESS", self.style1_right)

        self.worksheet.write(13, self.start_column + 2,
                             self.order_details["Billing Name"],
                             xlwt.easyxf('pattern: pattern solid, fore_colour white; align: horiz left'))
        self.worksheet.write(14, self.start_column + 2,
                             self.order_details["Shipping Address1"],
                             xlwt.easyxf('pattern: pattern solid, fore_colour white; align: horiz left'))

        self.worksheet.write(15, self.start_column + 2,
                             "Mobile:" + self.order_details["Shipping Phone"],
                             xlwt.easyxf(" align: horiz left; pattern: pattern solid, fore_colour white"))
        self.worksheet.write(16, self.start_column + 2,
                             "State ",
                             xlwt.easyxf('pattern: pattern solid, fore_colour white; align: horiz left'))
        self.worksheet.write(16, self.start_column + 3,
                             self.order_details["StateName"],
                             xlwt.easyxf('pattern: pattern solid, fore_colour white; align: horiz left'))
        self.worksheet.write(17, self.start_column + 2,
                             "State Code",
                             xlwt.easyxf('pattern: pattern solid, fore_colour white; align: horiz left'))
        self.worksheet.write(17, self.start_column + 3,
                             self.order_details["Shipping Province"],
                             xlwt.easyxf('pattern: pattern solid, fore_colour white; align: horiz left'))

        # Adding bill table titles
        self.worksheet.write(22, self.start_column + 0, "S. No.", self.style_table)
        self.worksheet.write_merge(22, 22, self.start_column + 1,
                                   self.start_column + 4,
                                   "Description of Goods", self.style_table)
        self.worksheet.write(22, self.start_column + 5, "HSN/SAC", self.style_table)
        self.worksheet.write(22, self.start_column + 6, "Quantity", self.style_table)
        self.worksheet.write(22, self.start_column + 7, "Rate", self.style_table)
        self.worksheet.write(22, self.start_column + 8, "Per    ", self.style_table)
        self.worksheet.write(22, self.start_column + 9, "Amount", self.style_table)

        # adding items to the bill table
        count = 1
        for item in self.order_details["item_details"]:
            self.worksheet.write(22 + count, self.start_column + 0, count,
                                 self.style_table)
            self.worksheet.write_merge(22 + count, 22 + count,
                                       self.start_column + 1,
                                       self.start_column + 4,
                                       item['Lineitem name'], self.style_table)
            self.worksheet.write(22 + count, self.start_column + 5,
                                 "6101", self.style_table)
            self.worksheet.write(22 + count, self.start_column + 6,
                                 item["Lineitem quantity"], self.style_table)
            self.worksheet.write(22 + count, self.start_column + 7,
                                 item['Lineitem price'], self.style_table)
            self.worksheet.write(22 + count, self.start_column + 8,
                                 "Nos", self.style_table)
            self.worksheet.write(22 + count, self.start_column + 9,
                                 int(item["Lineitem quantity"]) * int(item['Lineitem price']),self.style_table)
            count += 1

        self.worksheet.write(self.order_items_count + 28, self.start_column + 8,
                             "IGST", self.style_table_initial)
        self.worksheet.write_merge(self.order_items_count + 28,
                                   self.order_items_count + 28,
                                   self.start_column + 9,
                                   self.start_column + 10,
                                   self.order_details["Taxes"],
                                   self.style_table_initial)

        self.worksheet.write(self.order_items_count + 29,
                             self.start_column + 0, "", self.style_table)
        self.worksheet.write_merge(self.order_items_count + 29,
                                   self.order_items_count + 29,
                                   self.start_column + 1,
                                   self.start_column + 4, "Total", self.style_table)
        self.worksheet.write(self.order_items_count + 29,
                             self.start_column + 5, "", self.style_table)
        self.worksheet.write(self.order_items_count + 29,
                             self.start_column + 6, "", self.style_table)
        self.worksheet.write(self.order_items_count + 29, self.start_column + 7,
                             "", self.style_table)
        self.worksheet.write(self.order_items_count + 29, self.start_column + 8,
                             "", self.style_table)
        bill_amount = int(self.order_details["Subtotal"]) + int(self.order_details["Taxes"])
        self.worksheet.write_merge(self.order_items_count + 29,
                                   self.order_items_count + 29,
                                   self.start_column + 9, self.start_column + 10,
                                   bill_amount, self.style_table)
        self.worksheet.write_merge(self.order_items_count + 31,
                                   self.order_items_count + 31,
                                   self.start_column + 0,
                                   self.start_column + 1,
                                   "Tax Amount", self.style_bottom_invoice)
        self.worksheet.write(self.order_items_count + 31,
                             self.start_column + 2, bill_amount,
                             self.style_bottom_invoice)
        self.worksheet.write_merge(self.order_items_count + 31,
                                   self.order_items_count + 31,
                                   self.start_column + 3,
                                   self.start_column + 10,
                                   "In words : INR " + self.inflect_object.number_to_words(bill_amount) + " ONLY",
                                   self.style_bottom_invoice)
        self.worksheet.write_merge(self.order_items_count + 33,
                                   self.order_items_count + 33,
                                   self.start_column + 0, self.start_column + 10,
                                   self.decleration, self.style_bottom_invoice)

        self.worksheet.write_merge(self.order_items_count + 35,
                                   self.order_items_count + 35, self.start_column + 6,
                                   self.start_column + 10, "For Enceladus Internet PVT LTD",
                                   self.style_table)
        self.worksheet.write_merge(self.order_items_count + 36,
                                   self.order_items_count + 37, self.start_column + 6,
                                   self.start_column + 10, "Authorised Signatory", self.style_table)
