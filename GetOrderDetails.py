import pandas as pd


class GetOrderDetails():
    """
    This class helps to read the csv files.

    It provides function get_data_from_excel(document_name) to read data from the sheet.
    """

    def __init__(self, excel_document_names):
        """
        Initializes all the required attributes for the class
        """
        self.resources = {}

        for document_name in excel_document_names:
            sheet_name  = document_name[:document_name.find('.')]
            if sheet_name == "orders":
                self.resources[sheet_name] = self.process_order_details(self.get_data_from_excel(document_name))
            elif sheet_name == "state":
                self.resources[sheet_name] = self.process_state_abbrevations(self.get_data_from_excel(document_name))
            else:
                self.resources[sheet_name] = self.get_data_from_excel(document_name)

    def get_encoded(self, ip_string):
        """
        It converts the unicode to string.

        Takes unicode string as input and returns utf-8 string as output.
        """
        if isinstance(ip_string, unicode):
            return ip_string.encode('utf-8')
        return str(ip_string)

    def process_order_details(self, orders):
        """
        It simplifies the order details given.

        It returns  a dictionary or format
        {
          'Name': {
                    {'OrderAttribut': 'Value'},
                    {item_details: [{'Lineitem name': 'Product Name',
                                     'Lineitem quantity': 'Quantities',
                                     'Lineitem price': price] }, {}, {}]}
                   }
        }
        """
        simplified_order_details = {}
        invoice_numbers = []
        for order_detail in orders:
            invoice_numbers.append(order_detail["Name"])
        invoice_numbers = set(invoice_numbers)
        for invoice_number in invoice_numbers:
            invoice_complete_order = {}
            for order in orders:
                if order["Name"] == invoice_number:
                    item_details = {}
                    item_details["Lineitem name"] = order["Lineitem name"]
                    item_details["Lineitem quantity"] = order["Lineitem quantity"]
                    item_details["Lineitem price"] = order["Lineitem price"]
                    if len(order["Created at"]) > 3:
                        if invoice_number in invoice_complete_order:
                            invoice_complete_order[invoice_number].update(order)
                        else:
                            invoice_complete_order[invoice_number] = order
                            invoice_complete_order[invoice_number]["item_details"] =  [item_details]
                    else:
                        if invoice_number in invoice_complete_order:
                            if "item_details" in invoice_complete_order[invoice_number]:
                                invoice_complete_order[invoice_number]["item_details"].append(item_details)
                            else:
                                invoice_complete_order[invoice_number]["item_details"] =  [item_details]
                        else:
                            invoice_complete_order[invoice_number]["item_details"] = [item_details]
            simplified_order_details[invoice_number] = invoice_complete_order
        return simplified_order_details

    def process_state_abbrevations(self, state_details):
        """It simplifies the state abbrevations details given."""
        states_abbrevations = {}
        for state_deatil in state_details:
            states_abbrevations[state_deatil["State Abbreviations"]] = state_deatil['State Name']
        return states_abbrevations

    def get_data_from_excel(self, document_name):
        """
        It takes an excel file name and reads the excel file.

        Returns a list of format.
        rows = [
                  {# row1
                     'key_as_column_title': 'value_of_cell',
                     'key_as_column_title': 'value_of_cell',
                     'key_as_column_title': 'value_of_cell'
                  },
                  {row2
                     'key_as_column_title': 'value_of_cell',
                     'key_as_column_title': 'value_of_cell',
                     'key_as_column_title': 'value_of_cell'
                  },
                  .
                  .
                  .
                  {row n
                     'key_as_column_title': 'value_of_cell',
                     'key_as_column_title': 'value_of_cell',
                     'key_as_column_title': 'value_of_cell'
                  }
               ]
        """
        sheet = pd.read_excel(document_name, sheet_name=0, header=None)
        nrows = len(sheet.loc[:][0])
        ncols = len(sheet.loc[0][:])
        keys = [sheet.loc[0, col_index] for col_index in xrange(ncols)]
        keys = [self.get_encoded(sheet.loc[0, col_index]) for col_index in xrange(ncols)]
        rows = []
        for row_index in xrange(1, nrows):
            row = {str(keys[col_index]): self.get_encoded((sheet.loc[row_index, col_index]))
                 for col_index in xrange(ncols)}
            rows.append(row)
        return rows