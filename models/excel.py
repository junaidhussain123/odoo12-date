from odoo import models
from datetime import datetime
from datetime import timedelta
import math


class DetailXlsxTemplate(models.AbstractModel):
    _name = 'report.sale_excel_report.detail_xlsx_template'
    _description = "COD XLSX Report"
    _inherit = 'report.report_xlsx.abstract'

    def generate_xlsx_report(self, workbook, report_detail, wizard_data):

        print(1111111111111111111, wizard_data)

        merge_format = workbook.add_format({
            'bold': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': '17',
            "font_color": 'black',
            'font_name': 'Metropolis',
        })

        merge_format_date = workbook.add_format({
            'bold': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': '15',
            "font_color": 'black',
            'font_name': 'Metropolis',
            'num_format': 'd mmm yyyy'
        })
        merge_format_date_1 = workbook.add_format({
            'bold': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': '10',
            "font_color": 'black',
            'font_name': 'Metropolis',
            'num_format': '%D%M%Y',
            "bg_color": '#ffe3cc'
        })

        format_form = workbook.add_format({
            "bold": 1,
            "border": 1,
            "align": 'center',
            "valign": 'vcenter',
            "font_color": 'black',
            "bg_color": '#ffe3cc',
            'font_size': '10',
            'font_name': 'Metropolis',
            'num_format': '#,##0',
        })

        format_form_1 = workbook.add_format({
            "bold": 1,
            "border": 1,
            "align": 'left',
            "valign": 'vcenter',
            "font_color": 'black',
            "bg_color": '#ffe3cc',
            'font_size': '15',
            'font_name': 'Metropolis',
        })
        format_form_clean = workbook.add_format({
            "bold": 1,
            "border": 1,
            "align": 'center',
            "valign": 'vcenter',
            "font_color": 'black',
            # "bg_color": '#F7A859',
            'font_size': '10',
            'font_name': 'Metropolis',
            'num_format': '#,##0',

        })
        format_form_no_lines = workbook.add_format({
            "bold": 1,
            "border": 1,
            "align": 'center',
            "valign": 'vcenter',
            "font_color": 'black',
            "bg_color": '#ffd7d7',
            'font_size': '10',
            'font_name': 'Metropolis',
        })
        merge_format_time = workbook.add_format({
            'bold': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': '15',
            "font_color": 'black',
            'font_name': 'Metropolis',
            'num_format': "%H:%M:%S"
        })
        merge_format_time = workbook.add_format({
            'bold': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': '10',
            "font_color": 'black',
            'font_name': 'Metropolis',
            'num_format': "%H:%M:%S",
            "bg_color": '#ffd7d7',
        })

        worksheet = workbook.add_worksheet('Report')
        head = ' Report'
        worksheet.set_row(3, 30)
        worksheet.set_row(4, 30)
        worksheet.merge_range('A1:G2', head, merge_format)
        worksheet.set_column('B:I', 18)
        worksheet.set_column('A:A', 26)

        rows = 3
        new_time= datetime.now()+timedelta(hours=5)
        worksheet.write_string(rows, 0, 'FCS day wise tender summary', format_form)
        worksheet.write_string(rows, 1, 'Printed Date', format_form)
        worksheet.write_string(rows, 2, str(datetime.now().date()), merge_format_time)
        worksheet.write_string(rows, 3, 'Report Time', format_form)
        worksheet.write_string(rows, 4, (new_time.strftime("%H:%M:%S")), merge_format_time)
        worksheet.write_string(rows, 5, 'Printed Time', format_form)
        worksheet.write_string(rows, 6, (new_time.strftime("%H:%M:%S")), merge_format_time)
        worksheet.write_string(rows, 7, '', format_form)
        worksheet.write_string(rows, 8, '', format_form)
        rows = 4
        worksheet.write_string(rows, 0, 'Store Name', format_form)
        worksheet.write_string(rows, 1, 'Date', format_form)
        worksheet.write_string(rows, 2, 'Cash', format_form)
        worksheet.write_string(rows, 3, 'Bank', format_form)
        worksheet.write_string(rows, 4, 'Net Sale', format_form)

        diff = wizard_data.date_to - wizard_data.date_from
        shops = self.env['pos.config'].search([])
        all_data = []
        total_cash_sum = 0.0
        print(total_cash_sum)
        total_bank_sum = 0.0
        total_netSale_sum = 0.0
        for shop in shops:

            all_dates_list = []
            shop_total_cash_amount = 0.0
            shop_total_bank_amount = 0.0
            shop_total_net_amount = 0.0
            for i in range(diff.days + 1):

                this_date_list = []
                this_date_total_cash_amount = 0.0
                this_date_total_amount = 0.0

                this_date_total_bank_amount = 0.0


                date = wizard_data.date_from + timedelta(i)
                date_start = date.replace(hour=00, minute=00, second=00)
                date_end = date.replace(hour=23, minute=59, second=59)
                orders = self.env['pos.order'].search(
                    [('create_date', '>=', date_start), ('create_date', '<=', date_end),('config_id','=',shop.id)])
                # new_orders = orders.search([('config_id','=',shop.id)])
                for new_order in orders:
                    this_date_total_amount+=new_order.amount_total
                    if new_order.statement_ids:
                        type =new_order.statement_ids[0].journal_id.type
                        if type=='cash':
                            this_date_total_cash_amount+=new_order.amount_total
                        if type=='bank':
                            this_date_total_bank_amount+=new_order.amount_total
                shop_total_cash_amount += this_date_total_cash_amount
                shop_total_bank_amount += this_date_total_bank_amount
                shop_total_net_amount += this_date_total_amount
                this_date_list.append(str(date_start.date()))
                this_date_list.append(this_date_total_cash_amount)
                this_date_list.append(this_date_total_bank_amount)
                this_date_list.append(this_date_total_amount)
                all_dates_list.append(this_date_list)
            all_dates_list.insert(0,shop.name)
            all_dates_list.insert(1,shop_total_cash_amount)
            all_dates_list.insert(2, shop_total_bank_amount)
            all_dates_list.insert(3, shop_total_net_amount)
            all_data.append(all_dates_list)
            total_cash_sum+=shop_total_cash_amount
            total_bank_sum+=shop_total_bank_amount
            total_netSale_sum +=shop_total_net_amount



        for data in all_data:
            rows += 1
            worksheet.write_string(rows, 1, str(data[0]), format_form)

            for dyy  in data:
                 if isinstance(dyy,list):
                     rows += 1
                     worksheet.write_string(rows, 1, str(dyy[0]), format_form)
                     worksheet.write_string(rows, 2, str(int(dyy[1])), format_form)
                     worksheet.write_string(rows, 3, str(int(dyy[2])), format_form)
                     worksheet.write_string(rows, 4, str(int(dyy[3])), format_form)


            rows += 1
            worksheet.write_string(rows, 0, str(data[0]), format_form)
            worksheet.write_string(rows, 1, 'Total:', format_form)

            worksheet.write_string(rows, 2, str(int(data[1])), format_form)
            worksheet.write_string(rows, 3, str(int(data[2])) ,format_form)
            worksheet.write_string(rows, 4, str(int(data[3])), format_form)
        #
            rows+=1


        worksheet.write_string(rows, 1, 'Subsidiary Total:', format_form)
        worksheet.write_string(rows, 2, str(int(total_cash_sum)), format_form)
        worksheet.write_string(rows, 3, str(int(total_bank_sum)), format_form)
        worksheet.write_string(rows, 4, str(int(total_netSale_sum)), format_form)

