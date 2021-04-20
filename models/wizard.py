from odoo import models, api, fields

from datetime import datetime



class DetailWizard(models.TransientModel):
    _name = "excel.wizard"
    _description = "Detail Wizard"

    date_from = fields.Datetime(string="From Date", required=True)
    date_to=fields.Datetime(string="To Date", required=True)


    def generate_detail_xlsx_rep(self):

        return self.env.ref('sale_excel_report.detail_sale_wizard_report_abc').report_action(self)

