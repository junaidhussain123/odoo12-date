# -*- coding: utf-8 -*-
from odoo import http

# class SaleExcelReport(http.Controller):
#     @http.route('/sale_excel_report/sale_excel_report/', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/sale_excel_report/sale_excel_report/objects/', auth='public')
#     def list(self, **kw):
#         return http.request.render('sale_excel_report.listing', {
#             'root': '/sale_excel_report/sale_excel_report',
#             'objects': http.request.env['sale_excel_report.sale_excel_report'].search([]),
#         })

#     @http.route('/sale_excel_report/sale_excel_report/objects/<model("sale_excel_report.sale_excel_report"):obj>/', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('sale_excel_report.object', {
#             'object': obj
#         })