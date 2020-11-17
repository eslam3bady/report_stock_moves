# -*- coding: utf-8 -*-
from odoo import http

# class ReportStockMoves(http.Controller):
#     @http.route('/report_stock_moves/report_stock_moves/', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/report_stock_moves/report_stock_moves/objects/', auth='public')
#     def list(self, **kw):
#         return http.request.render('report_stock_moves.listing', {
#             'root': '/report_stock_moves/report_stock_moves',
#             'objects': http.request.env['report_stock_moves.report_stock_moves'].search([]),
#         })

#     @http.route('/report_stock_moves/report_stock_moves/objects/<model("report_stock_moves.report_stock_moves"):obj>/', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('report_stock_moves.object', {
#             'object': obj
#         })