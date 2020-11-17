# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import models, fields, api, _


class ProductVariantWizard(models.TransientModel):
    _name = 'stock.moves.report.wiz'
    _description = 'Stock Moves Report'

    date_from = fields.Date('Date From', )
    date_to = fields.Date('Date To', )
    compute_at_date = fields.Selection([
        (0, 'All'),
        (1, 'At a Specific Period')], default=0)
    report_type = fields.Selection(string="Report Type",
                                    selection=[('color', 'تقرير ارصده الاصناف وحركتها بالصور والالوان'),
                                               ('no_color', 'تقرير ارصده الاصناف وحركتها بدون صوره بالالوان')],
                                    required=True, default='color')

    branches = fields.Selection(string="", selection=[('all', 'All Company'), ('branch', 'By Branch'), ],
                                required=False, default='all')
    branch_ids = fields.Many2many(comodel_name="pos.config", relation="pos_config_id_moves_id", column1="config_id",
                                  column2="move_id", string="Branches", )
    categs = fields.Selection(string="", selection=[('all', 'All Categories'), ('categ', 'By Category'), ],
                              required=False, default='all')
    options = fields.Selection(string="", selection=[('image', 'With Images'), ('no_image', 'No Images'), ],
                               required=False, default='image')
    categ_ids = fields.Many2many(comodel_name="product.category", relation="categ_id_moves_id", column1="categ_id",
                                 column2="move_id", string="Categories", )
    vendor = fields.Selection(string="", selection=[('all', 'All Vendors'), ('vendor', 'By Vendor'), ],
                              required=False, default='all')
    vendor_ids = fields.Many2many(comodel_name="res.partner", relation="partner_id_moves_id", column1="partner_id",
                                  column2="moves_id", string="Vendors", domain=[('supplier', '=', True)])
    with_zero_values = fields.Selection(selection=[('all', 'All Values'), ('zero', 'No Zero Values')],
                                        required=False, default='all')
    sales_percent = fields.Selection(selection=[('all', 'All Sales'), ('percentage', 'Specific Percentage')],
                                     required=False, default='all')
    from_percent = fields.Float('Percentage From', default=0)
    to_percent = fields.Float('Percentage to', default=0)

    # def _print_report(self, data):
    #     return self.env.ref('sale_order_report.action_report_sales').report_action(self, data)

    def _print_report_xlsx(self, data):
        return self.env.ref('report_stock_moves.action_report_stock_moves_csv').report_action(self, data)

    # @api.multi
    # def view_report_pdf(self):
    #     self.ensure_one()
    #     data = {}
    #     data['ids'] = self.env.context.get('active_ids', [])
    #     data['model'] = self.env.context.get('active_model', 'ir.ui.menu')
    #     data['compute_at_date'] = self.compute_at_date
    #     data['sorting'] = self.sorting_type
    #     data['branches'] = self.branches
    #     data['categs'] = self.categs
    #     data['vendor'] = self.vendor
    #     data['branch_ids'] = self.branch_ids.ids
    #     data['categ_ids'] = self.categ_ids.ids
    #     data['vendor_ids'] = self.vendor_ids.ids
    #     if self.date_from and self.date_to:
    #         data['date_from'] = self.date_from
    #         data['date_to'] = self.date_to
    #     return self._print_report(data)

    @api.multi
    def view_report_xlsx(self):
        self.ensure_one()
        data = {}
        data['ids'] = self.env.context.get('active_ids', [])
        data['model'] = self.env.context.get('active_model', 'ir.ui.menu')
        data['report_type'] = self.report_type
        data['compute_at_date'] = self.compute_at_date
        data['branches'] = self.branches
        data['zero_values'] = self.with_zero_values
        data['sales_percent'] = self.sales_percent
        data['from_percent'] = self.from_percent
        data['to_percent'] = self.to_percent
        data['categs'] = self.categs
        data['options'] = self.options
        data['vendor'] = self.vendor
        data['branch_ids'] = self.branch_ids.ids
        data['categ_ids'] = self.categ_ids.ids
        data['vendor_ids'] = self.vendor_ids.ids
        if self.date_from and self.date_to:
            data['date_from'] = self.date_from
            data['date_to'] = self.date_to
        else:
            data['date_from'] = False
            data['date_to'] = False
        return self._print_report_xlsx(data)
