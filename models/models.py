# -*- coding: utf-8 -*-

from odoo import models, fields, api
import datetime


class StockPicking(models.Model):
    _inherit = 'product.template'

    display_date = fields.Date(string="Display Date", required=False)