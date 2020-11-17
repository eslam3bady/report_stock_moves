# -*- coding: utf-8 -*-
from odoo import models, fields, api, _
import datetime
from xlsxwriter.utility import xl_rowcol_to_cell
import io
from PIL import Image as PILImage
import base64
from odoo.exceptions import UserError
import collections


class ProductVariantReport(models.AbstractModel):
    _name = "report.stock_moves_report.stock_moves_report"
    _inherit = "report.report_xlsx.abstract"

    def get_stock_moves(self, date_from, date_to, branches, branch_ids, categs, categ_ids, vendor, zero_values,
                        vendor_ids, sales_percent, from_percent, to_percent):
        if date_from and date_to:
            domain = [('date_order', '>=', date_from), ('date_order', '<=', date_to)]
            moves_domain = [('date', '>=', date_from), ('date', '<=', date_to)]
        else:
            domain = []
            moves_domain = []
        if vendor == 'all':
            vendor_domain = []
        else:
            vendor_domain = [('id', 'in', vendor_ids)]
        vendor_domain.append(('supplier', '=', True))
        domain.append(('state', '=', 'done'))
        moves_domain.append(('state', '=', 'done'))
        vendors = self.env['res.partner'].search(vendor_domain)
        report_result = {}
        if categs == 'all':
            product_domain = []
        else:
            product_domain = [('categ_id', 'in', categ_ids)]
        product_domain.append(('available_in_pos', '=', True))
        products = self.env['product.template'].search(product_domain)
        if branches == 'all':
            pos_configs = self.env['pos.config'].search([])
        else:
            pos_configs = self.env['pos.config'].search([('id', 'in', branch_ids)])
            sessions = []
            pos_sessions = self.env['pos.session'].search([('config_id', 'in', branch_ids)])
            for session in pos_sessions:
                sessions.append(session.id)
            domain.append(('session_id', 'in', sessions))
        pos_orders = self.env['pos.order'].search(domain)
        moves = self.env['stock.move'].search(moves_domain)
        for vendor in vendors:
            total_incoming = 0
            total_sales = 0
            report_result[vendor.name] = [[]]
            for product in products:
                product_sales = 0
                product_incoming = 0
                product_vals = {}
                if product.variant_seller_ids and product.variant_seller_ids[0].name.id == vendor.id:
                    count = 0
                    if product.code_prefix:
                        code = product.code_prefix
                    else:
                        code = product.default_code
                    for attr in product.attribute_line_ids:
                        if attr.attribute_id.attr_type == "color":
                            count = len(attr.value_ids.ids)
                    if date_to and product.display_date:
                        days = (date_to - product.display_date).days
                    else:
                        if product.display_date:
                            days = (datetime.datetime.today().date() - product.display_date).days
                        else:
                            days = 0
                    product_vals[code] = {'id': product.id, 'Code': code, 'Colors': count, "Price": product.lst_price,
                                          "Display_date": product.display_date, "Days": days}
                    branch = {}
                    for config in pos_configs:
                        incoming = 0
                        sales = 0
                        for order in pos_orders:
                            if order.session_id.config_id.id == config.id:
                                for line in order.lines:
                                    if line.product_id.product_tmpl_id.id == product.id:
                                        sales += line.qty
                        for move in moves:
                            if move.location_dest_id.id == config.stock_location_id.id and move.product_id.product_tmpl_id.id == product.id:
                                incoming += move.product_uom_qty
                            if move.location_id.id == config.stock_location_id.id and move.product_id.product_tmpl_id.id == product.id:
                                incoming -= move.product_uom_qty
                        if incoming > 0:
                            first_percent = sales / incoming * 100
                        else:
                            first_percent = 0
                        if zero_values == 'zero':
                            if incoming or sales:
                                if sales_percent == 'percentage':
                                    if from_percent <= first_percent <= to_percent:
                                        branch[config.name] = {'incoming': incoming, 'sales': sales,
                                                               "balance": incoming - sales,
                                                               '1st_percent': round(first_percent, 2)}
                                else:
                                    branch[config.name] = {'incoming': incoming, 'sales': sales,
                                                           "balance": incoming - sales,
                                                           '1st_percent': round(first_percent, 2)}
                        else:
                            if sales_percent == 'percentage':
                                if from_percent <= first_percent <= to_percent:
                                    branch[config.name] = {'incoming': incoming, 'sales': sales,
                                                           "balance": incoming - sales,
                                                           '1st_percent': round(first_percent, 2)}
                            else:
                                branch[config.name] = {'incoming': incoming, 'sales': sales,
                                                       "balance": incoming - sales,
                                                       '1st_percent': round(first_percent, 2)}
                        total_incoming += incoming
                        product_incoming += incoming
                        # total_sales += sales
                        product_sales += sales
                    product_vals[code]['Branches'] = branch
                    product_vals[code]['Image'] = product.image_medium
                    product_vals[code]['sales'] = product_sales
                    product_vals[code]['incoming'] = product_incoming
                    report_result[vendor.name][0].append(product_vals)
            total_income = 0
            for purch in self.env['purchase.order'].search([('partner_id', '=', vendor.id)]):
                for line in purch.order_line:
                    total_income += line.product_uom_qty
            for pos in self.env['pos.order'].search([]):
                for line in pos.lines:
                    if line.product_id.product_tmpl_id.variant_seller_ids and \
                            line.product_id.product_tmpl_id.variant_seller_ids[0].name.id == vendor.id:
                        total_sales += line.qty
            report_result[vendor.name].append(total_income)
            report_result[vendor.name].append(total_sales)
            report_result[vendor.name].append(total_income - total_sales)
            # print(report_result[vendor.name])
        #     if not report_result[vendor.name][0]:
        #         # print(report_result[vendor.name][0])
        #         report_result.pop(vendor.name)
        # final_result = report_result
        # for item in report_result.keys():
        #     for key in report_result[item][0][0].keys():
        #         if not report_result[item][0][0][key]['Branches']:
        #             final_result[item][0][0].pop(key)
        #             # final_result[item] = report_result[item]
        #         print(report_result[0][0])
        #     if not report_result[item][0][0]:
        #         final_result.pop(item)
        # print(final_result)
                # print(report_result[item][0][0][key]['Branches'])
        return report_result

    def get_color_stock_moves(self, date_from, date_to, branches, branch_ids, categs, categ_ids, vendor, zero_values,
                              vendor_ids, sales_percent, from_percent, to_percent):
        if date_from and date_to:
            domain = [('date_order', '>=', date_from), ('date_order', '<=', date_to)]
            moves_domain = [('date', '>=', date_from), ('date', '<=', date_to)]
        else:
            domain = []
            moves_domain = []
        if vendor == 'all':
            vendor_domain = []
        else:
            vendor_domain = [('id', 'in', vendor_ids)]
        vendor_domain.append(('supplier', '=', True))
        domain.append(('state', '=', 'done'))
        moves_domain.append(('state', '=', 'done'))
        vendors = self.env['res.partner'].search(vendor_domain)
        report_result = {}
        if categs == 'all':
            product_domain = []
        else:
            product_domain = [('categ_id', 'in', categ_ids)]
        product_domain.append(('available_in_pos', '=', True))
        products = self.env['product.template'].search(product_domain)
        if branches == 'all':
            pos_configs = self.env['pos.config'].search([])
        else:
            pos_configs = self.env['pos.config'].search([('id', 'in', branch_ids)])
            sessions = []
            pos_sessions = self.env['pos.session'].search([('config_id', 'in', branch_ids)])
            for session in pos_sessions:
                sessions.append(session.id)
            domain.append(('session_id', 'in', sessions))
        pos_orders = self.env['pos.order'].search(domain)
        moves = self.env['stock.move'].search(moves_domain)
        prod_colors = {}
        for product in products:
            prod_colors[product.id] = []
            for attr in product.attribute_line_ids:
                if attr.attribute_id.attr_type == "color":
                    for value in attr.value_ids:
                        prod_colors[product.id].append(value.name)
        for vendor in vendors:
            total_incoming = 0
            total_sales = 0
            report_result[vendor.name] = [[]]
            for product in products:
                product_sales = 0
                product_incoming = 0
                product_vals = {}
                if product.variant_seller_ids and product.variant_seller_ids[0].name.id == vendor.id:
                    count = 0
                    if product.code_prefix:
                        code = product.code_prefix
                    else:
                        code = product.default_code
                    for attr in product.attribute_line_ids:
                        if attr.attribute_id.attr_type == "color":
                            count = len(attr.value_ids.ids)

                    if date_to and product.display_date:
                        days = (date_to - product.display_date).days
                    else:
                        if product.display_date:
                            days = (datetime.datetime.today().date() - product.display_date).days
                        else:
                            days = 0
                    product_vals[code] = {'id': product.id, 'Code': code, 'Colors': count, "Price": product.lst_price,
                                          "Display_date": product.display_date, "Days": days}
                    branch = {}
                    for color in prod_colors[product.id]:
                        incoming = 0
                        sales = 0
                        for order in pos_orders:
                            for line in order.lines:
                                if line.product_id.product_tmpl_id.id == product.id:
                                    for attr in line.product_id.attribute_value_ids:
                                        if attr.attribute_id.attr_type == 'color' and attr.name == color:
                                            sales += line.qty
                        for config in pos_configs:
                            for move in moves:
                                if move.location_dest_id.id == config.stock_location_id.id and move.product_id.product_tmpl_id.id == product.id:
                                    for attr in move.product_id.attribute_value_ids:
                                        if attr.attribute_id.attr_type == 'color' and attr.name == color:
                                            incoming += move.product_uom_qty
                            if move.location_id.id == config.stock_location_id.id and move.product_id.product_tmpl_id.id == product.id:
                                for attr in move.product_id.attribute_value_ids:
                                    if attr.attribute_id.attr_type == 'color' and attr.name == color:
                                        incoming -= move.product_uom_qty
                        if incoming > 0:
                            first_percent = sales / incoming * 100
                        else:
                            first_percent = 0
                        if zero_values == 'zero':
                            if incoming or sales:
                                if sales_percent == 'percentage':
                                    if from_percent <= first_percent <= to_percent:
                                        branch[color] = {'incoming': incoming, 'sales': sales,
                                                         "balance": incoming - sales,
                                                         '1st_percent': round(first_percent, 2)}
                                else:
                                    branch[color] = {'incoming': incoming, 'sales': sales, "balance": incoming - sales,
                                                     '1st_percent': round(first_percent, 2)}
                        else:
                            if sales_percent == 'percentage':
                                if from_percent <= first_percent <= to_percent:
                                    branch[color] = {'incoming': incoming, 'sales': sales, "balance": incoming - sales,
                                                     '1st_percent': round(first_percent, 2)}
                            else:
                                branch[color] = {'incoming': incoming, 'sales': sales, "balance": incoming - sales,
                                                 '1st_percent': round(first_percent, 2)}
                        total_incoming += incoming
                        product_incoming += incoming
                        total_sales += sales
                        product_sales += sales
                    product_vals[code]['Branches'] = branch
                    product_vals[code]['Image'] = product.image_medium
                    product_vals[code]['sales'] = product_sales
                    product_vals[code]['incoming'] = product_incoming
                    report_result[vendor.name][0].append(product_vals)
            total_income = 0
            for purch in self.env['purchase.order'].search([('partner_id', '=', vendor.id)]):
                for line in purch.order_line:
                    total_income += line.product_uom_qty
            report_result[vendor.name].append(total_income)
            report_result[vendor.name].append(total_sales)
            report_result[vendor.name].append(total_incoming - total_sales)
        return report_result

    def generate_xlsx_report(self, workbook, data, lines):
        report_lines = self.get_stock_moves(data['date_from'], data['date_to'], data['branches'], data['branch_ids'],
                                            data['categs'], data['categ_ids'], data['vendor'], data['zero_values'],
                                            data['vendor_ids'], data['sales_percent'], data['from_percent'],
                                            data['to_percent'])
        report_color_lines = self.get_color_stock_moves(data['date_from'], data['date_to'], data['branches'],
                                                        data['branch_ids'], data['categs'], data['categ_ids'],
                                                        data['vendor'], data['zero_values'], data['vendor_ids'],
                                                        data['sales_percent'], data['from_percent'], data['to_percent'])

        if report_lines:
            format_1 = workbook.add_format(
                {'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'green', 'font_color': 'white',
                 'border_color': 'white', 'border': 2})
            format_2 = workbook.add_format(
                {'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'black', 'font_color': 'white',
                 'border_color': 'white', 'border': 2})
            format_3 = workbook.add_format(
                {'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'black', 'font_color': 'white',
                 'border_color': 'white', 'border': 2})
            format_5 = workbook.add_format(
                {'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'white', 'font_color': 'black'})
            format_4 = workbook.add_format(
                {'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'black', 'font_color': 'white'})
            date_format = workbook.add_format(
                {'num_format': 'd mmmm yyyy', 'fg_color': 'black', 'valign': 'vcenter', 'font_color': 'white',
                 'border_color': 'white', 'border': 2})
            center = workbook.add_format(
                {'font_size': 14, 'align': 'center', 'fg_color': 'orange', 'valign': 'vcenter', 'font_color': 'white',
                 'border_color': 'white', 'border': 2})
            sheet = workbook.add_worksheet('Stocks Move Report')
            k = 0
            j = 0
            for categ in report_lines.items():
                if categ[1][0]:
                    aa = 0
                    sheet.merge_range(k, 0, k + 6, 1, categ[0], center)
                    k += 7
                    sheet.merge_range(k, 0, k + 1, 1, 'اجمالي الوارد', format_1)
                    k += 2
                    sheet.merge_range(k, 0, k + 1, 1, categ[1][1], format_2)
                    k += 2
                    sheet.merge_range(k, 0, k + 1, 1, 'اجمالي المبيعات', format_1)
                    k += 2
                    sheet.merge_range(k, 0, k + 1, 1, categ[1][2], format_2)
                    k += 2
                    sheet.merge_range(k, 0, k + 1, 1, 'اجمالي المتبقي', format_1)
                    k += 2
                    sheet.merge_range(k, 0, k + 1, 1, categ[1][3], format_2)
                    i = 0
                    for product in categ[1][0]:
                        for item in product.values():
                            if not item['Branches']:
                                continue
                            i += 2
                            v = j
                            if item['Days']:
                                sec_percentage = item['sales'] / item['Days']
                            else:
                                sec_percentage = 0
                            if data['options'] == 'image':
                                sheet.merge_range(v, i, v + 16, i + 2, '', format_3)
                                if item['Image']:
                                    bytes_data = base64.b64decode(item['Image'])
                                    data_img = io.BytesIO(bytes_data)
                                    x_scale = 1.5
                                    y_scale = 2.45
                                    sheet.insert_image(xl_rowcol_to_cell(v, i), 'data_img',
                                                       {'image_data': data_img, 'x_scale': x_scale, 'y_scale': y_scale})
                                i += 3
                            sheet.write(j, i, 'العرض', format_3)
                            if item['Display_date']:
                                sheet.merge_range(j, i + 1, j, i + 3, item['Display_date'], date_format)
                            else:
                                sheet.merge_range(j, i + 1, j, i + 3, item['Display_date'], format_3)

                            sheet.write(j, i + 4, 'السعر', format_3)
                            sheet.write(j, i + 5, item['Price'], format_3)
                            sheet.merge_range(j + 1, i, j + 2, i, 'عدد الايام', format_3)
                            sheet.merge_range(j + 1, i + 1, j + 2, i + 1, item['Days'], format_3)
                            sheet.merge_range(j + 3, i, j + 4, i, 'متوسط', format_3)
                            sheet.merge_range(j + 3, i + 1, j + 4, i + 1, sec_percentage, format_3)
                            sheet.merge_range(j + 1, i + 2, j + 2, i + 3, 'الوان', format_4)
                            sheet.merge_range(j + 3, i + 2, j + 4, i + 3, item['Colors'], format_4)
                            sheet.merge_range(j + 1, i + 4, j + 6, i + 5, item['Code'], format_1)
                            sheet.merge_range(j + 5, i, j + 6, i, 'الفرع', center)
                            sheet.merge_range(j + 5, i + 1, j + 6, i + 1, 'وارد', center)
                            sheet.merge_range(j + 5, i + 2, j + 6, i + 2, 'مبيعات', center)
                            sheet.merge_range(j + 5, i + 3, j + 6, i + 3, 'متبقي', center)
                            l = j + 7
                            aa = 9 - len(item["Branches"])
                            for line in item['Branches'].items():
                                sheet.write(l, i, line[0], format_3)
                                sheet.write(l, i + 1, line[1]['incoming'], format_3)
                                sheet.write(l, i + 2, line[1]['sales'], format_3)
                                sheet.write(l, i + 3, line[1]['balance'], format_3)
                                if line[1]['incoming'] > 0:
                                    sheet.write(l, i + 4,
                                                str(round(line[1]['sales'] / line[1]['incoming'] * 100, 2)) + '%',
                                                format_3)
                                else:
                                    sheet.write(l, i + 4, str(0) + '%', format_3)
                                if item['sales'] > 0:
                                    sheet.write(l, i + 5, str(round(line[1]['sales'] / item['sales'] * 100, 2)) + '%',
                                                format_3)
                                else:
                                    sheet.write(l, i + 5, str(0) + '%', format_3)
                                l += 1
                            if aa > 0:
                                sheet.merge_range(l, i, l + aa - 1, i + 5, "", format_4)
                                l += aa
                            sheet.write(l, i, 'الشركة', format_5)
                            sheet.write(l, i + 1, item['incoming'], format_5)
                            sheet.write(l, i + 2, item['sales'], format_5)
                            sheet.write(l, i + 3, item['incoming'] - item['sales'], format_5)
                            if item['incoming'] > 0:
                                sheet.write(l, i + 4, str(round(item['sales'] / item['incoming'] * 100, 2)) + '%',
                                            format_5)
                            else:
                                sheet.write(l, i + 4, str(0) + '%', format_5)
                            sheet.write(l, i + 5, '', format_5)
                            if data['options'] == 'image':
                                sheet.merge_range(l + 1, i - 3, l + 2, i + 3, "", format_1)
                                sheet.merge_range(l + 1, i + 4, l + 2, i + 5, "ملاحظات ادارية", center)
                            else:
                                sheet.merge_range(l + 1, i, l + 2, i + 3, "", format_1)
                                sheet.merge_range(l + 1, i + 4, l + 2, i + 5, "ملاحظات ادارية", center)
                            i += 5
                    if aa < 0:
                        sheet.merge_range(k + 2, 0, k - aa - 1, 1, '', format_2)
                        k -= aa
                    k += 3
                    j = k
        if report_color_lines:
            format_1 = workbook.add_format(
                {'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'green', 'font_color': 'white',
                 'border_color': 'white', 'border': 2})
            format_2 = workbook.add_format(
                {'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'black', 'font_color': 'white',
                 'border_color': 'white', 'border': 2})
            format_3 = workbook.add_format(
                {'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'black', 'font_color': 'white',
                 'border_color': 'white', 'border': 2})
            format_5 = workbook.add_format(
                {'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'white', 'font_color': 'black'})
            format_4 = workbook.add_format(
                {'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'fg_color': 'black', 'font_color': 'white'})
            date_format = workbook.add_format(
                {'num_format': 'd mmmm yyyy', 'fg_color': 'black', 'valign': 'vcenter', 'font_color': 'white',
                 'border_color': 'white', 'border': 2})
            center = workbook.add_format(
                {'font_size': 14, 'align': 'center', 'fg_color': 'orange', 'valign': 'vcenter', 'font_color': 'white',
                 'border_color': 'white', 'border': 2})
            sheet = workbook.add_worksheet('Color Stocks Move Report')
            k = 0
            j = 0
            for categ in report_color_lines.items():
                if categ[1][0]:
                    aa = 0
                    sheet.merge_range(k, 0, k + 6, 1, categ[0], center)
                    k += 7
                    sheet.merge_range(k, 0, k + 1, 1, 'اجمالي الوارد', format_1)
                    k += 2
                    sheet.merge_range(k, 0, k + 1, 1, categ[1][1], format_2)
                    k += 2
                    sheet.merge_range(k, 0, k + 1, 1, 'اجمالي المبيعات', format_1)
                    k += 2
                    sheet.merge_range(k, 0, k + 1, 1, categ[1][2], format_2)
                    k += 2
                    sheet.merge_range(k, 0, k + 1, 1, 'اجمالي المتبقي', format_1)
                    k += 2
                    sheet.merge_range(k, 0, k + 1, 1, categ[1][3], format_2)
                    i = 0
                    for product in categ[1][0]:
                        for item in product.values():
                            if not item['Branches']:
                                continue
                            i += 2
                            v = j
                            if item['Days']:
                                sec_percentage = item['sales'] / item['Days']
                            else:
                                sec_percentage = 0
                            if data['options'] == 'image':
                                sheet.merge_range(v, i, v + 16, i + 2, '', format_3)
                                if item['Image']:
                                    bytes_data = base64.b64decode(item['Image'])
                                    data_img = io.BytesIO(bytes_data)
                                    x_scale = 1.5
                                    y_scale = 2.45
                                    sheet.insert_image(xl_rowcol_to_cell(v, i), 'data_img',
                                                       {'image_data': data_img, 'x_scale': x_scale, 'y_scale': y_scale})
                                i += 3
                            sheet.write(j, i, 'العرض', format_3)
                            if item['Display_date']:
                                sheet.merge_range(j, i + 1, j, i + 3, item['Display_date'], date_format)
                            else:
                                sheet.merge_range(j, i + 1, j, i + 3, item['Display_date'], format_3)

                            sheet.write(j, i + 4, 'السعر', format_3)
                            sheet.write(j, i + 5, item['Price'], format_3)
                            sheet.merge_range(j + 1, i, j + 2, i, 'عدد الايام', format_3)
                            sheet.merge_range(j + 1, i + 1, j + 2, i + 1, item['Days'], format_3)
                            sheet.merge_range(j + 3, i, j + 4, i, 'متوسط', format_3)
                            sheet.merge_range(j + 3, i + 1, j + 4, i + 1, sec_percentage, format_3)
                            sheet.merge_range(j + 1, i + 2, j + 2, i + 3, 'الوان', format_4)
                            sheet.merge_range(j + 3, i + 2, j + 4, i + 3, item['Colors'], format_4)
                            sheet.merge_range(j + 1, i + 4, j + 6, i + 5, item['Code'], format_1)
                            sheet.merge_range(j + 5, i, j + 6, i, 'اللون', center)
                            sheet.merge_range(j + 5, i + 1, j + 6, i + 1, 'وارد', center)
                            sheet.merge_range(j + 5, i + 2, j + 6, i + 2, 'مبيعات', center)
                            sheet.merge_range(j + 5, i + 3, j + 6, i + 3, 'متبقي', center)
                            l = j + 7
                            aa = 9 - len(item["Branches"])
                            for line in item['Branches'].items():
                                sheet.write(l, i, line[0], format_3)
                                sheet.write(l, i + 1, line[1]['incoming'], format_3)
                                sheet.write(l, i + 2, line[1]['sales'], format_3)
                                sheet.write(l, i + 3, line[1]['balance'], format_3)
                                if line[1]['incoming'] > 0:
                                    sheet.write(l, i + 4,
                                                str(round(line[1]['sales'] / line[1]['incoming'] * 100, 2)) + '%',
                                                format_3)
                                else:
                                    sheet.write(l, i + 4, str(0) + '%', format_3)
                                if item['sales'] > 0:
                                    sheet.write(l, i + 5, str(round(line[1]['sales'] / item['sales'] * 100, 2)) + '%',
                                                format_3)
                                else:
                                    sheet.write(l, i + 5, str(0) + '%', format_3)
                                l += 1
                            if aa > 0:
                                sheet.merge_range(l, i, l + aa - 1, i + 5, "", format_4)
                                l += aa
                            sheet.write(l, i, 'الشركة', format_5)
                            sheet.write(l, i + 1, item['incoming'], format_5)
                            sheet.write(l, i + 2, item['sales'], format_5)
                            sheet.write(l, i + 3, item['incoming'] - item['sales'], format_5)
                            if item['incoming'] > 0:
                                sheet.write(l, i + 4, str(round(item['sales'] / item['incoming'] * 100, 2)) + '%',
                                            format_5)
                            else:
                                sheet.write(l, i + 4, str(0) + '%', format_5)
                            sheet.write(l, i + 5, '', format_5)
                            if data['options'] == 'image':
                                sheet.merge_range(l + 1, i - 3, l + 2, i + 3, "", format_1)
                                sheet.merge_range(l + 1, i + 4, l + 2, i + 5, "ملاحظات ادارية", center)
                            else:
                                sheet.merge_range(l + 1, i, l + 2, i + 3, "", format_1)
                                sheet.merge_range(l + 1, i + 4, l + 2, i + 5, "ملاحظات ادارية", center)
                            i += 5
                    if aa < 0:
                        sheet.merge_range(k + 2, 0, k - aa - 1, 1, '', format_2)
                        k -= aa
                    k += 3
                    j = k
        else:
            raise UserError("There is no Data available.")
