# -*- coding: utf-8 -*-
"""
Created : 2015-03-12

@author: Eric Lapouyade
"""

from openpyxl_template import ExcelTemplate

tpl = ExcelTemplate("templates/cell_data.xlsx", max_row_tmpl=100, max_col_tmpl=100)

context = {
    "data": {
        "date": "2025-09-04",
        "text": "hello world",
        "number": "12020",
    },
}

tpl.render(context)
tpl.save("output/cell_data.xlsx")