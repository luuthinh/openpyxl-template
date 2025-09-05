# -*- coding: utf-8 -*-
from openpyxl_template import ExcelTemplate

tpl = ExcelTemplate("templates/row_block.xlsx", max_row_tmpl=8, max_col_tmpl=5)

context = {
    "data": {
        "id": 1,
        "customer": "Alice Johnson",
        "mobile": "+1-202-555-0147",
        "address": "123 Main Street, Springfield, IL 62704",
        "date": "2025-09-04",
        "delivery_date": "2025-09-10",
        "items": [
            {"seq": 1,"product": "Laptop", "qty": 2, "price": 800},
            {"seq": 2,"product": "Mouse", "qty": 5, "price": 20},
        ],
        "total": 2 * 800 + 5 * 20,
    },
}

tpl.render(context)
tpl.save("output/row_block.xlsx")