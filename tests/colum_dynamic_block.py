# -*- coding: utf-8 -*-

from openpyxl_template import ExcelTemplate

tpl = ExcelTemplate("templates/column_dynamic_block.xlsx", max_row_tmpl=8, max_col_tmpl=5)

context = {
    "data": {
        "id": 1,
        "name": "Alice Johnson",
        "class": "10A1",
        "programs": [
            {
                "name": "Math",
                "scores": [
                    {"subject": "Quiz1", "score": 8.5},
                    {"subject": "Quiz2", "score": 9.0},
                    {"subject": "Midterm", "score": 7.5},
                    {"subject": "Final", "score": 8.0},
                ],
            },
            {
                "name": "English",
                "scores": [
                    {"subject": "Listening", "score": 7.0},
                    {"subject": "Speaking", "score": 8.0},
                    {"subject": "Reading", "score": 7.5},
                    {"subject": "Writing", "score": 8.5},
                ],
            },
            {
                "name": "Science",
                "scores": [
                    {"subject": "Lab1", "score": 9.0},
                    {"subject": "Lab2", "score": 8.5},
                    {"subject": "Final", "score": 8.0},
                ],
            },            
        ],
    },
}

tpl.render(context)
tpl.save("output/column_dynamic_block.xlsx")