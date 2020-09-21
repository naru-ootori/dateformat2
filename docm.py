# coding=utf-8

from docx.shared import *
from docx.enum.text import *
from docx.enum.table import *

def table_format(table):

    table.style            = 'Table Grid'
    table.alignment        = WD_TABLE_ALIGNMENT.CENTER
    table.rows.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    table.rows[0].height   = Cm(0.8)
    
    for col in table.columns:
        
        for cell in col.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            for par in cell.paragraphs:
                par.paragraph_format.alignment         = WD_ALIGN_PARAGRAPH.LEFT
                par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                par.paragraph_format.space_before      = Pt(0)
                par.paragraph_format.space_after       = Pt(0)
                
                for run in par.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)
                    
    for cell in table.row_cells(0):
    
        for par in cell.paragraphs:
            par.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            for run in par.runs:
                run.font.bold = True
                