"""评分器公共内核：包解析、三种文档读取、统一报告。"""
from .pkg import Package
from .xlsx import ExcelBook, Styles
from .docx import WordDoc, extract_word_parts
from .pptx import PptDeck, slide_no
from .pptcheck import (presets as ppt_presets, has_preset, filters as ppt_filters,
                       transition_types, transition_attrs, actions as ppt_actions,
                       shape_names, hlink_ids, has_gradfill, auto_number,
                       line_spacing, layout_short, placeholder_texts, title_text)
from .report import Report, run, setup_stdout
from .checks import (norm_formula, formula_has, func_name, style_name_is,
                     as_float, close, font_color, fill_color, has_border,
                     border_styles, ref_range_covers, norm_ref, has_field, field_name)

__all__ = ['Package', 'ExcelBook', 'Styles', 'WordDoc', 'extract_word_parts',
           'PptDeck', 'slide_no', 'Report', 'run', 'setup_stdout',
           'norm_formula', 'formula_has', 'func_name', 'style_name_is',
           'as_float', 'close', 'font_color', 'fill_color', 'has_border',
           'border_styles', 'ref_range_covers', 'norm_ref', 'has_field', 'field_name',
           'ppt_presets', 'has_preset', 'ppt_filters', 'transition_types',
           'transition_attrs', 'ppt_actions', 'shape_names', 'hlink_ids',
           'has_gradfill', 'auto_number', 'line_spacing', 'layout_short',
           'placeholder_texts', 'title_text']
