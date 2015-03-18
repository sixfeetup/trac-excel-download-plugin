# -*- coding: utf-8 -*-

import re
from cStringIO import StringIO
from datetime import datetime
from decimal import Decimal
from unicodedata import east_asian_width
from xlwt import XFStyle, Style, Alignment, Borders, Pattern, Font

from trac.core import TracError
from trac.util.text import to_unicode

from tracexceldownload.translation import _


class WorksheetWriterError(TracError): pass


class WorksheetWriter(object):

    MAX_ROWS = 65536

    def __init__(self, sheet, req):
        self.sheet = sheet
        self.req = req
        self.tz = req.tz
        if hasattr(req, 'locale'):
            self.ambiwidth = (1, 2)[str(req.locale)[:2] in ('ja', 'kr', 'zh')]
        else:
            self.ambiwidth = 1
        self.styles = self._get_excel_styles()
        self.row_idx = 0
        self._col_widths = {}
        self._metrics_cache = {}
        self._cells_count = 0

    _normalize_newline = re.compile(r'\r\n?').sub

    def move_row(self):
        self.row_idx += 1
        if self.row_idx >= self.MAX_ROWS:
            raise WorksheetWriterError(_(
                "Number of rows in the Excel sheet exceeded the limit of "
                "65536 rows"))
        self._flush_row()

    def write_row(self, cells):
        _get_style = self._get_style
        _set_col_width = self._set_col_width
        _normalize_newline = self._normalize_newline
        get_metrics = self.get_metrics
        tz = self.tz
        has_tz_normalize = hasattr(tz, 'normalize')  # pytz

        row = self.sheet.row(self.row_idx)
        max_line = 1
        max_height = 0
        for idx, (value, style, width, line) in enumerate(cells):
            if isinstance(value, basestring):
                if isinstance(value, str):
                    value = to_unicode(value)
                value = value.rstrip()
                if '\r' in value:
                    value = _normalize_newline('\n', value)
            elif isinstance(value, datetime):
                value = value.astimezone(tz)
                if has_tz_normalize: # pytz
                    value = tz.normalize(value)
                value = datetime(*(value.timetuple()[0:6]))
                if style == '[date]':
                    width = len('YYYY-MM-DD')
                elif style == '[time]':
                    width = len('HH:MM:SS')
                else:
                    width = len('YYYY-MM-DD HH:MM:SS')
                _set_col_width(idx, width)
                row.set_cell_date(idx, value, _get_style(style))
                continue
            elif isinstance(value, (int, long, float, Decimal)):
                _set_col_width(idx, len('%g' % value))
                row.set_cell_number(idx, value, _get_style(style))
                continue
            elif value is True or value is False:
                _set_col_width(idx, 1)
                row.set_cell_number(idx, int(value), _get_style(style))
                continue
            if width is None or line is None:
                metrics = get_metrics(value)
                if width is None:
                    width = metrics[0]
                if line is None:
                    line = metrics[1]
            if max_line < line:
                max_line = line
            _set_col_width(idx, width)
            style = _get_style(style)
            if max_height < style.font.height:
                max_height = style.font.height
            row.write(idx, value, style)
            self._cells_count += 1
        row.height = min(max_line, 10) * max(max_height * 255 / 180, 255)
        row.height_mismatch = True
        self.move_row()

    def _flush_row(self):
        if self.row_idx % 512 == 0 or self._cells_count >= 4096:
            self.sheet.flush_row_data()
            self._cells_count = 0

    def _get_style(self, style):
        if isinstance(style, basestring):
            if style not in self.styles:
                if style.endswith(':change'):
                    style = '*:change'
                else:
                    style = '*'
            style = self.styles[style]
        return style

    def _set_col_width(self, idx, width):
        widths = self._col_widths
        widths.setdefault(idx, 1)
        if widths[idx] < width:
            widths[idx] = width

    def set_col_widths(self):
        for idx, width in self._col_widths.iteritems():
            self.sheet.col(idx).width = (1 + min(width, 50)) * 256

    def get_metrics(self, value):
        if not value:
            return 0, 1
        if isinstance(value, str):
            value = to_unicode(value)
        if value not in self._metrics_cache:
            lines = value.splitlines()
            doubles = ('WFA', 'WF')[self.ambiwidth == 1]
            width = max(sum((1, 2)[east_asian_width(ch) in doubles]
                            for ch in line)
                        for line in lines)
            if len(value) > 64:
                return width, len(lines)
            self._metrics_cache[value] = (width, len(lines))
        return self._metrics_cache[value]

    def _get_excel_styles(self):
        def style_base():
            style = XFStyle()
            style.alignment.vert = Alignment.VERT_TOP
            style.alignment.wrap = True
            style.font.height = 180 # 9pt
            borders = style.borders
            borders.left = Borders.THIN
            borders.right = Borders.THIN
            borders.top = Borders.THIN
            borders.bottom = Borders.THIN
            return style

        header = XFStyle()
        header.font.height = 400 # 20pt
        header2 = XFStyle()
        header2.font.height = 320 # 16pt

        thead = style_base()
        thead.font.bold = True
        thead.font.colour_index = Style.colour_map['white']
        thead.pattern.pattern = Pattern.SOLID_PATTERN
        thead.pattern.pattern_fore_colour = Style.colour_map['black']
        thead.borders.colour = 'white'
        thead.borders.left = Borders.THIN
        thead.borders.right = Borders.THIN
        thead.borders.top = Borders.THIN
        thead.borders.bottom = Borders.THIN

        def style_change(style):
            pattern = style.pattern
            pattern.pattern = Pattern.SOLID_PATTERN
            pattern.pattern_fore_colour = Style.colour_map['light_orange']
            return style

        def style_id():
            style = style_base()
            style.font.underline = Font.UNDERLINE_SINGLE
            style.font.colour_index = Style.colour_map['blue']
            style.num_format_str = '"#"0'
            return style

        def style_milestone():
            style = style_base()
            style.font.underline = Font.UNDERLINE_SINGLE
            style.font.colour_index = Style.colour_map['blue']
            style.num_format_str = '@'
            return style

        def style_time():
            style = style_base()
            style.num_format_str = 'HH:MM:SS'
            return style

        def style_date():
            style = style_base()
            style.num_format_str = 'YYYY-MM-DD'
            return style

        def style_datetime():
            style = style_base()
            style.num_format_str = 'YYYY-MM-DD HH:MM:SS'
            return style

        def style_default():
            style = style_base()
            style.num_format_str = '@'    # String
            return style

        styles = {'header': header, 'header2': header2, 'thead': thead}
        for key, func in (('id', style_id),
                          ('milestone', style_milestone),
                          ('[time]', style_time),
                          ('[date]', style_date),
                          ('[datetime]', style_datetime),
                          ('*', style_default)):
            styles[key] = func()
            styles['%s:change' % key] = style_change(func())
        return styles


def get_workbook_content(book):
    out = StringIO()
    book.save(out)
    return out.getvalue()


def get_literal(text):
    return u'"%s"' % to_unicode(text).replace('"', '""')
