"""Report Writer workbook contracts, imports, calculation, and report preparation."""
from copy import copy
import csv
from io import StringIO
from datetime import date
import hashlib
import re
import colorsys
import posixpath
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path, PureWindowsPath
from io import BytesIO

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries

from .formulas import Calculator, FormulaError, format_value, excel_text


class WorkbookError(ValueError):
    pass


def read_source(path):
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f'Input file does not exist: {path}')
    if path.suffix.lower() in ('.xlsx', '.xlsm'):
        source = openpyxl.load_workbook(path, data_only=False)
        calc = Calculator(source)
        sheet = source.worksheets[0]
        try:
            return [[calc.cell(sheet.title, c.coordinate) for c in row] for row in sheet]
        finally:
            source.close()
    if path.suffix.lower() != '.csv':
        raise WorkbookError(f'Unsupported input format: {path.suffix}; use CSV or XLSX')
    raw = path.read_bytes()
    try:
        text = raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        text = raw.decode('cp1252')
    def typed(value):
        value = value.strip()
        if value == '':
            return None
        # Match numeric CSV type inference, including Excel trailing-minus syntax.
        numeric = '-' + value[:-1] if value.endswith('-') else value
        if re.fullmatch(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?', numeric):
            value = float(numeric)
            return int(value) if value.is_integer() else value
        return value
    return [[typed(v) for v in row] for row in csv.reader(StringIO(text))]


def parse_formats(specification, count):
    # The legacy contract concatenates marker and format, e.g. format10.0%:
    # that is column 1 with format 0.0%, not a column-10 marker.
    markers = '|'.join(sorted((str(i) for i in range(1, count + 1)), key=len, reverse=True))
    matches = list(re.finditer(r'format(' + markers + r')', specification, re.I))
    result = {}
    for i, match in enumerate(matches):
        result[int(match[1])] = specification[match.end():matches[i + 1].start() if i + 1 < len(matches) else None]
    if [int(match[1]) for match in matches] != list(range(1, count + 1)):
        raise WorkbookError(f'Mass-table format string must define format1 through format{count}')
    return [result[i] for i in range(1, count + 1)]


class ReportWorkbook:
    def __init__(self, path, project_dir=None, output_dir=None, today=None, macro_hook=None):
        self.path = Path(path).resolve()
        self.project_dir = Path(project_dir).resolve() if project_dir else self.path.parent
        self.output_dir = Path(output_dir).resolve() if output_dir else self.project_dir / 'Final Reports'
        self.book = openpyxl.load_workbook(self.path, keep_vba=self.path.suffix.lower() == '.xlsm')
        self.calc = Calculator(self.book, today=today)
        self.macro_hook = macro_hook
        self.warnings = []
        self.events = []
        self.imported = []
        self.chart_names = self._chart_names()
        # Portable project roots explicitly replace workstation paths, in memory only.
        self.sheet('Current Parameters')['A2'] = str(self.project_dir)
        self.sheet('Current Parameters')['A8'] = str(self.output_dir)
        self.calc.invalidate()

    def close(self):
        self.book.close()

    def _chart_names(self):
        """Use the actual drawing object names, not assumed Chart N numbering."""
        ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'xdr':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
        result={}
        def target(base,value):
            return value.lstrip('/') if value.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(base),value))
        def relationships(archive,part):
            path=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
            return {r.attrib['Id']:target(part,r.attrib['Target']) for r in ET.fromstring(archive.read(path))} if path in archive.namelist() else {}
        with zipfile.ZipFile(self.path) as archive:
            rels=relationships(archive,'xl/workbook.xml')
            for sheet in ET.fromstring(archive.read('xl/workbook.xml')).find('s:sheets',ns):
                part=rels[sheet.attrib['{'+ns['r']+'}id']]
                links=relationships(archive,part)
                names=[]
                for drawing in ET.fromstring(archive.read(part)).findall('s:drawing',ns):
                    location=links[drawing.attrib['{'+ns['r']+'}id']]
                    for node in ET.fromstring(archive.read(location)).findall('.//xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr',ns):
                        names.append(node.attrib.get('name',''))
                result[sheet.attrib['name'].casefold()]=names
        return result

    def color(self,color):
        if color is None:
            return None
        if color.type=='rgb':
            return color.rgb[-6:]
        if color.type=='indexed':
            from openpyxl.styles.colors import COLOR_INDEX
            return COLOR_INDEX[color.indexed][-6:] if color.indexed<len(COLOR_INDEX) else None
        if color.type=='theme' and self.book.loaded_theme:
            root=ET.fromstring(self.book.loaded_theme)
            scheme=root.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme')
            indexes=[1,0,3,2,4,5,6,7,8,9,10,11]
            rgb=list(scheme)[indexes[color.theme]][0].attrib
            value=rgb.get('val') if re.fullmatch(r'[0-9A-Fa-f]{6}',rgb.get('val','')) else rgb.get('lastClr')
            if value and color.tint:
                h,l,s=colorsys.rgb_to_hls(*(int(value[i:i+2],16)/255 for i in (0,2,4)))
                l=l*(1+color.tint) if color.tint<0 else l*(1-color.tint)+color.tint
                value=''.join(f'{round(v*255):02X}' for v in colorsys.hls_to_rgb(h,l,s))
            return value
        return None

    def sheet(self, name):
        return self.calc.sheet(name)

    def value(self, sheet, address):
        return self.calc.cell(sheet, address)

    def option(self, row):
        return excel_text(self.value('Current Parameters', f'B{row}')).casefold() == 'on'

    def report_ids(self):
        result = []
        sheet = self.sheet('Report Loop List')
        for row in range(2, sheet.max_row + 1):
            value = self.value(sheet.title, f'A{row}')
            if value in (None, ''):
                break
            result.append(excel_text(value))
        return result

    def resolve_path(self, value):
        value = str(value)
        candidate = Path(value.replace('\\', '/'))
        if candidate.is_absolute() and (candidate.exists() or candidate.is_relative_to(self.project_dir)
                                       or str(candidate).casefold().startswith(str(self.project_dir).casefold())):
            return candidate
        if candidate.is_absolute() and not PureWindowsPath(value).drive:
            return candidate
        if not PureWindowsPath(value).is_absolute() and not candidate.is_absolute():
            return self.project_dir / candidate
        # Relocate absolute paths embedded in mass-index rows under project Results.
        parts = PureWindowsPath(value).parts if '\\' in value else candidate.parts
        for i, part in enumerate(parts):
            if part.casefold() in ('results', 'additional docs'):
                return self.project_dir.joinpath(*parts[i:])
        # A path outside the portable project requires an existing explicit destination.
        if PureWindowsPath(value).is_absolute() and Path(value).exists():
            return Path(value)
        raise WorkbookError(f'Cannot relocate external path {value!r}; put inputs under the project or update its configuration')

    def import_data(self):
        sheet = self.sheet('FigureDataImport')
        base = self.resolve_path(self.value(sheet.title, 'A2'))
        self.imported = []
        for row in range(4, min(100, max(sheet.max_row, 4)) + 1):
            if self.value(sheet.title, f'E{row}') in (None, ''):
                continue
            filename = self.value(sheet.title, f'J{row}')
            if not filename:
                raise WorkbookError(f'FigureDataImport!J{row}: missing input filename')
            path = base / str(filename)
            data = read_source(path)
            target = self.sheet(str(self.value(sheet.title, f'F{row}')))
            start = str(self.value(sheet.title, f'G{row}'))
            source_start = str(self.value(sheet.title, f'K{row}'))
            c1, r1, _, _ = range_boundaries(start)
            sc, sr, _, _ = range_boundaries(source_start)
            rows = int(self.value(sheet.title, f'H{row}'))
            cols = int(self.value(sheet.title, f'I{row}'))
            if rows < 0 or cols < 0 or rows * cols > 1_000_000:
                raise WorkbookError(f'FigureDataImport row {row}: invalid import dimensions')
            for i in range(rows):
                for j in range(cols):
                    value = data[sr + i - 1][sc + j - 1] if sr + i <= len(data) and sc + j <= len(data[sr + i - 1]) else None
                    target.cell(r1 + i, c1 + j).value = value
            self.calc.invalidate()
            self.imported.append({'row': row, 'file': str(path), 'sheet': target.title, 'start': start, 'rows': rows, 'columns': cols})
        self.events.append('imports')

    def substitute_index(self):
        identifier = self.value('Current Parameters', 'B31')
        source = self.sheet('index_altered' if identifier not in (None, '') else 'index')
        destination = self.sheet('index_temp')
        # Original clears a fixed scratch region; no mutation of the saved workbook.
        for row in destination.iter_rows(min_row=1, max_row=1000, min_col=1, max_col=8):
            for cell in row:
                cell.value = None
        for row in range(1, source.max_row + 1):
            path = self.value(source.title, f'A{row}')
            if path in (None, ''):
                break
            for col in range(1, 9):
                value = self.value(source.title, f'{get_column_letter(col)}{row}')
                if col == 1 and identifier not in (None, ''):
                    value = str(value).replace(str(identifier), str(self.value('Current Parameters', 'A11')))
                destination.cell(row, col).value = value
        self.calc.invalidate()
        self.events.append('index_substitution')

    def run_macro_hook(self):
        if self.macro_hook is not None:
            self.macro_hook(self)
            self.calc.invalidate()
            self.events.append('python_macro_hook')
            return
        # User explicitly deferred execution of boutique VBA queries in this version.
        self.events.append('vba_bypassed')
        self.warnings.append('User Defined Macros is On; VBA execution is deferred by user direction. No boutique queries were run.')

    def prepare(self, report_id):
        self.events, self.warnings = [], []
        self.sheet('Current Parameters')['A11'] = report_id
        self.calc.invalidate()
        if self.option(29):
            self.substitute_index()
        if self.option(19):
            self.import_data()
        if self.option(33):
            self.run_macro_hook()
        if self.option(35):
            self.events.append('demo_mode_without_word_window')
            self.warnings.append('Demo Mode is On; this engine reports progress but does not display a live Word automation window.')
        self.events.append('prepared')

    def formatted(self, sheet, address):
        cell = self.sheet(sheet)[address]
        return format_value(self.value(sheet, address), cell.number_format)

    def range_data(self, sheet, address):
        return self.calc.rectangle(sheet, address)

    def cell_style(self, cell):
        rgb=self.color
        style = {'font': cell.font.name, 'font_size': cell.font.sz, 'bold': cell.font.b,
                 'italic': cell.font.i, 'font_color': rgb(cell.font.color),
                 'fill': rgb(cell.fill.fgColor) if cell.fill.patternType == 'solid' else None,
                 'horizontal': cell.alignment.horizontal, 'vertical': cell.alignment.vertical,
                 'borders': {side: {'style': getattr(cell.border, side).style,
                                    'color': rgb(getattr(cell.border, side).color)}
                             for side in ('left', 'right', 'top', 'bottom')}}
        return style

    def table_element(self, row):
        sheet = str(self.value('Report Numbers', f'P{row}'))
        address = str(self.value('Report Numbers', f'Q{row}'))
        source = self.sheet(sheet)
        c1, r1, c2, r2 = range_boundaries(address)
        styles, values = [], []
        for r in range(r1, r2 + 1):
            values.append([self.formatted(sheet, f'{get_column_letter(c)}{r}') for c in range(c1, c2 + 1)])
            styles.append([self.cell_style(source.cell(r, c)) for c in range(c1, c2 + 1)])
        merges = []
        for merged in source.merged_cells.ranges:
            if merged.min_row >= r1 and merged.max_row <= r2 and merged.min_col >= c1 and merged.max_col <= c2:
                merges.append([merged.min_row - r1, merged.min_col - c1, merged.max_row - r1, merged.max_col - c1])
            elif not (merged.max_row < r1 or merged.min_row > r2 or merged.max_col < c1 or merged.min_col > c2):
                raise WorkbookError(f'{sheet}!{address} cuts through merged range {merged}')
        mode = 'shell' if self.value('Report Numbers', f'T{row}') else 'new'
        types = self.value('Report Numbers', f'U{row}')
        return {'type': 'legacy_table', 'values': values, 'cell_styles': styles, 'merges': merges,
                'mode': mode, 'variable_rows': types == 'Variable Row',
                'header_rows': int(self.value('Report Numbers', f'V{row}') or 0),
                'align': 'center' if self.value('Report Numbers', f'S{row}') else 'left',
                'fit': bool(self.value('Report Numbers', f'R{row}')),
                'widths_inches': [((source.column_dimensions[get_column_letter(c)].width or 13) * 7 + 5) / 96 for c in range(c1, c2 + 1)],
                'row_heights_points': [source.row_dimensions[r].height for r in range(r1, r2 + 1)],
                'padding_px': self.value('Report Numbers', f'Y{row}'),
                'padding_sides': str(self.value('Report Numbers', f'Z{row}') or ''),
                'title': excel_text(self.value('Report Numbers', f'W{row}')),
                'alt_text': excel_text(self.value('Report Numbers', f'X{row}'))}

    def chart_element(self, row):
        source_name = str(self.value('Report Numbers', f'I{row}'))
        source = self.sheet(source_name)
        name = str(self.value('Report Numbers', f'J{row}'))
        names=self.chart_names.get(source_name.casefold(),[])
        index=next((i for i,n in enumerate(names) if n.casefold()==name.casefold()),-1)
        if index<0:
            raise WorkbookError(f'Chart {name!r} does not exist on {source_name}; available names: {names}')
        if index >= len(source._charts):
            raise WorkbookError(f'Chart {name} does not exist on {source_name}')
        chart = source._charts[index]
        if not isinstance(chart, openpyxl.chart.BarChart):
            raise WorkbookError(f'Unsupported Excel chart type: {chart.__class__.__name__}')
        address = str(self.value('Report Numbers', f'K{row}'))
        data = self.range_data(source_name, address)
        if not data or any(len(r) < 2 for r in data):
            raise WorkbookError(f'Chart range {source_name}!{address} needs category plus numeric series')
        c1, r1, c2, r2 = range_boundaries(address)
        if chart.grouping not in ('clustered', 'standard'):
            raise WorkbookError(f'Unsupported Excel bar grouping: {chart.grouping}')
        categories = [excel_text(r[0]) for r in data]
        series = [[float(r[i] or 0) for r in data] for i in range(1, c2 - c1 + 1)]
        percent = '%' in source.cell(r1, c1 + 1).number_format
        override = self.value('Report Numbers', f'L{row}')
        if override:
            column = column_index_from_string(str(override))
            alt = '\n'.join(excel_text(self.value(source_name, f'{get_column_letter(column)}{r}')) for r in range(r1, r2 + 1))
        else:
            alt = '\n'.join(' '.join(
                f'The {self.formatted(source_name, f"{get_column_letter(c)}1")} for the '
                f'{self.formatted(source_name, f"{get_column_letter(c1)}1")} of {categories[i]} is '
                f'{self.formatted(source_name, f"{get_column_letter(c)}{r1+i}")}'
                for c in range(c1 + 1, c2 + 1)) for i in range(len(data)))
        width, height = chart.width / 2.54, chart.height / 2.54
        if hasattr(chart.anchor, 'ext'):
            width, height = chart.anchor.ext.cx / 914400, chart.anchor.ext.cy / 914400
        elif hasattr(chart.anchor, 'to'):
            # Two-cell anchors: preserve actual span instead of openpyxl's defaults.
            start, end = chart.anchor._from, chart.anchor.to
            width = (sum((source.column_dimensions[get_column_letter(c+1)].width or 13)*7+5 for c in range(start.col, end.col)) / 96
                     + (end.colOff-start.colOff)/914400)
            height = sum((source.row_dimensions[r+1].height or 15)/72 for r in range(start.row,end.row)) + (end.rowOff-start.rowOff)/914400
        return {'type': 'chart', 'labels': categories, 'values': series[0], 'series': series,
                'series_names': [self.formatted(source_name, f'{get_column_letter(c)}1') for c in range(c1+1,c2+1)],
                'direction': chart.type, 'percent': percent, 'width_inches': width, 'height_inches': height,
                'alt_text': alt if self.option(21) else '', 'title': ''}

    def transform_bullets(self,sheet_name,mode):
        """Port all three bullets.vb transforms; caller is opt-in, as in legacy code."""
        if mode not in (1,2,3):
            raise WorkbookError(f'Unsupported legacy bullet mode: {mode}')
        sheet=self.sheet(sheet_name)
        from openpyxl.styles import Font
        for row in range(2,52):
            label=self.value(sheet_name,f'E{row}')
            is_heading=(self.value(sheet_name,f'F{row}') or 0)>=1
            if is_heading:
                sheet.cell(row,1).value=label
                sheet.cell(row,2).value=None
                sheet.cell(row,1).font=Font(name='Times New Roman',size=12)
            else:
                sheet.cell(row,1).value='  l'
                sheet.cell(row,1).font=Font(name='Wingdings',size=8)
                extra=self.value(sheet_name,f'G{row}')
                sheet.cell(row,2).value=excel_text(label)+(f' ({excel_text(extra)})' if mode==1 and extra not in (None,'') else '')
            sheet.cell(row,2).font=Font(name='Times New Roman',size=12)
        if self.value(sheet_name,'F2') in (None,''):
            fallback={1:'A common cause is qualified by a cause that occurs more than once, thus there are no common causes listed.',
                      2:'No HCBS Medical Review errors were found in the sample.',
                      3:'No HCBS Data Processing errors were found in the sample.'}
            sheet['B2']=fallback[mode]
            sheet['B2'].font=Font(name='Times New Roman',size=12)
        self.calc.invalidate()

    def make_bullets(self,import_row):
        mode=int(self.value('FigureDataImport',f'L{import_row}') or 0)
        if mode in (1,2,3):
            self.transform_bullets(str(self.value('FigureDataImport',f'F{import_row}')),mode)


    def elements(self):
        elements = {}
        sheet = self.sheet('Report Numbers')
        for id_col, builder in [('A', lambda r: {'type': 'text', 'value': self.formatted(sheet.title, f'C{r}') }),
                                ('G', self.chart_element), ('N', self.table_element)]:
            for row in range(2, sheet.max_row + 1):
                tag = self.value(sheet.title, f'{id_col}{row}')
                if tag in (None, ''):
                    break
                key = '@#$_' + str(tag)
                if key.casefold() in {k.casefold() for k in elements}:
                    raise WorkbookError(f'Duplicate element tag: {key}')
                elements[key] = builder(row)
        return elements

    def mass_elements(self):
        result = {}
        if not self.option(29):
            return result
        sheet = self.sheet('index_temp')
        for row in range(1, sheet.max_row + 1):
            tag = self.value(sheet.title, f'C{row}')
            if tag in (None, ''):
                break
            data = read_source(self.resolve_path(self.value(sheet.title, f'A{row}')))
            address = str(self.value(sheet.title, f'B{row}'))
            c1, r1, c2, r2 = range_boundaries(address)
            formats = parse_formats(str(self.value(sheet.title, f'F{row}')), c2-c1+1)
            mode = str(self.value(sheet.title, f'E{row}'))
            if mode not in ('1', '2', '3'):
                raise WorkbookError(f'index_temp!E{row}: unsupported mass-table mode {mode}')
            if mode == '2':
                r2 = r1 - 1
                while r2 < len(data) and c1 <= len(data[r2]) and data[r2][c1-1] not in (None, ''):
                    r2 += 1
            rows = [[format_value(data[r-1][c-1] if r <= len(data) and c <= len(data[r-1]) else None, formats[c-c1])
                     for c in range(c1,c2+1)] for r in range(r1,r2+1)]
            result['@#$_'+str(tag)] = ({'type':'text','value':rows[0][0]} if c1==c2 and r1==r2 else
                                     {'type':'legacy_table','values':rows,'mode':'shell','variable_rows':mode=='2','mass':True})
        return result

    def sections(self):
        result = []
        if not self.option(15):
            return result
        sheet = self.sheet('SmartSections')
        for row in range(2, sheet.max_row+1):
            identifier = self.value(sheet.title,f'A{row}')
            if identifier in (None,''):
                break
            result.append({'start':f'Xsection{excel_text(identifier)}S', 'end':f'Xsection{excel_text(identifier)}E',
                           'include':bool(self.value(sheet.title,f'B{row}')), 'reason':excel_text(self.value(sheet.title,f'D{row}')),
                           'bullet_mode':int(self.value(sheet.title,f'G{row}') or 0)})
        return result

    def postprocess(self):
        replacements = []
        try:
            sheet = self.sheet('SmartSections')
        except FormulaError:
            sheet = None
        if sheet:
            for start, before, after, prefix in [('J','K','L','Table'),('P','Q','R','Figure')]:
                for row in range(2,sheet.max_row+1):
                    if self.value(sheet.title,f'{start}{row}') in (None,''):
                        break
                    replacements.append((f'{prefix}X {excel_text(self.value(sheet.title,f"{before}{row}"))}',
                                         f'{prefix} {excel_text(self.value(sheet.title,f"{after}{row}"))}'))
        replacements += [(s,'N/A') for s in ('-3,149','-$3,149','-314,900.0%','-21,700.0%')]
        if self.option(37):
            sheet = self.sheet('GlobalReplace')
            for row in range(2,sheet.max_row+1):
                before = self.value(sheet.title,f'A{row}')
                if before in (None,''):
                    break
                replacements.append((str(before),excel_text(self.value(sheet.title,f'B{row}'))))
        headers=[]
        if self.option(17):
            sheet=self.sheet('Header and Footer')
            for row in range(2,7):
                headers.append({'section':row-1,'header':self.value(sheet.title,f'B{row}'),
                                'header_align':self.value(sheet.title,f'C{row}'),
                                'footer':self.value(sheet.title,f'D{row}'), 'footer_align':self.value(sheet.title,f'E{row}')})
        documents=[]
        sheet=self.sheet('Add Docx')
        for row in range(5,sheet.max_row+1):
            tag=self.value(sheet.title,f'A{row}')
            if tag in (None,''):
                break
            filename=str(self.value(sheet.title,f'B{row}'))
            path=self.resolve_path(self.value(sheet.title,'D2'))/filename
            documents.append({'tag':'@#$_'+str(tag),'path':str(path),'filename':filename})
        return {'replacements':replacements,'headers':headers,'documents':documents}

    def output_name(self, report_id, extension='.docx'):
        template=str(self.value('Current Parameters','A5'))
        if '_Template' not in template:
            raise WorkbookError('Template filename must contain _Template, or provide an explicit output pattern')
        name=template.split('_Template')[0]+'_'+str(report_id)+excel_text(self.value('Current Parameters','B27'))+extension
        if any(c in name for c in ('/','\\',':','*','?','"','<','>','|')):
            raise WorkbookError(f'Output filename has invalid characters: {name}')
        return name

    def save_prepared(self,destination):
        """Save a separate imported workbook with freshly computed formula caches.

        The selected workbook is never overwritten. openpyxl preserves the VBA
        archive for XLSM; arbitrary Excel extensions outside its support may not
        round-trip and are not claimed supported.
        """
        stream=BytesIO();self.book.save(stream)
        namespace='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        with zipfile.ZipFile(stream) as original:
            parts={i.filename:original.read(i.filename) for i in original.infolist()}
        for index,sheet in enumerate(self.book.worksheets,1):
            part=f'xl/worksheets/sheet{index}.xml'
            tree=ET.fromstring(parts[part])
            for cell in tree.iter('{'+namespace+'}c'):
                if cell.find('{'+namespace+'}f') is None:
                    continue
                value=self.value(sheet.title,cell.attrib['r'])
                cached=cell.find('{'+namespace+'}v')
                if cached is None:
                    cached=ET.SubElement(cell,'{'+namespace+'}v')
                cell.set('t','b' if isinstance(value,bool) else ('str' if isinstance(value,str) else 'n'))
                cached.text=('1' if value else '0') if isinstance(value,bool) else excel_text(value)
            parts[part]=ET.tostring(tree,encoding='UTF-8',xml_declaration=True)
        output=BytesIO()
        with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as package:
            for name,data in parts.items():
                package.writestr(name,data)
        from .files import save_bytes
        save_bytes(output.getvalue(),destination)
        return str(destination)
