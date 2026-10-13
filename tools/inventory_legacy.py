"""Read the original archives and enumerate VB.NET, VBA, and workbook structures.

Read-only: never builds the add-in, executes VBA, or writes the uploaded source.
Optional audit dependencies: openpyxl==3.1.5, oletools==0.60.2.
This is a declaration scanner, not a VB compiler or proof of dynamic reachability.
"""
import argparse
from collections import Counter
import hashlib
from io import BytesIO
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

MODIFIERS = r'(?:(?:Public|Private|Friend|Protected|Shared|ReadOnly|WriteOnly|Overrides|Overloads|Default|NotOverridable|Overridable|MustOverride|NotInheritable|MustInherit|Static|Partial)\s+)*'
DECL = re.compile(r'^\s*' + MODIFIERS + r'(?:(Declare)\s+(?:Ansi\s+|Unicode\s+|Auto\s+)?(Sub|Function)|(Sub|Function|Property))\s+(?:(Get|Let|Set)\s+)?(\[?\w+\]?)', re.I)
SCOPE = re.compile(r'^\s*' + MODIFIERS + r'(Module|Class)\s+(\w+)', re.I)


def mask(line):
    """Blank string literals and remove VB comments; retain token positions."""
    result, quoted, i = [], False, 0
    while i < len(line):
        char = line[i]
        if char == '"':
            if quoted and i + 1 < len(line) and line[i + 1] == '"':
                result.extend('  ')
                i += 2
                continue
            quoted = not quoted
            result.append(' ')
        elif not quoted and char == "'":
            break
        else:
            result.append(' ' if quoted else char)
        i += 1
    return ''.join(result)


def routines(text, file, language):
    lines = text.splitlines()
    result, scopes = [], []
    for i, line in enumerate(lines):
        code = mask(line)
        if re.match(r'^\s*End\s+(Class|Module)\b', code, re.I):
            if scopes:
                scopes.pop()
        scope = SCOPE.match(code)
        if scope:
            scopes.append(scope[2])
        match = DECL.match(code)
        if not match:
            continue
        kind = match[2] or match[3]
        name = match[5].strip('[]')
        signature = line.strip()
        signature_end = i
        while signature.count('(') > signature.count(')') or signature.endswith('_'):
            signature_end += 1
            if signature_end >= len(lines):
                raise ValueError('Unterminated declaration: ' + file)
            signature = signature.rstrip('_').rstrip() + ' ' + lines[signature_end].strip()
        if match[1]:
            end = signature_end
            kind = 'Declare ' + kind
        else:
            end = next((j for j in range(signature_end + 1, len(lines))
                        if re.match(r'^\s*End\s+' + kind + r'\b', mask(lines[j]), re.I)), None)
            if end is None:
                raise ValueError('Missing End ' + kind + ': ' + file + ':' + str(i + 1))
        body = '\n'.join(lines[signature_end + 1:end])
        code_body = '\n'.join(mask(l) for l in lines[signature_end + 1:end])
        result.append({
            'id': f'{language}:{file}:{i+1}', 'file': file,
            'scope': '.'.join(scopes), 'name': name, 'kind': kind,
            'start_line': i + 1, 'end_line': end + 1, 'signature': signature,
            'body_sha256': hashlib.sha256(body.encode()).hexdigest(),
            'empty_body': not code_body.strip(),
            'handles': re.findall(r'Handles\s+(.+)', signature, re.I),
            'worksheet_names': sorted(set(re.findall(r'Worksheets\("([^"]+)"\)', body, re.I))),
            '_code': code_body,
        })
    return result


def build(code_archive, package_archive):
    import openpyxl
    from oletools.olevba import VBA_Parser
    ns = {'m': 'http://schemas.microsoft.com/developer/msbuild/2003'}
    files, members, project_info = [], [], {}
    with zipfile.ZipFile(code_archive) as archive:
        projects = [n for n in archive.namelist() if n.endswith('.vbproj')]
        if len(projects) != 1:
            raise ValueError('Expected exactly one VB project')
        project_name = projects[0]
        project = ET.fromstring(archive.read(project_name))
        prefix = project_name.rsplit('/', 1)[0] + '/'
        compiled = [e.attrib['Include'].replace('\\', '/') for e in project.findall('.//m:Compile', ns)]
        project_info = {
            'archive_project': project_name,
            'target_framework': project.find('.//m:TargetFrameworkVersion', ns).text,
            'compiled_files': compiled,
            'references': [e.attrib['Include'].split(',')[0] for e in project.findall('.//m:Reference', ns)],
            'com_references': [e.attrib['Include'] for e in project.findall('.//m:COMReference', ns)],
            'imports': [e.attrib['Project'] for e in project.findall('.//m:Import', ns) if 'Project' in e.attrib],
        }
        names = [n for n in archive.namelist() if n.startswith(prefix) and n.endswith('.vb')
                 and '/bin/' not in n and '/obj/' not in n]
        if {n.removeprefix(prefix) for n in names} != set(compiled):
            raise ValueError('Compiled VB source and archive membership do not match')
        for name in sorted(names):
            data = archive.read(name)
            relative = name.removeprefix(prefix)
            text = data.decode('utf-8-sig')
            declarations = routines(text, relative, 'vbnet')
            members.extend(declarations)
            files.append({'file': relative, 'line_count': len(text.splitlines()),
                          'sha256': hashlib.sha256(data).hexdigest(),
                          'generated': '.Designer.vb' in relative,
                          'declaration_count': len(declarations)})
        project_info['excluded_vb_entries'] = sorted(n for n in archive.namelist()
                                                   if n.endswith('.vb') and n not in names)
    # Evidence only: module-qualified calls, or same-file unqualified calls.
    # Does not resolve reflection, VBA Application.Run strings, or all VB overloads.
    for target in members:
        refs = []
        for caller in members:
            if caller['id'] == target['id'] or target['name'].lower() == 'new':
                continue
            qualified = re.compile(r'\b' + re.escape(target['scope'].split('.')[0]) + r'\.' + re.escape(target['name']) + r'\b', re.I)
            unqualified = re.compile(r'(?<![\w.])' + re.escape(target['name']) + r'\s*\(', re.I)
            if qualified.search(caller['_code']) or (caller['file'] == target['file'] and unqualified.search(caller['_code'])):
                refs.append(caller['id'])
        target['static_reference_candidates'] = sorted(refs)
    for member in members:
        del member['_code']
    workbooks = []
    documents = []
    with zipfile.ZipFile(package_archive) as archive:
        for name in archive.namelist():
            if name.endswith('.docx'):
                data = archive.read(name)
                word_ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                parts, tags, external = [], set(), 0
                with zipfile.ZipFile(BytesIO(data)) as package:
                    for part in package.namelist():
                        if part.endswith('.rels'):
                            external += sum(r.attrib.get('TargetMode') == 'External'
                                            for r in ET.fromstring(package.read(part)))
                        if not part.startswith('word/') or not part.endswith('.xml'):
                            continue
                        tree = ET.fromstring(package.read(part))
                        for paragraph in tree.findall('.//w:p', word_ns):
                            text = ''.join(n.text or '' for n in paragraph.findall('.//w:t', word_ns))
                            tags.update(re.findall(r'@#\$_[A-Za-z0-9_]+', text))
                        fields = [n.text or '' for n in tree.findall('.//w:instrText', word_ns)]
                        if fields:
                            parts.append({'part': part, 'field_instructions': fields})
                    media = sum(p.startswith('word/media/') for p in package.namelist())
                documents.append({'archive_path': name, 'name': Path(name).name,
                                  'sha256': hashlib.sha256(data).hexdigest(),
                                  'tags': sorted(tags), 'field_parts': parts,
                                  'external_relationship_count': external,
                                  'media_count': media})
            if not name.endswith('.xlsm'):
                continue
            data = archive.read(name)
            workbook = openpyxl.load_workbook(BytesIO(data), data_only=False, keep_vba=True)
            chart_definitions = []
            with zipfile.ZipFile(BytesIO(data)) as package:
                chart_ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
                for part in package.namelist():
                    if not re.match(r'xl/charts/chart\d+\.xml$', part):
                        continue
                    tree = ET.fromstring(package.read(part))
                    plot = tree.find('.//c:plotArea', chart_ns)
                    chart_types = [node.tag.rsplit('}', 1)[-1] for node in plot
                                   if node.tag.rsplit('}', 1)[-1].endswith('Chart')] if plot is not None else []
                    chart_definitions.append({'part': part, 'types': chart_types,
                                              'series_formula_references': [n.text for n in tree.findall('.//c:f', chart_ns)],
                                              'bar_directions': [n.attrib.get('val') for n in tree.findall('.//c:barDir', chart_ns)]})
            parser = VBA_Parser(name, data=data)
            modules, macros = [], []
            try:
                for _, stream, filename, code in parser.extract_macros():
                    declarations = routines(code, filename, 'vba')
                    for declaration in declarations:
                        del declaration['_code']
                        declaration['id'] = f'vba:{Path(name).name}:{filename}:{declaration["start_line"]}'
                    macros.extend(declarations)
                    modules.append({'file': filename, 'stream': stream,
                                    'sha256': hashlib.sha256(code.encode()).hexdigest(),
                                    'line_count': len(code.splitlines()),
                                    'declaration_count': len(declarations)})
            finally:
                parser.close()
            tabs = []
            for sheet in workbook:
                formula_count = 0
                formulas = []
                families = Counter()
                samples = {}
                for row in sheet:
                    for cell in row:
                        if cell.data_type != 'f':
                            continue
                        formula_count += 1
                        formula = str(cell.value)
                        formulas.append({'cell': cell.coordinate, 'formula': formula,
                                         'number_format': cell.number_format})
                        for func in re.findall(r'\b([A-Z][A-Z0-9_.]*)\s*\(', formula, re.I):
                            families[func.upper()] += 1
                            samples.setdefault(func.upper(), cell.coordinate)
                tabs.append({'name': sheet.title, 'visibility': sheet.sheet_state,
                             'formula_cells': formula_count, 'formula_functions': dict(sorted(families.items())),
                             'formula_example_cells': dict(sorted(samples.items())),
                             'formulas': formulas,
                             'chart_count': len(sheet._charts),
                             'merged_ranges': [str(r) for r in sheet.merged_cells.ranges]})
            workbooks.append({'archive_path': name, 'name': Path(name).name,
                              'sha256': hashlib.sha256(data).hexdigest(),
                              'tabs': tabs, 'chart_definitions': chart_definitions,
                              'vba_modules': modules, 'vba_declarations': macros})
            workbook.close()
    return {'schema_version': 1,
            'source_archives': [{'name': p.name, 'sha256': hashlib.sha256(p.read_bytes()).hexdigest()} for p in [code_archive, package_archive]],
            'project': project_info, 'vbnet_files': files, 'vbnet_declarations': members,
            'workbooks': workbooks, 'word_documents': documents}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('code_archive', type=Path)
    parser.add_argument('package_archive', type=Path)
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args()
    result = build(args.code_archive, args.package_archive)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2) + '\n')
    print(f'Inventoried {len(result["vbnet_files"])} VB.NET files, '
          f'{len(result["vbnet_declarations"])} declarations, '
          f'{len(result["workbooks"])} workbooks')


if __name__ == '__main__':
    main()
