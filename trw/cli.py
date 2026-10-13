"""A small manifest adapter: CSV inputs become resolved report elements."""
import argparse
import csv
import json
import os
from pathlib import Path
import sys


def resolve(config, base, report_id):
    sources = {}
    for name, filename in config.get('sources', {}).items():
        with (base / filename.format(report_id=report_id)).open(newline='', encoding='utf-8-sig') as handle:
            sources[name] = list(csv.DictReader(handle))
        if not sources[name]:
            raise ValueError(f'Empty input: {name}')
    elements = {}
    for tag, spec in config['elements'].items():
        element = dict(spec)
        rows = sources.get(spec.get('source'), [])
        kind = spec['type']
        if kind == 'text':
            if 'sum_column' in spec:
                total = sum(float(row[spec['sum_column']]) for row in rows)
                element['value'] = str(int(total)) if total.is_integer() else str(total)
            else:
                element['value'] = str(spec['value']).format(report_id=report_id)
        elif kind == 'table' and 'source' in spec:
            element['rows'] = [[row[column] for column in spec['columns']] for row in rows]
        elif kind == 'chart' and 'source' in spec:
            element['labels'] = [row[spec['label_column']] for row in rows]
            element['values'] = [float(row[spec['value_column']]) for row in rows]
            element['alt_text'] = '\n'.join(
                f'{label}: {value:.1%}' if spec.get('percent') else f'{label}: {value:g}'
                for label, value in zip(element['labels'], element['values']))
        elif kind == 'bullets' and 'source' in spec:
            element['items'] = [spec['item_template'].format(
                report_id=report_id,
                **{key: value.lower() if key in spec.get('lowercase_columns', []) else value
                   for key, value in row.items()}) for row in rows]
        elements[tag] = element
    return elements


def main():
    parser = argparse.ArgumentParser(description='Generate template-based Word reports without Office')
    parser.add_argument('manifest', type=Path)
    parser.add_argument('--report', action='append', help='Report ID; repeat for several. Defaults to manifest batch.')
    parser.add_argument('--output', type=Path, default=Path('output'))
    parser.add_argument('--project-dir', type=Path, help='Portable project root for an XLSM/XLSX workbook (default: workbook folder)')
    parser.add_argument('--as-of', type=str, help='Date for workbook TODAY(), YYYY-MM-DD')
    parser.add_argument('--skip-fields', action='store_true', help='Explicitly generate a draft without requested TOC/page-number refresh')
    parser.add_argument('--update-toc',action=argparse.BooleanOptionalAction,default=None,
                        help='Override workbook TOC option; --no-update-toc keeps its cached contents')
    parser.add_argument('--format', choices=['docx','pdf','doc'], help='Override workbook output format')
    parser.add_argument('--import-only', action='store_true', help='Validate/prepare workbook imports without creating reports')
    parser.add_argument('--list-charts', action='store_true', help='List actual Excel chart object names without generating reports')
    parser.add_argument('--soffice', help='LibreOffice executable path')
    args = parser.parse_args()
    os.environ.setdefault('MPLCONFIGDIR', '/tmp/trw-matplotlib')
    os.environ.setdefault('XDG_CACHE_HOME', '/tmp/trw-cache')
    try:
        from .engine import render
        manifest = args.manifest.resolve()
        if manifest.suffix.lower() in ('.xlsm','.xlsx'):
            from datetime import date
            from .pipeline import run_workbook
            from .renderer import LibreOfficeRenderer
            if args.list_charts:
                from .workbook import ReportWorkbook
                workbook=ReportWorkbook(manifest,project_dir=args.project_dir)
                try:
                    print(json.dumps(workbook.chart_names,indent=2))
                finally:
                    workbook.close()
                return 0
            renderer=LibreOfficeRenderer(args.soffice)
            results=run_workbook(manifest,output=args.output,project_dir=args.project_dir,report_ids=args.report,
                                 today=date.fromisoformat(args.as_of) if args.as_of else None,
                                 skip_fields=args.skip_fields,export_format=args.format,renderer=renderer,
                                 update_toc=args.update_toc,
                                 import_only=args.import_only,
                                 progress=lambda event: print(json.dumps(event),file=sys.stderr))
            print(json.dumps(results,indent=2))
            return 0
        config = json.loads(manifest.read_text())
        ids = args.report or config['report_ids']
        results = []
        for report_id in ids:
            if not report_id or any(c not in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-' for c in report_id):
                raise ValueError('Report IDs must contain only letters, digits, underscores, or hyphens')
            destination = args.output / config['output_pattern'].format(report_id=report_id)
            if destination.resolve().parent != args.output.resolve():
                raise ValueError('Output pattern must be a filename inside the output directory')
            results.append(render(manifest.parent / config['template'], destination,
                                  resolve(config, manifest.parent, report_id)))
        print(json.dumps(results, indent=2))
        return 0
    except (ValueError, KeyError, OSError) as exc:
        print(f'Report generation failed: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
