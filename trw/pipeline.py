"""Run one or all reports from the original Excel configuration workbook."""
from datetime import date
from pathlib import Path
import tempfile

from .workbook import ReportWorkbook, WorkbookError
from .document import generate


def run_workbook(path, output=None, project_dir=None, report_ids=None, today=None,
                 skip_fields=False, export_format=None, renderer=None, macro_hook=None,
                 progress=None, import_only=False,update_toc=None):
    workbook=ReportWorkbook(path,project_dir=project_dir,output_dir=output,today=today,macro_hook=macro_hook)
    try:
        ids=report_ids or workbook.report_ids()
        if not ids:
            raise WorkbookError('Report Loop List contains no report IDs')
        results=[]
        for position,report_id in enumerate(ids,1):
            if progress:
                progress({'stage':'prepare','report':report_id,'position':position,'total':len(ids)})
            workbook.prepare(report_id)
            if import_only:
                prepared=workbook.output_dir/(workbook.path.stem+'_'+str(report_id)+'_prepared'+workbook.path.suffix)
                workbook.save_prepared(prepared)
                results.append({'report':report_id,'imports':workbook.imported,'events':workbook.events,
                                'warnings':workbook.warnings,'output':str(prepared),'status':'prepared_workbook_saved'})
                continue
            requested=str(export_format or workbook.value('Current Parameters','B23'))
            extension={ 'Word (*.docx)':'.docx','97-2003 Word (*.doc)':'.doc','PDF (*.pdf)':'.pdf',
                        'docx':'.docx','doc':'.doc','pdf':'.pdf'}.get(requested)
            if extension is None:
                raise WorkbookError(f'Unsupported report output format: {requested}')
            destination=workbook.output_dir/workbook.output_name(report_id,extension)
            if destination.exists():
                raise FileExistsError(f'Refusing to overwrite {destination}')
            update_fields=workbook.option(25) if update_toc is None else bool(update_toc)
            needs_renderer=extension!='.docx' or (update_fields and not skip_fields)
            if needs_renderer and renderer is None:
                from .renderer import LibreOfficeRenderer
                renderer=LibreOfficeRenderer()
            if progress:
                progress({'stage':'document','report':report_id})
            with tempfile.TemporaryDirectory(prefix='trw-report-') as folder:
                intermediate=Path(folder)/'assembled.docx'
                result=generate(workbook.project_dir/str(workbook.value('Current Parameters','A5')),
                                intermediate,workbook.elements(),workbook.mass_elements(),
                                workbook.sections(),workbook.postprocess())
                if needs_renderer:
                    if progress:
                        progress({'stage':'render','report':report_id})
                    rendering=renderer.finalize(intermediate,destination,update_fields=update_fields and not skip_fields,
                                                extension=extension)
                    result['rendering']=rendering
                else:
                    from .files import save_document
                    from docx import Document
                    save_document(Document(intermediate),destination)
                result.update(output=str(destination),report=report_id,imports=workbook.imported,
                              events=list(workbook.events),warnings=list(workbook.warnings))
                if update_fields and skip_fields:
                    result['warnings'].append('TOC/field refresh was explicitly skipped; cached page numbers are not final.')
                result['status']='partial' if result['warnings'] else 'generated'
                results.append(result)
            if progress:
                progress({'stage':'complete','report':report_id,'output':str(destination)})
        return results
    finally:
        workbook.close()
