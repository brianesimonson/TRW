"""Automatic heading TOC/pagination and DOC/PDF export through LibreOffice CLI."""
import os
from pathlib import Path
import shutil
import subprocess
import tempfile


class RendererError(ValueError):
    pass


class LibreOfficeRenderer:
    """Render PDFs and rebuild heading TOCs until page references converge.

Uses LibreOffice's supported command-line export, not its optional UNO bindings.
The source templates' heading-based TOC is supported; unsupported fields/switches
are explicit errors. Word's own pagination may differ from LibreOffice's.
"""
    def __init__(self,soffice=None,timeout=120):
        self.soffice=soffice or shutil.which('soffice') or shutil.which('libreoffice')
        if not self.soffice and os.name=='nt':
            for variable in ('ProgramFiles','ProgramFiles(x86)'):
                path=Path(os.environ.get(variable,''))/'LibreOffice/program/soffice.exe'
                if path.is_file():
                    self.soffice=str(path)
                    break
        self.timeout=timeout

    def export(self,source,directory,extension='.pdf'):
        if not self.soffice:
            raise RendererError('Automatic TOC/PDF/DOC output requires LibreOffice. Install LibreOffice or use --skip-fields for an explicitly unfinished DOCX draft.')
        source,directory=Path(source).resolve(),Path(directory).resolve()
        directory.mkdir(parents=True,exist_ok=True)
        with tempfile.TemporaryDirectory(prefix='trw-lo-profile-') as folder:
            filters={'.pdf':'pdf:writer_pdf_Export:{"ExportBookmarks":{"type":"boolean","value":"true"},"UseTaggedPDF":{"type":"boolean","value":"true"}}',
                     '.doc':'doc:MS Word 97'}
            try:
                result=subprocess.run([self.soffice,'-env:UserInstallation='+Path(folder).as_uri(),
                                       '--headless','--norestore','--convert-to',filters[extension],
                                       '--outdir',str(directory),str(source)],capture_output=True,text=True,timeout=self.timeout)
            except subprocess.TimeoutExpired as exc:
                raise RendererError(f'Document export exceeded {self.timeout}s') from exc
        expected=directory/(source.stem+extension)
        if result.returncode or not expected.is_file():
            raise RendererError('LibreOffice export failed: '+(result.stderr or result.stdout)[-2000:])
        return expected

    def finalize(self,input_path,destination,update_fields=True,extension='.docx'):
        destination=Path(destination).resolve()
        if destination.exists():
            raise FileExistsError(f'Refusing to overwrite {destination}')
        if extension not in ('.docx','.pdf','.doc'):
            raise RendererError('Unsupported output extension: '+extension)
        from .toc import rebuild_toc,pdf_outline
        with tempfile.TemporaryDirectory(prefix='trw-finalize-') as folder:
            root=Path(folder)
            current=root/'pass0.docx';shutil.copyfile(input_path,current)
            entries=[];pages=None;converged=not update_fields
            pdf=None
            if update_fields:
                for iteration in range(1,7):
                    pdf=self.export(current,root/f'layout{iteration}')
                    next_doc=root/f'pass{iteration}.docx'
                    new_entries,pages=rebuild_toc(current,pdf,next_doc)
                    if new_entries==entries:
                        converged=True
                        current=next_doc
                        break
                    entries=new_entries;current=next_doc
                if not converged:
                    raise RendererError('TOC page references did not converge after six layout passes')
            # Export the final document, then verify the exact published layout.
            if extension in ('.pdf','.doc'):
                artifact=self.export(current,root/'final',extension)
            else:
                artifact=current
            if update_fields:
                checked_pdf=artifact if extension=='.pdf' else self.export(current,root/'verified')
                actual,pages=pdf_outline(checked_pdf)
                expected=[e for e in actual if e['level'] in {v['level'] for v in entries}]
                if expected!=entries:
                    raise RendererError('Final rendered headings/pages differ from the populated TOC')
            destination.parent.mkdir(parents=True,exist_ok=True)
            with tempfile.NamedTemporaryFile(dir=destination.parent,suffix=extension,delete=False) as handle:
                temporary=Path(handle.name)
            try:
                shutil.copyfile(artifact,temporary);os.chmod(temporary,0o644);os.link(temporary,destination)
            finally:
                temporary.unlink(missing_ok=True)
            return {'status':'ok','renderer':'LibreOffice CLI + Python heading TOC',
                    'indexes_updated':1 if update_fields else 0,'page_count':pages,
                    'toc_entries':entries,'layout_passes':iteration if update_fields else 0}
