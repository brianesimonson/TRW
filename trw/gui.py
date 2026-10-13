"""Small desktop front end for the Python engine (no Excel/Word automation)."""
from datetime import datetime
import json
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog,messagebox,ttk


class ReportWriterApp:
    def __init__(self,root):
        self.root=root;root.title('Report Writer — Python engine');root.geometry('780x490')
        self.messages=queue.Queue();self.running=False
        self.workbook=tk.StringVar();self.output=tk.StringVar();self.report=tk.StringVar()
        self.format=tk.StringVar(value='Workbook setting');self.draft=tk.BooleanVar(value=False)
        frame=ttk.Frame(root,padding=16);frame.pack(fill='both',expand=True)
        ttk.Label(frame,text='Report Writer',font=('Segoe UI',18)).grid(row=0,column=0,sticky='w',pady=(0,14))
        for row,label,variable,callback in [(1,'Excel configuration workbook',self.workbook,self.select_workbook),
                                             (2,'Output folder',self.output,self.select_output)]:
            ttk.Label(frame,text=label).grid(row=row,column=0,sticky='w',pady=5)
            ttk.Entry(frame,textvariable=variable,width=60).grid(row=row,column=1,sticky='ew',padx=8)
            ttk.Button(frame,text='Browse…',command=callback).grid(row=row,column=2)
        ttk.Label(frame,text='Single report ID').grid(row=3,column=0,sticky='w',pady=5)
        ttk.Entry(frame,textvariable=self.report).grid(row=3,column=1,sticky='w',padx=8)
        ttk.Label(frame,text='Leave blank to run every ID in Report Loop List.').grid(row=4,column=1,sticky='w',padx=8)
        ttk.Label(frame,text='Output format').grid(row=5,column=0,sticky='w',pady=5)
        ttk.Combobox(frame,textvariable=self.format,values=['Workbook setting','docx','pdf','doc'],state='readonly').grid(row=5,column=1,sticky='w',padx=8)
        ttk.Checkbutton(frame,text='Draft only: skip requested table-of-contents refresh',variable=self.draft).grid(row=6,column=0,columnspan=3,sticky='w',pady=8)
        buttons=ttk.Frame(frame);buttons.grid(row=7,column=0,columnspan=3,sticky='w',pady=8)
        self.run_button=ttk.Button(buttons,text='Generate reports',command=self.run);self.run_button.pack(side='left',padx=(0,8))
        self.import_button=ttk.Button(buttons,text='Import only',command=lambda:self.run(import_only=True));self.import_button.pack(side='left',padx=(0,8))
        ttk.Button(buttons,text='Open output folder',command=self.open_output).pack(side='left')
        self.log=tk.Text(frame,height=11,wrap='word',state='disabled');self.log.grid(row=8,column=0,columnspan=3,sticky='nsew',pady=8)
        frame.columnconfigure(1,weight=1);frame.rowconfigure(8,weight=1)
        root.after(100,self.poll)

    def select_workbook(self):
        value=filedialog.askopenfilename(title='Choose Report Writer workbook',filetypes=[('Excel workbooks','*.xlsm *.xlsx')])
        if value:
            self.workbook.set(value)
            if not self.output.get():self.output.set(str(Path(value).parent/'Python Reports'))

    def select_output(self):
        value=filedialog.askdirectory(title='Choose output folder')
        if value:self.output.set(value)

    def write(self,text):
        self.log.configure(state='normal');self.log.insert('end',text+'\n');self.log.see('end');self.log.configure(state='disabled')

    def run(self,import_only=False):
        if self.running:return
        path=Path(self.workbook.get())
        if not path.is_file() or not self.output.get():
            messagebox.showerror('Select project','Choose an Excel configuration workbook and output folder.');return
        # Capture Tk state on the UI thread, never from the worker.
        ids=[self.report.get()] if self.report.get().strip() else None
        export=None if self.format.get()=='Workbook setting' else self.format.get()
        draft=self.draft.get()
        output=Path(self.output.get())/datetime.now().strftime('run-%Y%m%d-%H%M%S-%f')
        self.running=True;self.run_button.configure(state='disabled');self.import_button.configure(state='disabled')
        self.write('Starting '+str(path))
        def worker():
            try:
                from .pipeline import run_workbook
                results=run_workbook(path,output=output,report_ids=ids,skip_fields=draft,export_format=export,
                                     import_only=import_only,progress=lambda event:self.messages.put(('progress',event)))
                self.messages.put(('done',results))
            except Exception as exc:
                self.messages.put(('error',str(exc)))
        threading.Thread(target=worker,daemon=True).start()

    def poll(self):
        try:
            while True:
                kind,value=self.messages.get_nowait()
                if kind=='progress':self.write(f'{value["report"]}: {value["stage"]}')
                elif kind=='error':
                    self.write('Failed: '+value);messagebox.showerror('Report generation failed',value);self.running=False
                else:
                    for result in value:
                        self.write('Saved '+result['output'])
                        for warning in result.get('warnings',[]):self.write('Note: '+warning)
                    self.running=False
                if not self.running:
                    self.run_button.configure(state='normal');self.import_button.configure(state='normal')
        except queue.Empty:pass
        self.root.after(100,self.poll)

    def open_output(self):
        import os,subprocess,sys
        path=Path(self.output.get())
        if not path.is_dir():return
        if os.name=='nt':os.startfile(path)
        elif sys.platform=='darwin':subprocess.Popen(['open',str(path)])
        else:subprocess.Popen(['xdg-open',str(path)])


def main():
    root=tk.Tk();ReportWriterApp(root);root.mainloop()


if __name__=='__main__':main()
