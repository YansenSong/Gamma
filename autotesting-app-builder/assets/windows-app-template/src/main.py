"""IVISTA visual workbench; all business data comes from the generated project."""
from __future__ import annotations
import argparse,json,sys
from pathlib import Path
from PySide6.QtCore import QSettings,QThread,Qt,QUrl,Signal
from PySide6.QtGui import QDesktopServices,QIcon,QPixmap
from PySide6.QtWidgets import (QApplication,QCheckBox,QComboBox,QDialog,QDoubleSpinBox,QFileDialog,QFormLayout,QFrame,QHBoxLayout,QLabel,QLineEdit,QListWidget,QListWidgetItem,QMainWindow,QMessageBox,QProgressBar,QPushButton,QScrollArea,QVBoxLayout,QWidget)
from app_service import load_facts_from_excel,load_spec,resource_path,run_job,run_job_from_facts
from ui_design_system.components import FileDropArea
from ui_design_system.settings import DesktopSettings,SettingsDialog
from field_review_ui import FieldReviewDialog

def read_facts(path,spec):
    return load_facts_from_excel(path,spec) if path.suffix.lower()=='.xlsx' else json.loads(path.read_text(encoding='utf-8-sig'))

class BatchWorker(QThread):
    progress=Signal(int,int,str);generated=Signal(object);failed=Signal(str,str)
    def __init__(self,jobs,output,spec):super().__init__();self.jobs=jobs;self.output=output;self.spec=spec
    def run(self):
        for index,(label,value) in enumerate(self.jobs):
            self.progress.emit(index,len(self.jobs),'正在处理：'+label)
            try:self.generated.emit(run_job_from_facts(read_facts(value,self.spec) if isinstance(value,Path) else value,self.output))
            except Exception as exc:self.failed.emit(label,str(exc))
            self.progress.emit(index+1,len(self.jobs),f'已处理 {index+1}/{len(self.jobs)}')

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__();self.spec=load_spec();app=self.spec['application'];ui=app.get('ui',{})
        self.setWindowTitle(app['name']);self.resize(1100,850);self.setMinimumSize(980,760)
        if app.get('logo'):self.setWindowIcon(QIcon(str(resource_path(app['logo']))))
        self.store=QSettings('AutoTestingGenerated',app['id']);self.preferences=DesktopSettings.load(self.store,app['id'])
        self.setWindowFlag(Qt.WindowStaysOnTopHint,self.preferences.always_on_top)
        self.selected_files=[];self.results=[];self.errors=[];self.worker=None;self.controls={};self.input_path=None
        root=QWidget();root.setObjectName('root');self.setCentralWidget(root);outer=QVBoxLayout(root);outer.setContentsMargins(32,28,32,28);outer.setSpacing(18)
        hero=QFrame();hero.setObjectName('heroCard');row=QHBoxLayout(hero);row.setContentsMargins(26,22,26,22);copy=QVBoxLayout();copy.setSpacing(5)
        brand=QLabel(ui.get('brand',app['id']));brand.setObjectName('brandLabel');copy.addWidget(brand,alignment=Qt.AlignLeft)
        title=QLabel(app['name']);title.setObjectName('pageTitle');copy.addWidget(title)
        mode=self.spec['scoring'].get('mode')
        subtitle=QLabel(ui.get('subtitle',f"导入测试数据，生成评分表、检验报告与审计记录 · Excel 原生公式结果"));subtitle.setObjectName('subtitle');subtitle.setWordWrap(True);copy.addWidget(subtitle);row.addLayout(copy,1)
        if app.get('logo'):
            logo=QLabel();logo.setPixmap(QPixmap(str(resource_path(app['logo']))).scaled(64,64,Qt.KeepAspectRatio,Qt.SmoothTransformation));row.addWidget(logo)
        self.settings_button=QPushButton('设置');self.settings_button.setObjectName('settingsButton');self.settings_button.setFixedSize(72,32);self.settings_button.clicked.connect(self.open_settings);row.addWidget(self.settings_button);outer.addWidget(hero)
        workspace=QHBoxLayout();workspace.setSpacing(18);left=QVBoxLayout();right=QVBoxLayout();left.setSpacing(18);right.setSpacing(18);workspace.addLayout(left,5);workspace.addLayout(right,6);outer.addLayout(workspace,1)
        card,layout=self.card('1','测评项目与数据录入','填写当前记录，或从文件列表导入后编辑');category=QLabel(ui.get('category',app['name']));category.setObjectName('countBadge');layout.addWidget(category)
        scroll=QScrollArea();scroll.setWidgetResizable(True);host=QWidget();host.setObjectName('formHost');form=QFormLayout(host);form.setSpacing(10);form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapAllRows);scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        for definition in self.spec['scoring']['inputs']:
            key=definition['id'];kind=definition.get('type','number')
            if kind=='boolean':control=QCheckBox();control.setChecked(bool(definition.get('default',False)))
            elif kind=='string' and definition.get('enum'):
                control=QComboBox();control.addItems([str(x) for x in definition['enum']]);control.setCurrentText(str(definition.get('default',definition['enum'][0])))
            elif kind=='number':
                control=QDoubleSpinBox();control.setDecimals(int(definition.get('decimals',3)));control.setRange(float(definition.get('minimum',-1e9)),float(definition.get('maximum',1e9)));control.setValue(float(definition.get('default',definition.get('minimum',0) or 0)))
            else:control=QLineEdit(str(definition.get('default','')))
            control.setObjectName(key);self.controls[key]=control;field_label=QLabel(definition.get('label',key)+(f" ({definition['unit']})" if definition.get('unit') else ''));field_label.setWordWrap(True);field_label.setMaximumWidth(360);form.addRow(field_label,control)
        scroll.setWidget(host);layout.addWidget(scroll,1);left.addWidget(card,1)
        card,layout=self.card('2','添加测试数据','批量处理使用各文件的输入；当前表单可单独生成');self.drop_area=FileDropArea()
        self.drop_area.set_hint('JSON / Excel · 支持一次拖入多个文件' if self.spec.get('excel_input') else 'JSON · 支持一次拖入多个文件');self.drop_area.files_dropped.connect(self.add_files);self.drop_area.clicked.connect(self.choose_input);layout.addWidget(self.drop_area)
        row=QHBoxLayout();self.choose_button=QPushButton('选择文件');self.choose_button.setObjectName('secondaryButton');self.choose_button.clicked.connect(self.choose_input);self.clear_button=QPushButton('清空列表');self.clear_button.setObjectName('ghostButton');self.clear_button.clicked.connect(self.clear_files);row.addWidget(self.choose_button);row.addWidget(self.clear_button);layout.addLayout(row);left.addWidget(card)
        card=QFrame();card.setObjectName('panelCard');layout=QVBoxLayout(card);layout.setContentsMargins(22,18,22,18);row=QHBoxLayout();title=QLabel('待处理文件');title.setObjectName('listTitle');row.addWidget(title);self.count_label=QLabel('0 个文件');self.count_label.setObjectName('countBadge');row.addWidget(self.count_label);row.addStretch();layout.addLayout(row)
        self.file_list=QListWidget();self.file_list.setMinimumHeight(116);self.file_list.setSpacing(4);self.file_list.itemDoubleClicked.connect(self.import_selected);layout.addWidget(self.file_list,1);hint=QLabel('双击文件，将记录导入左侧表单');hint.setObjectName('sectionHint');layout.addWidget(hint);right.addWidget(card,1)
        card,layout=self.card('3','生成评分与报告','三件套成组输出，同名目录自动编号，不覆盖已有文件');card.setObjectName('resultCard');self.output=QLineEdit(self.preferences.output_directory);self.output.setObjectName('outputLocation');layout.addWidget(self.output)
        self.progress=QProgressBar();self.progress.setRange(0,1);self.progress.setValue(0);self.progress.setFormat('等待开始');layout.addWidget(self.progress);self.status_label=QLabel('可直接录入，或选择 / 拖入测试数据');self.status_label.setObjectName('statusLabel');self.status_label.setWordWrap(True);layout.addWidget(self.status_label)
        self.result_list=QListWidget();self.result_list.setMaximumHeight(150);self.result_list.itemDoubleClicked.connect(self.open_selected_result);layout.addWidget(self.result_list)
        self.run_form_button=QPushButton('生成当前表单');self.run_form_button.setObjectName('secondaryButton');self.run_form_button.clicked.connect(self.run);self.run_batch_button=QPushButton('开始批量生成');self.run_batch_button.setObjectName('primaryButton');self.run_batch_button.clicked.connect(self.run_batch);self.run_batch_button.setEnabled(False)
        row=QHBoxLayout();row.addWidget(self.run_form_button);row.addWidget(self.run_batch_button);layout.addLayout(row);row=QHBoxLayout();self.open_report_button=QPushButton('打开最近报告');self.open_report_button.clicked.connect(self.open_report);self.open_report_button.setEnabled(False);folder=QPushButton('打开输出目录');folder.setObjectName('ghostButton');folder.clicked.connect(self.open_output);row.addWidget(self.open_report_button);row.addWidget(folder);layout.addLayout(row);right.addWidget(card)
        if self.spec.get('report_review'):
            self.review_button=QPushButton('审核报告字段来源');self.review_button.setObjectName('secondaryButton');self.review_button.clicked.connect(self.open_field_review);layout.addWidget(self.review_button)
        if self.spec.get('review_preview'):
            self.preview_button=QPushButton('生成字段审核预览（不评分）');self.preview_button.setObjectName('secondaryButton');self.preview_button.clicked.connect(self.create_review_preview);layout.addWidget(self.preview_button)
        self.setStyleSheet(resource_path('src/ui_design_system/workbench.qss').read_text(encoding='utf-8'))
    @staticmethod
    def card(number,title,hint):
        card=QFrame();card.setObjectName('panelCard');layout=QVBoxLayout(card);layout.setContentsMargins(22,18,22,20);layout.setSpacing(12);row=QHBoxLayout();badge=QLabel(number);badge.setObjectName('stepBadge');badge.setAlignment(Qt.AlignCenter);badge.setFixedSize(30,30);row.addWidget(badge,alignment=Qt.AlignTop);copy=QVBoxLayout();heading=QLabel(title);heading.setObjectName('sectionTitle');detail=QLabel(hint);detail.setObjectName('sectionHint');detail.setWordWrap(True);copy.addWidget(heading);copy.addWidget(detail);row.addLayout(copy,1);layout.addLayout(row);return card,layout
    def set_status(self,text,state='neutral'):
        self.status_label.setText(text);self.status_label.setProperty('state',state);self.status_label.style().unpolish(self.status_label);self.status_label.style().polish(self.status_label)
    def choose_input(self):
        paths,_=QFileDialog.getOpenFileNames(self,'选择测评数据','','测评数据 (*.json *.xlsx)' if self.spec.get('excel_input') else '测评记录 (*.json)');self.add_files(paths)
    def add_files(self,paths):
        allowed={'.json','.xlsx'} if self.spec.get('excel_input') else {'.json'};rejected=[]
        for value in paths:
            path=Path(value).resolve()
            if not path.is_file() or path.suffix.lower() not in allowed:rejected.append(path.name);continue
            if path not in self.selected_files:
                self.selected_files.append(path);item=QListWidgetItem(path.name);item.setToolTip(str(path));item.setData(Qt.UserRole,str(path));self.file_list.addItem(item)
        self.count_label.setText(f'{len(self.selected_files)} 个文件');self.run_batch_button.setEnabled(bool(self.selected_files));self.set_status('无法添加：'+', '.join(rejected) if rejected else f'已添加 {len(self.selected_files)} 个文件。双击可导入表单。','error' if rejected else 'neutral')
    def clear_files(self):self.selected_files.clear();self.file_list.clear();self.count_label.setText('0 个文件');self.run_batch_button.setEnabled(False)
    def import_selected(self,item):
        try:self.input_path=Path(item.data(Qt.UserRole));self.populate(read_facts(self.input_path,self.spec));self.set_status('已导入：'+self.input_path.name)
        except Exception as exc:self.set_status('无法导入：'+str(exc),'error');QMessageBox.critical(self,'无法导入',str(exc))
    def open_field_review(self):
        review=json.loads(resource_path(self.spec['report_review']['path']).read_text(encoding='utf-8'))
        package_path=self.spec.get('business_review',{}).get('package')
        package=json.loads(resource_path(package_path).read_text(encoding='utf-8')) if package_path else None
        FieldReviewDialog(review,self,package=package).exec()
    def create_review_preview(self):
        from review_preview import create_review_preview
        try:
            result=create_review_preview(self.spec,self.form_facts(),self.output.text(),resource_path)
            self.set_status('已生成非正式字段审核预览（未评分）：'+result['bundle'])
            QDesktopServices.openUrl(QUrl.fromLocalFile(result['report']))
        except Exception as exc:
            self.set_status('预览失败：'+str(exc),'error');QMessageBox.critical(self,'预览失败',str(exc))
    def populate(self,facts):
        for key,control in self.controls.items():
            if key not in facts:
                definition=next(x for x in self.spec['scoring']['inputs'] if x['id']==key)
                if definition.get('required',True):continue
                value=definition.get('default','')
            else:value=facts[key]
            if isinstance(control,QCheckBox):control.setChecked(bool(value))
            elif isinstance(control,QComboBox):
                if control.findText(str(value))<0:raise ValueError(f'{key} 的值不在可选范围中')
                control.setCurrentText(str(value))
            elif isinstance(control,QDoubleSpinBox):
                if value is None or not control.minimum()<=float(value)<=control.maximum():raise ValueError(f'{key} 的值缺失或超出范围')
                control.setValue(float(value))
            else:control.setText('' if value is None else str(value))
    def form_facts(self):return {key:(control.isChecked() if isinstance(control,QCheckBox) else control.currentText() if isinstance(control,QComboBox) else control.value() if isinstance(control,QDoubleSpinBox) else control.text()) for key,control in self.controls.items()}
    def run(self):self.start_jobs([('当前表单',self.form_facts())])
    def run_batch(self):self.start_jobs([(path.name,path) for path in self.selected_files])
    def start_jobs(self,jobs):
        if not jobs or (self.worker and self.worker.isRunning()):return
        if not self.output.text().strip():QMessageBox.warning(self,'缺少输出目录','请选择输出目录。');return
        self.preferences.output_directory=self.output.text().strip();self.preferences.save(self.store);self.errors=[]
        for control in self.busy_controls():control.setEnabled(False)
        self.worker=BatchWorker(jobs,self.output.text(),self.spec);self.worker.progress.connect(self.on_progress);self.worker.generated.connect(self.on_generated);self.worker.failed.connect(self.on_failed);self.worker.finished.connect(self.on_finished);self.worker.start()
    def busy_controls(self):return [self.run_form_button,self.run_batch_button,self.choose_button,self.clear_button,self.drop_area,self.settings_button,self.output,self.file_list]
    def on_progress(self,index,total,text):self.progress.setRange(0,total);self.progress.setValue(index);self.progress.setFormat('%v / %m');self.set_status(text,'busy')
    def on_generated(self,result):
        self.results.append(result);item=QListWidgetItem(result['report'].name+' · '+result['bundle'].name);item.setToolTip(str(result['report']));item.setData(Qt.UserRole,str(result['report']));self.result_list.addItem(item);self.open_report_button.setEnabled(True)
    def on_failed(self,label,message):self.errors.append((label,message));self.result_list.addItem(f'失败：{label} · {message}')
    def on_finished(self):
        for control in self.busy_controls():control.setEnabled(True)
        self.run_batch_button.setEnabled(bool(self.selected_files));self.set_status(f'处理完成；失败 {len(self.errors)} 项。'+('请检查列表中的错误信息。' if self.errors else '报告、评分表和审计记录已成组保存。'),'error' if self.errors else 'success')
        if self.errors:QMessageBox.warning(self,'部分记录未完成','\n'.join(f'{label}: {message}' for label,message in self.errors))
        elif self.preferences.completion_action=='folder':self.open_output()
        elif self.preferences.completion_action=='report':self.open_report()
    def open_settings(self):
        dialog=SettingsDialog(DesktopSettings(self.output.text(),self.preferences.completion_action,self.preferences.always_on_top),self)
        if dialog.exec()==QDialog.DialogCode.Accepted:
            self.preferences=dialog.values();self.preferences.save(self.store);self.output.setText(self.preferences.output_directory);self.setWindowFlag(Qt.WindowStaysOnTopHint,self.preferences.always_on_top);self.show();self.set_status('设置已保存')
    def choose_output(self):
        path=QFileDialog.getExistingDirectory(self,'选择输出目录',self.output.text())
        if path:self.output.setText(path)
    def open_output(self):
        path=Path(self.output.text());path.mkdir(parents=True,exist_ok=True);QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.resolve())))
    def open_report(self):
        if self.results:QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.results[-1]['report'])))
    def open_selected_result(self,item):
        target=item.data(Qt.UserRole)
        if not target:return
        path=Path(target)
        if path.is_file():QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.resolve())))
    def closeEvent(self,event):
        if self.worker and self.worker.isRunning():QMessageBox.information(self,'正在生成','请等待当前处理完成后关闭窗口。');event.ignore();return
        super().closeEvent(event)

def main():
    parser=argparse.ArgumentParser();mode=parser.add_mutually_exclusive_group();mode.add_argument('--run-json');mode.add_argument('--review-json');parser.add_argument('--output');args=parser.parse_args()
    if args.review_json:
        if not args.output:parser.error('--review-json requires --output')
        try:
            from review_preview import create_review_preview
            spec=load_spec();facts=json.loads(Path(args.review_json).read_text(encoding='utf-8-sig'))
            result=create_review_preview(spec,facts,args.output,resource_path)
            print(json.dumps({'status':'engineering_preview','business_status':'blocked',**result},ensure_ascii=False));return 0
        except Exception as exc:print(json.dumps({'status':'failed','error':str(exc)},ensure_ascii=False));return 2
    if args.run_json:
        if not args.output:parser.error('--run-json requires --output')
        try:result=run_job(args.run_json,args.output);print(json.dumps({'status':'passed','bundle':str(result['bundle']),'scores':result['scores']},ensure_ascii=False));return 0
        except Exception as exc:print(json.dumps({'status':'failed','error':str(exc)},ensure_ascii=False));return 2
    app=QApplication(sys.argv);window=MainWindow();window.show();return app.exec()
if __name__=='__main__':raise SystemExit(main())
