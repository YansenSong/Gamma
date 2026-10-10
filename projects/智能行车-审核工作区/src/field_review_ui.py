"""Read-only source review; user entries are data, never business approvals."""
import csv
import json
from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QLineEdit, QTableWidget, QTableWidgetItem, QPushButton, QFileDialog, QMessageBox

class FieldReviewDialog(QDialog):
    def __init__(self,review,parent=None,package=None):
        super().__init__(parent)
        self.review=review
        self.package=package
        self.setWindowTitle('报告字段来源审核 · 待确认')
        self.resize(1060,680)
        layout=QVBoxLayout(self)
        label=QLabel('候选映射、用户补录和“未提供”均不代表批准。导出的清单可填写审核意见；正式报告仍须完整业务审批。')
        label.setWordWrap(True);layout.addWidget(label)
        self.filter=QLineEdit();self.filter.setPlaceholderText('搜索字段、分类、来源或缺失原因');layout.addWidget(self.filter)
        self.table=QTableWidget();self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(['字段 / 含义','分类 / 状态','Word位置','来源','缺失原因 / 问题','正式阻断'])
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table)
        self.export_button=QPushButton('导出绑定版本的业务审核表（CSV）' if package else '导出可填写审核表（CSV）');self.export_button.clicked.connect(self.export);layout.addWidget(self.export_button)
        self.filter.textChanged.connect(self.populate);self.populate()
    def populate(self,*args):
        query=self.filter.text().casefold()
        fields=[x for x in self.review['fields'] if query in json.dumps(x,ensure_ascii=False).casefold()]
        self.table.setRowCount(len(fields))
        for row,item in enumerate(fields):
            values=[item['field_id']+' / '+item['meaning'],item['category']+' / '+item['review_status'],
                    '; '.join(item['word_locations']),json.dumps(item['source'],ensure_ascii=False),
                    item['missing_reason']+' / '+item['question'],'是' if item['blocks_official'] else '工程说明']
            for col,text in enumerate(values):
                cell=QTableWidgetItem(text);cell.setToolTip(text);self.table.setItem(row,col,cell)
        for col,width in enumerate([250,140,200,240,350,100]):self.table.setColumnWidth(col,width)
    def export(self):
        filename,_=QFileDialog.getSaveFileName(self,'导出审核清单','report-field-review.csv','CSV (*.csv)')
        if not filename:return
        if self.package:
            from pathlib import Path
            keys=['package_sha256','item_id','content_sha256','subject','question','locations','reviewer','reviewed_at','decision','reason','evidence_paths','evidence_sha256','record_ids']
            try:
                with open(filename,'w',encoding='utf-8-sig',newline='') as stream:
                    writer=csv.DictWriter(stream,fieldnames=keys,lineterminator='\n');writer.writeheader()
                    for item in self.package['items']:
                        writer.writerow({'package_sha256':self.package['package_sha256'],**{k:item[k] for k in ('item_id','content_sha256','subject','question')},'locations':'; '.join(item['locations'])})
                Path(filename).with_suffix('.package.json').write_text(json.dumps(self.package,ensure_ascii=False,indent=2)+'\n',encoding='utf-8',newline='\n')
                QMessageBox.information(self,'已导出','CSV 与同名 package.json 已导出。填写真实审核人、带时区日期、意见和依据；证据文件路径、SHA-256、记录编号各填 JSON 数组。导入后仍须复验，不会自动批准。')
            except OSError as exc:QMessageBox.warning(self,'导出失败',str(exc))
            return
        fields=self.review['fields'];keys=list(fields[0])+['reviewer','decision','reason']
        try:
            with open(filename,'w',encoding='utf-8-sig',newline='') as stream:
                writer=csv.DictWriter(stream,fieldnames=keys);writer.writeheader()
                for item in fields:
                    writer.writerow({key:json.dumps(value,ensure_ascii=False) if isinstance(value,(list,dict)) else value for key,value in item.items()})
        except OSError as exc:QMessageBox.warning(self,'导出失败',str(exc))
