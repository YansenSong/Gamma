"""Project-scoped settings, independent of IVISTA report categories."""
from dataclasses import dataclass
from pathlib import Path
from PySide6.QtWidgets import QDialog,QDialogButtonBox,QFileDialog,QFormLayout,QLineEdit,QPushButton,QVBoxLayout,QComboBox,QCheckBox,QMessageBox,QFrame,QLabel

@dataclass
class DesktopSettings:
    output_directory: str
    completion_action: str = 'none'
    always_on_top: bool = False
    @classmethod
    def load(cls,store,project_id):
        return cls(str(store.value('output/directory',str(Path.home()/'Documents'/project_id))),str(store.value('completion/action','none')),store.value('window/always_on_top',False,type=bool))
    def save(self,store):
        store.setValue('output/directory',self.output_directory);store.setValue('completion/action',self.completion_action);store.setValue('window/always_on_top',self.always_on_top);store.sync()

class SettingsDialog(QDialog):
    def __init__(self,values,parent=None):
        super().__init__(parent);self.setWindowTitle('应用设置');self.setMinimumWidth(560)
        layout=QVBoxLayout(self);layout.setContentsMargins(26,24,26,22);layout.setSpacing(18)
        title=QLabel('应用设置');title.setStyleSheet('font-size:22px;font-weight:700;color:#172033');layout.addWidget(title)
        hint=QLabel('设置保存位置和常用桌面行为，修改后会自动保留。');hint.setStyleSheet('color:#778397;font-size:12px');layout.addWidget(hint)
        output_card=QFrame();output_card.setObjectName('panelCard');output_layout=QVBoxLayout(output_card);output_layout.setContentsMargins(18,16,18,18)
        heading=QLabel('报告输出');heading.setObjectName('sectionTitle');output_layout.addWidget(heading);form=QFormLayout();output_layout.addLayout(form)
        self.output_path=QLineEdit(values.output_directory);form.addRow('输出目录',self.output_path)
        browse=QPushButton('选择输出目录');browse.clicked.connect(self.choose_directory);form.addRow('',browse)
        layout.addWidget(output_card)
        general_card=QFrame();general_card.setObjectName('panelCard');general_layout=QVBoxLayout(general_card);general_layout.setContentsMargins(18,16,18,18)
        heading=QLabel('通用设置');heading.setObjectName('sectionTitle');general_layout.addWidget(heading);form=QFormLayout();general_layout.addLayout(form)
        self.completion=QComboBox()
        for label,value in [('不执行额外操作','none'),('打开输出目录','folder'),('打开最近报告','report')]:self.completion.addItem(label,value)
        self.completion.setCurrentIndex(max(0,self.completion.findData(values.completion_action)));form.addRow('生成完成后',self.completion)
        self.top=QCheckBox('窗口始终置顶');self.top.setChecked(values.always_on_top);form.addRow('',self.top);layout.addWidget(general_card)
        buttons=QDialogButtonBox(QDialogButtonBox.Save|QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Save).setText('保存');buttons.button(QDialogButtonBox.Cancel).setText('取消')
        buttons.accepted.connect(self.accept);buttons.rejected.connect(self.reject);layout.addWidget(buttons)
    def choose_directory(self):
        path=QFileDialog.getExistingDirectory(self,'选择输出目录',self.output_path.text())
        if path:self.output_path.setText(path)
    def values(self):return DesktopSettings(self.output_path.text().strip(),self.completion.currentData(),self.top.isChecked())
    def accept(self):
        if not self.output_path.text().strip():QMessageBox.warning(self,'缺少输出目录','请选择输出目录。');return
        super().accept()
