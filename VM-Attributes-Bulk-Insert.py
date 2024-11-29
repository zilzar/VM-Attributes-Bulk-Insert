

import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QProgressBar, QTextEdit
)
from PyQt5 import QtCore
from pyVim.connect import SmartConnect, Disconnect
from pyVmomi import vim
import ssl

def connect_to_vcenter(host, user, password):
    try:
        context = ssl._create_unverified_context()
        si = SmartConnect(host=host, user=user, pwd=password, sslContext=context)
        return si
    except Exception as e:
        return str(e)

def find_vm_by_name(folder, vm_name):
    for entity in folder.childEntity:
        if isinstance(entity, vim.VirtualMachine) and entity.name == vm_name:
            return entity
        elif hasattr(entity, 'childEntity'):  
            vm = find_vm_by_name(entity, vm_name)
            if vm:
                return vm
    return None

def process_excel_and_add_attributes(excel_file, si, progress_bar, log_text_edit):
    try:
        df = pd.read_excel(excel_file, dtype=str).fillna('')  
        total_vms = len(df)
        progress_bar.setMaximum(total_vms)
        progress_bar.setValue(0)

        for index, row in df.iterrows():
            vm_name = row['VM Name']
            attributes = row.drop(labels=['VM Name']).to_dict()

            content = si.RetrieveContent()
            vm = None

            for datacenter in content.rootFolder.childEntity:
                if isinstance(datacenter, vim.Datacenter):
                    vm_folder = datacenter.vmFolder
                    vm = find_vm_by_name(vm_folder, vm_name)
                    if vm:
                        break

            if vm is None:
                log_text_edit.append(f"VM '{vm_name}' not found!")
                continue

            for key, value in attributes.items():
                log_text_edit.append(f"'{vm.name}' Assigning '{value}' to '{key}' field for VM...")
                try:
                    vm.SetCustomValue(key=key, value=value)
                except Exception as e:
                    log_text_edit.append(f"'{vm.name}' Error occurred while adding Custom Attributes to VM: {e}")
            
            progress_bar.setValue(index + 1)
            QApplication.processEvents() 

        log_text_edit.append("Process Completed!")
    except Exception as e:
        log_text_edit.append(f"Error occurred while processing Excel file: {e}")


class VCenterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("vCenter Credantials Entry")
        self.layout = QVBoxLayout()

        self.layout.addWidget(QLabel("vCenter FQDN:"))
        self.vcenter_host = QLineEdit()
        self.layout.addWidget(self.vcenter_host)

        self.layout.addWidget(QLabel("User Name:"))
        self.vcenter_user = QLineEdit()
        self.layout.addWidget(self.vcenter_user)

        self.layout.addWidget(QLabel("Password:"))
        self.vcenter_password = QLineEdit()
        self.vcenter_password.setEchoMode(QLineEdit.Password)
        self.layout.addWidget(self.vcenter_password)

        self.layout.addWidget(QLabel("Excel File:"))
        self.excel_file_path = QLineEdit()
        self.layout.addWidget(self.excel_file_path)
        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.browse_file)
        self.layout.addWidget(self.browse_button)


        self.submit_button = QPushButton("Connect and Embed")
        self.submit_button.clicked.connect(self.submit)
        self.layout.addWidget(self.submit_button)


        self.progress_bar = QProgressBar()
        self.layout.addWidget(self.progress_bar)


        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)
        self.layout.addWidget(self.log_text_edit)


        self.signature_label = QLabel("Zil Zar", self)
        self.signature_label.setAlignment(QtCore.Qt.AlignCenter)  
        self.layout.addWidget(self.signature_label)

        self.setLayout(self.layout)

    def browse_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Select excel file", "", "Excel Files (*.xlsx)")
        if file_path:
            self.excel_file_path.setText(file_path)

    def submit(self):
        vcenter_host = self.vcenter_host.text()
        vcenter_user = self.vcenter_user.text()
        vcenter_password = self.vcenter_password.text()
        excel_file = self.excel_file_path.text()

        if not vcenter_host or not vcenter_user or not vcenter_password or not excel_file:
            QMessageBox.warning(self, "Missing Information", "Please fill in all fields!")
            return


        si = connect_to_vcenter(vcenter_host, vcenter_user, vcenter_password)
        if isinstance(si, str): 
            QMessageBox.critical(self, "Connection Error", f"Error connecting to vCenter: {si}")
        else:
            QMessageBox.information(self, "Success", "Connected to vCenter successfully.")
            try:
                process_excel_and_add_attributes(excel_file, si, self.progress_bar, self.log_text_edit)
                QMessageBox.information(self, "Success", "Data in Excel file was processed.")
            except Exception as e:
                QMessageBox.critical(self, "Processing Error", f"An error occurred while processing the data: {e}")
            finally:
                Disconnect(si)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VCenterApp()
    window.show()
    sys.exit(app.exec_())
