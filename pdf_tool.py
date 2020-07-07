#!/usr/bin/env python3
# -*- coding: utf-8 -*

import sys
import subprocess
import re
from pathlib import Path

from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize, Qt


class PdfTool(QtWidgets.QDialog):
    """Main Window containing the three tabs 'Compress', 'Split' and 'Merge'.
    """
    def __init__(self):
        super().__init__(parent=None)
        self.setWindowTitle('Pdf Tool')
        self.vertical_layout = QtWidgets.QVBoxLayout()
        self.tab_widget = QtWidgets.QTabWidget()
        self.tab_widget.addTab(TabCompress(), 'Compress')
        self.tab_widget.addTab(TabSplit(), 'Split')
        self.tab_widget.addTab(TabMerge(), 'Merge')
        self.vertical_layout.addWidget(self.tab_widget)
        self.setLayout(self.vertical_layout)

    @staticmethod
    def get_all_files(folder):
        """Returns a list of all pdf files existing in the given folder.
        """
        file_list = []
        for item in folder.iterdir():
            if item.is_file() and item.suffix == '.pdf':
                file_list.append(item)
        return file_list

    @staticmethod
    def refresh_list_widget(file_list, widget):
        """Refresh the given list widget with the given list.
        """
        file_list = list(dict.fromkeys(file_list))
        widget.clear()
        widget.addItems([str(file) for file in file_list])

    @staticmethod
    def remove_file(file_list, widget):
        """Removes selected item in given list and given list widget.
        """
        try:
            selected_item = widget.selectedItems()[0]
            file_list.remove(Path(selected_item.text()))
            widget.takeItem(widget.row(selected_item))
        except IndexError:
            pass


class TabCompress(QtWidgets.QWidget):
    """Tab containing the elements for pdf compression.
    """
    def __init__(self):
        super().__init__()

        self.horizontal_layout = QtWidgets.QHBoxLayout(self)
        self.horizontal_layout.setContentsMargins(10, 10, 10, 10)
        self.horizontal_layout.setSpacing(10)
        self.file_dialog_input = QtWidgets.QFileDialog()
        self.folder_dialog_output = QtWidgets.QFileDialog()

        self.folder_dialog = QtWidgets.QFileDialog()
        self.file_list = []
        self.output_path = Path().home()
        self.file_list_widget = QtWidgets.QListWidget()
        self.entry_output_path = QtWidgets.QLineEdit()
        self.entry_output_path.setText(str(self.output_path))

        self.make_layout_compress()

    def make_layout_compress(self):
        """Create and arrange the layout for the compression elements.
        """
        vertical_layout_compress = QtWidgets.QVBoxLayout()
        vertical_layout_compress.setContentsMargins(10, 10, 10, 10)
        vertical_layout_compress.setSpacing(10)
        self.horizontal_layout.addLayout(vertical_layout_compress)

        push_button_load_files_input = QtWidgets.QPushButton()
        push_button_load_files_input.setMinimumSize(QSize(100, 30))
        push_button_load_files_input.setText('Add pdf files')
        push_button_load_files_input.clicked.connect(self.open_file_dialog_input)
        push_button_load_folder_input = QtWidgets.QPushButton()
        push_button_load_folder_input.setMinimumSize(QSize(100, 30))
        push_button_load_folder_input.setText('Add all pdf files from a folder')
        push_button_load_folder_input.clicked.connect(self.open_folder_dialog_input)
        label_list_widget = QtWidgets.QLabel('Pdf files to compress:')
        push_button_remove_file = QtWidgets.QPushButton()
        push_button_remove_file.setText('Remove selected file')
        push_button_remove_file.clicked.connect(self.remove_file)
        push_button_choose_path_output = QtWidgets.QPushButton()
        push_button_choose_path_output.setMinimumSize(QSize(100, 30))
        push_button_choose_path_output.setText('Change output path')
        push_button_choose_path_output.clicked.connect(self.open_folder_dialog_output)

        push_button_start_compress = QtWidgets.QPushButton()
        push_button_start_compress.setText('Start compression')
        push_button_start_compress.clicked.connect(self.start_compression)

        vertical_layout_compress.addWidget(push_button_load_files_input)
        vertical_layout_compress.addWidget(push_button_load_folder_input)
        vertical_layout_compress.addWidget(label_list_widget)
        vertical_layout_compress.addWidget(self.file_list_widget)
        vertical_layout_compress.addWidget(push_button_remove_file)
        vertical_layout_compress.addWidget(push_button_choose_path_output)
        vertical_layout_compress.addWidget(self.entry_output_path)
        vertical_layout_compress.addWidget(push_button_start_compress)

    def start_compression(self):
        """Start the compression process by calling self.run_gs(). Opens messagebox when finished.
        """
        self.output_path = Path(self.entry_output_path.text())
        if self.check_if_output_is_valid_and_different_to_input(self.file_list, self.output_path):
            for file in self.file_list:
                TabCompress.run_gs(str(file), str(self.output_path / file.name))
            message_box = QtWidgets.QMessageBox(self)
            message_box.setText('Compression finished!')
            message_box.show()

    @staticmethod
    def run_gs(input_file, output_file):
        """Runs the tool ghostscript to compress the given pdf file. Takes strings for the input
        and the output file as arguments.
        """
        command = ('gs', '-sDEVICE=pdfwrite', '-dNOPAUSE', '-dBATCH', f'-sOutputFile={output_file}', input_file)
        subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

    def check_if_output_is_valid_and_different_to_input(self, input_file_list, output_path):
        """Returns True if the given output path is valid and different to all paths in the given list of input files.
        Returns False otherwise.
        """
        if input_file_list:
            if output_path.root and output_path.is_dir():
                for file in input_file_list:
                    if file.parent == output_path:
                        message_box = QtWidgets.QMessageBox(self)
                        message_box.setText('Output path should be different to the destination of your input files!')
                        message_box.show()
                        return False

                return True
            else:
                message_box = QtWidgets.QMessageBox(self)
                message_box.setText('No valid output path selected!')
                message_box.show()
                return False
        else:
            message_box = QtWidgets.QMessageBox(self)
            message_box.setText('No input files selected')
            message_box.show()
            return False

    def open_file_dialog_input(self):
        """Opens the file dialog to choose the input file(s). Writes its value(s) to self.file_list.
        """
        self.file_dialog_input.setFileMode(QtWidgets.QFileDialog.ExistingFiles)
        self.file_dialog_input.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)
        file_list_temp = self.file_dialog_input.getOpenFileNames(self, 'Select pdf files to compress!', '', 'Pdf files (*.pdf)')[0]
        if file_list_temp:
            for file in file_list_temp:
                self.file_list.append(Path(file))
            PdfTool.refresh_list_widget(self.file_list, self.file_list_widget)

    def open_folder_dialog_output(self):
        """Opens the folder dialog to choose the destination of the output files. Writes its value to self.output_path.
        """
        path = self.folder_dialog_output.getExistingDirectory(self, 'Change output folder!')
        if path:
            self.entry_output_path.setText(path)
            self.output_path = Path(path)

    def open_folder_dialog_input(self):
        """Opens the folder dialog to choose the folder containing the input files.
        Writes its value to self.file_list via the method PdfTool.get_all_files.
        """
        self.folder_dialog.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)
        folder = Path(self.folder_dialog.getExistingDirectory(self, 'Select folder!'))
        if folder.root:
            self.file_list += PdfTool.get_all_files(folder)
            PdfTool.refresh_list_widget(self.file_list, self.file_list_widget)

    def remove_file(self):
        """Call PdfTool.remove_file to remove selected file from list and widget.
        """
        PdfTool.remove_file(self.file_list, self.file_list_widget)


class TabSplit(QtWidgets.QWidget):
    """Tab containing the elements for pdf splitting.
    """
    def __init__(self):
        super().__init__()

        self.horizontal_layout = QtWidgets.QHBoxLayout(self)
        self.horizontal_layout.setContentsMargins(10, 10, 10, 10)
        self.horizontal_layout.setSpacing(10)
        self.file_dialog_input = QtWidgets.QFileDialog()
        self.folder_dialog_output = QtWidgets.QFileDialog()
        self.file = Path()
        self.label_file = QtWidgets.QLabel('Select a pdf file!')
        self.output_filename_line_edit = QtWidgets.QLineEdit()
        self.output_filename_line_edit.textChanged.connect(self.refresh_output_label)
        self.label_output_path = QtWidgets.QLabel()
        self.output_path = Path().home()
        self.line_edit_split_pattern = QtWidgets.QLineEdit('1-2,7-9')
        self.compress_radio_button = QtWidgets.QRadioButton()
        self.compress_radio_button.setText('Compress output pdf file?')
        self.compress_radio_button.setChecked(True)
        self.make_layout_split()

    def make_layout_split(self):
        """Create and arrange the layout for the pdf splitting elements.
        """
        vertical_layout_split = QtWidgets.QVBoxLayout()
        vertical_layout_split.setContentsMargins(10, 10, 10, 10)
        vertical_layout_split.setSpacing(10)
        self.horizontal_layout.addLayout(vertical_layout_split)
        push_button_load_files_input = QtWidgets.QPushButton('Load pdf file')
        push_button_load_files_input.clicked.connect(self.open_file_dialog_input)
        label_split_pattern = QtWidgets.QLabel('Split pattern:')
        push_button_start_splitting = QtWidgets.QPushButton('Start Splitting')
        push_button_start_splitting.clicked.connect(self.start_splitting)

        push_button_choose_path_output = QtWidgets.QPushButton('Change output path')
        push_button_choose_path_output.clicked.connect(self.open_folder_dialog_output)

        label_filename = QtWidgets.QLabel('Type name of the output file:')
        vertical_layout_split.addWidget(push_button_load_files_input)
        vertical_layout_split.addWidget(self.label_file)
        vertical_layout_split.addWidget(label_split_pattern)
        vertical_layout_split.addWidget(self.line_edit_split_pattern)
        vertical_layout_split.addWidget(push_button_choose_path_output)
        horizontal_layout_filename = QtWidgets.QHBoxLayout()
        horizontal_layout_filename.addWidget(label_filename)
        self.output_filename_line_edit.setText('output')
        horizontal_layout_filename.addWidget(self.output_filename_line_edit)
        vertical_layout_split.addLayout(horizontal_layout_filename)
        vertical_layout_split.addWidget(self.label_output_path)

        horizontal_layout_bottom = QtWidgets.QHBoxLayout()
        horizontal_layout_bottom.addWidget(self.compress_radio_button)
        horizontal_layout_bottom.addWidget(push_button_start_splitting)
        vertical_layout_split.addLayout(horizontal_layout_bottom)

    def open_file_dialog_input(self):
        """Opens the file dialog to choose the input file. Writes its value to self.file.
        """
        self.file_dialog_input.setFileMode(QtWidgets.QFileDialog.ExistingFile)
        self.file_dialog_input.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)
        self.file = self.file_dialog_input.getOpenFileName(self, 'Select pdf file to split!', '', 'Pdf files (*.pdf)')[0]
        if self.file:
            self.label_file.setText(f'Selected pdf file:   {self.file}')

    def open_folder_dialog_output(self):
        """Opens the folder dialog to choose the destination of the output files. Writes its value to self.output_path.
        """
        path = self.folder_dialog_output.getExistingDirectory(self, 'Select output folder!')
        if path:
            self.label_output_path.setText(f'Output File:    {path}/{self.output_filename_line_edit.text()}.pdf')
            self.output_path = Path(path)

    def refresh_output_label(self):
        """Refresh output label to selected output path.
        """
        self.label_output_path.setText(f'Output File:    {self.output_path}/{self.output_filename_line_edit.text()}.pdf')

    def start_splitting(self):
        """Starts splitting process. Informs when finished or the split pattern has a wrong format.
        """
        list_start_stop = TabSplit.analyze_split_pattern(self.line_edit_split_pattern.text())
        list_indices = []
        output_file = f'{self.output_path}/{self.output_filename_line_edit.text()}.pdf'
        if list_start_stop:
            for item in list_start_stop:
                split_succeeded = self.split_pdf(*item, self.file, output_file)
                if not split_succeeded:
                    return
                list_indices += [n for n in range(int(item[0]), int(item[1]) + 1)]
            command = ['pdfunite']
            for index in list_indices:
                command.append(output_file + str(index))
            command.append(output_file)
            subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

            for index in list_indices:
                Path(output_file + str(index)).unlink()
            if self.compress_radio_button.toggled:
                TabCompress.run_gs(output_file, output_file + '_')
                Path(output_file + '_').rename(Path(output_file))

            message_box = QtWidgets.QMessageBox(self)
            message_box.setText('Splitting finished.')
            message_box.show()
        else:
            message_box = QtWidgets.QMessageBox(self)
            message_box.setText('Wrong split format. Example: 1, 2, 4-6, 8-9')
            message_box.show()

    @staticmethod
    def analyze_split_pattern(string_split_pattern):
        """Takes the split pattern string input as argument. Returns a list with of the
        list [start-page, stop-page] for each element seperated by ',' of the input string.
        Returns False if the elements are not in the right format: int, or int-int.
        """
        list_old = string_split_pattern.replace(' ', '').split(',')
        list_new = []
        r = re.compile('[0-9][0-9]*-[0-9][0-9]*')
        for item in list_old:
            if r.match(item) is not None:
                list_new.append(item.split('-'))
            elif item.isnumeric():
                list_new.append([item, item])
            else:
                return False

        return list_new

    def split_pdf(self, start, stop, input_file, output_file):
        """Start single splitting process with tool pdfseperate. Takes start page, stop page, input file and output file in
        string format as arguments. Returns True if successful, False otherwise.
        """
        command = ['pdfseparate', '-f', start, '-l', stop, input_file, f'{output_file}%d']
        with subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT) as proc:
            list_log_split = proc.stdout.readlines()
            try:
                log_split = list_log_split[0]
            except IndexError:
                log_split = b''
            if b'Illegal pageNo' in log_split:
                page_string = log_split.strip()[-4:].decode()
                message_box = QtWidgets.QMessageBox(self)
                message_box.setText(f'Page {page_string[0]} doesn\'t exist. The pdf file only contains {page_string[2]} pages.')
                message_box.show()
                return False
        return True


class TabMerge(QtWidgets.QWidget):
    """Tab containing the elements for pdf merging.
    """
    def __init__(self):
        super().__init__()
        self.setMinimumWidth(600)
        self.horizontal_layout = QtWidgets.QHBoxLayout(self)
        self.horizontal_layout.setContentsMargins(10, 10, 10, 10)
        self.horizontal_layout.setSpacing(10)
        self.file_dialog_input = QtWidgets.QFileDialog()
        self.folder_dialog_output = QtWidgets.QFileDialog()

        self.folder_dialog = QtWidgets.QFileDialog()
        self.file_list = []
        self.output_filename_line_edit = QtWidgets.QLineEdit()
        self.output_filename_line_edit.textChanged.connect(self.refresh_output_label)
        self.label_output_path = QtWidgets.QLabel()
        self.output_path = Path().home()
        self.file_list_widget = QtWidgets.QListWidget()

        self.compress_radio_button = QtWidgets.QRadioButton()
        self.compress_radio_button.setText('Compress output pdf file?')
        self.compress_radio_button.setChecked(True)
        self.make_layout_merge()

    def make_layout_merge(self):
        """Create and arrange the layout for the pdf merging elements.
        """
        vertical_layout_merge = QtWidgets.QVBoxLayout()
        vertical_layout_merge.setContentsMargins(10, 10, 10, 10)
        vertical_layout_merge.setSpacing(10)
        self.horizontal_layout.addLayout(vertical_layout_merge)
        push_button_load_files_input = QtWidgets.QPushButton()
        push_button_load_files_input.setMinimumSize(QSize(100, 30))
        push_button_load_files_input.setText('Add pdf files')
        push_button_load_files_input.clicked.connect(self.open_file_dialog_input)
        push_button_load_folder_input = QtWidgets.QPushButton()
        push_button_load_folder_input.setMinimumSize(QSize(100, 30))
        push_button_load_folder_input.setText('Add all pdf files from a folder')
        push_button_load_folder_input.clicked.connect(self.open_folder_dialog_input)
        push_button_up = QtWidgets.QPushButton()
        push_button_up.setIcon(QIcon.fromTheme('arrow-up'))
        push_button_up.clicked.connect(self.move_selected_item_up)
        label_list_widget = QtWidgets.QLabel('Pdf files to merge:')
        push_button_down = QtWidgets.QPushButton(self)
        push_button_down.setIcon(QIcon.fromTheme('arrow-down'))
        push_button_down.clicked.connect(self.move_selected_item_down)
        push_button_remove_file = QtWidgets.QPushButton(self)
        push_button_remove_file.setIcon(QIcon.fromTheme('remove'))
        push_button_remove_file.clicked.connect(self.remove_file)
        push_button_choose_path_output = QtWidgets.QPushButton('Change output path')
        push_button_choose_path_output.clicked.connect(self.open_folder_dialog_output)
        label_filename = QtWidgets.QLabel('Type name of the output file:')

        push_button_start_merge = QtWidgets.QPushButton()
        push_button_start_merge.setText('Merge selected pdf files to one single file')
        push_button_start_merge.clicked.connect(self.start_merge)

        vertical_layout_merge.addWidget(push_button_load_files_input)
        vertical_layout_merge.addWidget(push_button_load_folder_input)
        vertical_layout_merge.addWidget(label_list_widget)

        horizontal_layout_file_list = QtWidgets.QHBoxLayout()
        vertical_layout_buttons = QtWidgets.QVBoxLayout()
        vertical_layout_buttons.addWidget(push_button_up)
        vertical_layout_buttons.addWidget(push_button_remove_file)
        vertical_layout_buttons.addWidget(push_button_down)

        horizontal_layout_file_list.addWidget(self.file_list_widget)
        horizontal_layout_file_list.addLayout(vertical_layout_buttons)
        vertical_layout_merge.addLayout(horizontal_layout_file_list)
        vertical_layout_merge.addWidget(push_button_choose_path_output)
        horizontal_layout_filename = QtWidgets.QHBoxLayout()
        horizontal_layout_filename.addWidget(label_filename)
        self.output_filename_line_edit.setText('output')
        horizontal_layout_filename.addWidget(self.output_filename_line_edit)
        vertical_layout_merge.addLayout(horizontal_layout_filename)
        vertical_layout_merge.addWidget(self.label_output_path)

        horizontal_layout_bottom = QtWidgets.QHBoxLayout()
        horizontal_layout_bottom.addWidget(self.compress_radio_button)
        horizontal_layout_bottom.addWidget(push_button_start_merge)
        vertical_layout_merge.addLayout(horizontal_layout_bottom)

    def refresh_output_label(self):
        """Refresh output label to selected output path.
        """
        file_name = self.output_filename_line_edit.text()
        if file_name[-4:] == '.pdf':
            output_file = file_name[:-4]
        self.label_output_path.setText(f'Output File:     {self.output_path}/{file_name}.pdf')

    def move_selected_item_up(self):
        """Moves the position of the selected item in the list widget and related list up.
        """
        if self.file_list:
            current_row = self.file_list_widget.currentRow()
            current_item = self.file_list_widget.takeItem(current_row)
            self.file_list.insert(current_row - 1, self.file_list.pop(current_row))
            self.file_list_widget.insertItem(current_row - 1, current_item)
            self.file_list_widget.setCurrentRow(current_row - 1)

    def move_selected_item_down(self):
        """Moves the position of the selected item in the list widget and related list down.
        """
        if self.file_list:
            current_row = self.file_list_widget.currentRow()
            current_item = self.file_list_widget.takeItem(current_row)
            self.file_list.insert(current_row + 1, self.file_list.pop(current_row))
            self.file_list_widget.insertItem(current_row + 1, current_item)
            self.file_list_widget.setCurrentRow(current_row + 1)

    def open_file_dialog_input(self):
        """Opens the file dialog to choose the input file(s). Writes its value(s) to self.file_list.
        """
        self.file_dialog_input.setFileMode(QtWidgets.QFileDialog.ExistingFiles)
        self.file_dialog_input.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)
        file_list_temp = self.file_dialog_input.getOpenFileNames(self, 'Select pdf files to compress!', '', 'Pdf files (*.pdf)')[0]
        if file_list_temp:
            for file in file_list_temp:
                self.file_list.append(Path(file))
            PdfTool.refresh_list_widget(self.file_list, self.file_list_widget)

    def open_folder_dialog_output(self):
        """Opens the folder dialog to choose the destination of the output file. Writes its value to self.output_path.
        """
        path = self.folder_dialog_output.getExistingDirectory(self, 'Select output folder!')
        if path:
            self.label_output_path.setText(f'Output File:     {path}/{self.output_filename_line_edit.text()}.pdf')
            self.output_path = Path(path)

    def open_folder_dialog_input(self):
        """Opens the folder dialog to choose the folder containing the input files.
        Writes its value to self.file_list via the method PdfTool.get_all_files.
        """
        self.folder_dialog.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)
        folder = Path(self.folder_dialog.getExistingDirectory(self, 'Select folder!'))
        if folder.root:
            self.file_list += PdfTool.get_all_files(folder)
            PdfTool.refresh_list_widget(self.file_list, self.file_list_widget)

    def remove_file(self):
        """Call PdfTool.remove_file to remove selected file from list and widget.
        """
        PdfTool.remove_file(self.file_list, self.file_list_widget)

    def start_merge(self):
        """Start merging process with the tool pdfunite. Informs when finished or no input or output file is given.
        """
        message_box = QtWidgets.QMessageBox(self)

        if self.output_filename_line_edit.text():
            if self.file_list:
                output_file = str(self.output_path / self.output_filename_line_edit.text())
                if output_file[-4:] == '.pdf':
                    output_file = output_file[:-4]
                command = ['pdfunite'] + self.file_list + [output_file + '.pdf']
                subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
                if self.compress_radio_button.toggled:
                    TabCompress.run_gs(output_file + '.pdf', output_file + '_.pdf')
                    Path(output_file + '_.pdf').rename(Path(output_file + '.pdf'))
                message_box.setText('Emerging finished')
                message_box.show()
            else:
                message_box.setText('No pdf files selected!')
                message_box.show()
        else:
            message_box.setText('Choose a file name')
            message_box.show()

def main():
    app = QtWidgets.QApplication(sys.argv)
    main.pdf_tool = PdfTool()
    main.pdf_tool.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()