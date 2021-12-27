import json
import os
import re
import sys
import time
import threading
from PySide2.QtCore import Slot
from PySide2.QtGui import QTextCursor, QIcon, QPixmap
from PySide2.QtWidgets import QApplication, QMainWindow, QLineEdit, QTextEdit, QAbstractSpinBox, \
    QPushButton, QSpinBox, QDoubleSpinBox, QDateTimeEdit, QCheckBox, \
    QFileDialog
from PySide2.QtCore import Signal, Slot, QDateTime, QTimer, QEventLoop, QCoreApplication
from main_ui import Ui_water_mainwd
import imgs

from functools import partial


class UI(QMainWindow, Ui_water_mainwd):
    signal_log = Signal(str, bool, object)

    def __init__(self, *wd, **kw):
        Ui_water_mainwd.__init__(self)
        QMainWindow.__init__(self, parent=None)
        self.setupUi(self)
        icon = QIcon()
        icon.addPixmap(QPixmap(":res/main.ico"), QIcon.Normal, QIcon.Off)
        self.setWindowIcon(icon)

        # 所有tool button统一注册
        for k, item in self.__dict__.items():
            if isinstance(item, QPushButton):
                self.__getattribute__(k).clicked.connect(self.on_btn_clicked)
        self.dir_selector_map = {
            self.sel_output_btn: self.sel_output_lbl,
            self.sel_input_btn: self.sel_input_lbl
        }
        self.file_selector_map = {

        }
        self.file_exec_map = {
            # self.load_setting_btn: self.load_setting,
            # self.dump_setting_btn: self.dump_setting
        }

        self.signal_log.connect(self._log_msg)
        # self.clear_input_btn.hide()

        self.date_creatfile_st.setDisplayFormat("yyyy-MM-dd HH:mm:ss")
        self.date_creatfile_ed.setDisplayFormat("yyyy-MM-dd HH:mm:ss")
        self.tabWidget.setCurrentIndex(0)
        self.sel_output_btn.hide()
        self.sel_output_lbl.hide()

        self.load_input_set()

        self.__task_map = {}
        self.__task_lock = threading.RLock()

        self.start()

    def dump_input_set(self):
        dct = {}
        for k, item in self.__dict__.items():
            if isinstance(item, QLineEdit):
                dct[k] = item.text()
            elif isinstance(item, QTextEdit):
                dct[k] = item.toPlainText()
            elif isinstance(item, QSpinBox) or isinstance(item, QDoubleSpinBox):
                dct[k] = item.value()
            elif isinstance(item, QDateTimeEdit):
                dct[k] = item.text()
            elif isinstance(item, QCheckBox):
                dct[k] = item.isChecked()
        with open(".ui.dump", "w+", encoding="utf-8") as f:
            f.write(json.dumps(dct, indent=1, ensure_ascii=False))

    def load_input_set(self):
        try:
            with open(".ui.dump", "r", encoding="utf-8") as f:
                dct = json.load(f)
                if dct:
                    for k, v in dct.items():
                        if hasattr(self, k):
                            item = self.__getattribute__(k)
                            if isinstance(item, QLineEdit) \
                                    or isinstance(item, QTextEdit):
                                item.setText(v)
                            elif isinstance(item, QSpinBox) or isinstance(item, QDoubleSpinBox):
                                item.setValue(v)
                            elif isinstance(item, QDateTimeEdit):
                                item.setDateTime(QDateTime.fromString(v, "yyyy-MM-dd HH:mm:ss"))
                            elif isinstance(item, QCheckBox):
                                item.setChecked(True if v else False)
        except FileNotFoundError:
            pass
        except Exception as e:
            print(e)

    def closeEvent(self, event):
        print("main window close event")
        self.on_clear_output_btn_clicked()
        self.dump_input_set()
        # os._exit(0)
        sys.exit(0)

    @Slot()
    def on_btn_clicked(self):
        obj = self.sender()
        # 文件选择器
        if self.file_selector_map.get(obj, None) is not None:
            return self.on_common_file_choice_btn_clicked(obj, sel="file",
                                                          txt_show=self.file_selector_map.get(obj, None))

        # 文件夹选择器
        elif self.dir_selector_map.get(obj, None) is not None:
            return self.on_common_file_choice_btn_clicked(obj, sel="dir",
                                                          txt_show=self.dir_selector_map.get(obj, None))
        # 文件选择后执行回调
        elif self.file_exec_map.get(obj, None) is not None:
            return self.on_common_file_choice_btn_clicked(obj, sel="file", callback=self.file_exec_map[obj])

    def on_common_file_choice_btn_clicked(self, obj, file_filter="*.*", sel="file", txt_show=None, callback=None):
        if txt_show is not None:
            curpath = txt_show.text()
            curpath = os.path.split(curpath)[0]
            curpath = curpath if curpath else "."
        else:
            curpath = ""
        if "file" == sel:
            # file_name = QFileDialog.getOpenFileName(None, "文件选择", curpath, file_filter)
            file_name = QFileDialog.getOpenFileNames(None, "文件选择", curpath, file_filter)
        else:
            file_name = [QFileDialog.getExistingDirectory(None, "文件选择", curpath)]
        show_text = ""
        if isinstance(file_name, tuple):
            if isinstance(file_name[0], list):
                show_text = "||".join(file_name[0])
            else:
                show_text = file_name[0]
        elif isinstance(file_name, list):
            show_text = file_name[0]
        if txt_show:
            txt_show.setText(show_text)

        if callable(callback):
            for item in show_text.split("||"):
                callback(item)

    def load_setting(self, path):
        pass
        # try:
        #     with open(path, "r", encoding="utf-8") as f:
        #         dct = json.load(f)
        #         if dct:
        #             # 先清空
        #             for i in range(99):
        #                 if hasattr(self, "{0}_{1}".format("item_name", i + 1)):
        #                     self.__getattribute__("{0}_{1}".format("item_name", i + 1)).setText("")
        #                     self.__getattribute__("{0}_{1}".format("item_reg", i + 1)).setText("")
        #                     self.__getattribute__("{0}_{1}".format("item_index", i + 1)).setText("0")
        #                     self.__getattribute__("{0}_{1}".format("item_skip", i + 1)).setText("0")
        #                 else:
        #                     break
        #             for k, v in dct.items():
        #                 if hasattr(self, k):
        #                     item = self.__getattribute__(k)
        #                     if isinstance(item, QLineEdit) or isinstance(item, QTextEdit):
        #                         item.setText(v)
        #             # 更新读取字典
        #             if dct.get("items_lists", None) is not None and isinstance(dct["items_lists"], list):
        #                 for idx, item in enumerate(dct["items_lists"]):
        #                     if not isinstance(item, dict):
        #                         self.log_msg(u"{0} 读取异常".format(item))
        #                         continue
        #                     for k_name in ["item_name", "item_reg", "item_index", "item_skip"]:
        #                         if hasattr(self, "{0}_{1}".format(k_name, idx + 1)):
        #                             self.__getattribute__("{0}_{1}".format(k_name, idx + 1)).setText(
        #                                 str(item.get(k_name, "")))
        #             else:
        #                 self.log_msg(u"未发现items_lists 或其不为列表")
        #
        #         self.log_msg(u"从:{0} 读取配置成功".format(path))
        # except Exception as e:
        #     self.log_msg(str(e))


    def dump_setting(self, path):
        pass
        # try:
        #     with open(path, "w+", encoding="utf-8") as f:
        #         dct = {}
        #         for k, item in self.__dict__.items():
        #             if isinstance(item, QLineEdit):
        #                 if k.endswith("_filter"):
        #                     dct[k] = item.text()
        #         dct["items_lists"] = []
        #         idx_cnt = 1
        #         while 1:
        #             if not hasattr(self, "{0}_{1}".format("item_name", idx_cnt)):
        #                 break
        #             per_line = {}
        #             for k_name in ["item_name", "item_reg", "item_index", "item_skip"]:
        #                 if hasattr(self, "{0}_{1}".format(k_name, idx_cnt)):
        #                     per_line[k_name] = self.__getattribute__("{0}_{1}".format(k_name, idx_cnt)).text()
        #             dct["items_lists"].append(per_line)
        #             idx_cnt += 1
        #         f.write(json.dumps(dct, indent=1, ensure_ascii=False))
        #         self.log_msg(u"保存配置文件到:{0} 成功".format(path))
        # except Exception as e:
        #     self.log_msg(str(e))

    def log_msg(self, msg, mv_end=True, replace_pattern: str or re.RegexFlag = ""):
        self.signal_log.emit(msg, mv_end, replace_pattern)

    def _log_msg(self, msg, mv_end=True, replace_pattern: str or re.RegexFlag = ""):
        cursor = self.result_lbl.textCursor()
        if replace_pattern:
            ret = []
            found = False
            for item in self.result_lbl.toPlainText().split("\n"):
                if re.search(replace_pattern, item):
                    ret.append(msg)
                    found = True
                else:
                    ret.append(item)
            if not found:
                ret.append(msg)
            self.result_lbl.setText("\n".join(ret))
        else:
            cursor.insertText(msg + "\n")

        if mv_end:
            cursor.movePosition(QTextCursor.End)
            self.result_lbl.setTextCursor(cursor)

    @Slot()
    def on_clear_input_btn_clicked(self):
        for k, item in self.__dict__.items():
            if isinstance(item, QLineEdit):
                if k.startswith("item_"):
                    item.clear()

    @Slot()
    def on_clear_output_btn_clicked(self):
        self.result_lbl.clear()

    @Slot()
    def on_clear_input_btn_clicked(self):
        for i in range(self.filter_num):
            self.__getattribute__("filter_input_{0}".format(i+1)).clear()
            # self.__getattribute__("title_start_row_{0}".format(i + 1)).clear()

    def open_explore(self, drt: str):
        try:
            os.system("explorer \\e,\\root,{0}".format(drt.replace("/", os.sep)))
        except Exception as e:
            self.log_msg(str(e))

    def copy_files(self, source_dir: str, target_dir: str):
        # 拷贝文件
        from shutil import copy
        for p in os.listdir(source_dir):
            filepath = target_dir + '/' + p
            oldpath = source_dir + '/' + p
            if os.path.isdir(oldpath):
                if not os.path.exists(filepath):
                    os.mkdir(filepath)
                self.copy_files(oldpath, filepath)
            if os.path.isfile(oldpath):
                copy(oldpath, filepath)

    @Slot()
    def on_exec_btn_clicked(self):
        self.clear_output_btn.click()
        func_map = {
            self.tab_1: self.do_file_attr_modify,
            self.tab_2: self.do_file_content_modify,
            self.tab_3: self.do_gen_root_tree,
            self.tab_4: self.do_file_content_date_modify
        }
        try:
            func_map[self.tabWidget.currentWidget()]()
        except Exception as e:
            self.log_msg(str(e))
            raise e

    def do_file_attr_modify(self):
        src = self.sel_input_lbl.text().strip()
        dst = self.sel_output_lbl.text().strip()
        if dst and os.path.isdir(dst): # 拷贝所有文件到新目录
            self.log_msg("拷贝 {0} 到 {1}".format(src, dst))
            self.copy_files(src, dst)
            src = dst
        self.log_msg("开始修改目录内文件时间属性 {0}...".format(src))
        filter_lst = self.modattr_filter.toPlainText().strip()
        if filter_lst.__len__() == 0:
            filter_lst = None
        else:
            filter_lst = filter_lst.split("\n")
        time_format = "%Y-%m-%d %H:%M:%S"
        create_st = self.date_creatfile_st.text()
        create_ed = self.date_creatfile_ed.text()
        mod_range_st = float(self.date_modfile_st.text()) * 60 * 60
        mod_range_ed = float(self.date_modfile_ed.text()) * 60 * 60
        vst_range_st = float(self.date_visfile_st.text()) * 60 * 60
        vst_range_ed = float(self.date_visfile_ed.text()) * 60 * 60

        def _t():
            self.__task_lock.acquire()
            from mod_files_attr import mod_files_attr
            try:
                mod_files_attr(src, create_st, create_ed,
                               mod_range_st, mod_range_ed,
                               vst_range_st, vst_range_ed,
                               time_format=time_format,
                               logger=self.log_msg,
                               filter_list=filter_lst)
                self.log_msg("修改目录内文件时间属性成功! {0}...".format(src))
                if self.open_when_fin_btn.isChecked():
                    self.open_explore(src)
                self.__task_lock.release()
            except Exception as e:
                self.log_msg(str(e))
                self.log_msg("修改目录内文件时间属性异常! {0}...".format(src))
                self.__task_lock.release()
                raise e

        t = threading.Thread(target=_t)
        t.setDaemon(True)
        t.start()

    def do_file_content_modify(self):
        src = self.sel_input_lbl.text().strip()
        dst = self.sel_output_lbl.text().strip()
        if dst and os.path.isdir(dst): # 拷贝所有文件到新目录
            self.log_msg("拷贝 {0} 到 {1}".format(src, dst))
            self.copy_files(src, dst)
            src = dst
        self.log_msg("开始查找替换内容, 目录 {0}...".format(src))
        replace_map = {}
        for i in range(1, 999):
            if hasattr(self, "replace_input_lbl_{0}".format(i)):
                key = self.__getattribute__("replace_input_lbl_{0}".format(i)).text()
                if key:
                    val = self.__getattribute__("replace_output_lbl_{0}".format(i)).text()
                    if val:
                        replace_map[key] = val
            else:
                break
        filter_lst = []
        if self.replacefile_filter.toPlainText().strip():
            filter_lst = self.replacefile_filter.toPlainText().strip().split('\n')
            for i in range(filter_lst.__len__()):
                if "\." not in filter_lst[i]:
                    filter_lst[i] = "\." + filter_lst[i]

        def _t():
            self.__task_lock.acquire()
            try:
                from document_edit import replace_dir, office2officeX
                office2officeX(src, logger=self.log_msg)
                tt_cnt = replace_dir(src, replace_dct=replace_map, filter_lst=filter_lst, logger=self.log_msg)
                self.log_msg("完成！总替换次数：【{0}】， 目录：{1}".format(tt_cnt, src))
                if self.open_when_fin_btn.isChecked():
                    self.open_explore(src)
                self.__task_lock.release()
            except Exception as e:
                self.log_msg(str(e))
                self.__task_lock.release()
                raise e
        t = threading.Thread(target=_t)
        t.setDaemon(True)
        t.start()

    def do_gen_root_tree(self):
        dst = self.sel_output_lbl.text().strip()
        if not dst:
            dst = self.sel_input_lbl.text().strip()
        dst = dst + os.sep + "目录结构.docx"
        self.log_msg("开始生成目录树到 {0}...".format(dst))
        src = self.sel_input_lbl.text().strip()
        ignore_list = [r"~.*\..*"]  # 默认不列出win32下的临时文件
        if self.rootshow_filter.toPlainText():
            ignore_list += self.rootshow_filter.toPlainText().split('\n')
        assert src
        from list_tree import list_tree
        list_tree(src, dst, ignore_lst=ignore_list,
                  logger=self.log_msg, use_to_word=self.radioButton_toword.isChecked())
        self.log_msg("生成目录树到 {0} 成功！".format(dst))
        if self.open_when_fin_btn.isChecked():
            self.open_explore(os.path.split(dst)[0])

    def do_file_content_date_modify(self):
        src = self.sel_input_lbl.text().strip()
        dst = self.sel_output_lbl.text().strip()
        if dst and os.path.isdir(dst):  # 拷贝所有文件到新目录
            self.log_msg("拷贝 {0} 到 {1}".format(src, dst))
            self.copy_files(src, dst)
            src = dst
        self.log_msg("开始查找日期替换内容, 目录 {0}...".format(src))
        replace_call_lst = []
        from document_edit import replace_date
        for i in range(1, 999):
            if hasattr(self, "input_year_{0}".format(i)):
                y = self.__getattribute__("input_year_{0}".format(i)).text()
                m = self.__getattribute__("input_month_{0}".format(i)).text()
                n_y = self.__getattribute__("output_year_{0}".format(i)).text()
                n_m = self.__getattribute__("output_month_{0}".format(i)).text()
                is_enable = self.__getattribute__("date_enable_{0}".format(i)).isChecked()
                if is_enable and y and m and n_y and n_m:
                    # 将年月输入转为正则表达式
                    i_y = int(y)
                    i_m = int(m)
                    i_n_y = int(n_y)
                    i_n_m = int(n_m)
                    src_date = [i_y, i_m]
                    dst_date = [i_n_y, i_n_m]
                    fun = partial(replace_date, ori_date=src_date,
                                  dst_date=dst_date,
                                  fmt_index_used=(0, 1),
                                  for_excel_sp=True)
                    replace_call_lst.append(fun)
            else:
                break
        filter_lst = []
        if self.replacefile_filter.toPlainText().strip():
            filter_lst = self.replacedate_filter.toPlainText().strip().split('\n')
            for i in range(filter_lst.__len__()):
                if "\." not in filter_lst[i]:
                    filter_lst[i] = "\." + filter_lst[i]

        def _t():
            self.__task_lock.acquire()
            try:
                from document_edit import replace_dir, office2officeX
                office2officeX(src, logger=self.log_msg)
                tt_cnt = replace_dir(src, replace_dct={}, filter_lst=filter_lst, logger=self.log_msg,
                                     replace_cb_lst=replace_call_lst)
                self.log_msg("完成！总替换次数：【{0}】， 目录：{1}".format(tt_cnt, src))
                if self.open_when_fin_btn.isChecked():
                    self.open_explore(src)
                self.__task_lock.release()
            except Exception as e:
                self.log_msg(str(e))
                self.__task_lock.release()
                raise e

        t = threading.Thread(target=_t)
        t.setDaemon(True)
        t.start()

    @Slot()
    def on_year_input_to_all_clicked(self):
        start_year = 0
        start_out_year = 0
        for i in range(1, 999):
            if hasattr(self, "input_year_{0}".format(i)):
                is_enable = self.__getattribute__("date_enable_{0}".format(i)).isChecked()
                if is_enable:
                    if 0 == start_year:
                        start_year = int(self.__getattribute__("input_year_{0}".format(i)).text())
                        start_out_year = int(self.__getattribute__("output_year_{0}".format(i)).text())
                    else:
                        self.__getattribute__("input_year_{0}".format(i)).setValue(start_year)
                        self.__getattribute__("output_year_{0}".format(i)).setValue(start_out_year)
            else:
                break

    @Slot()
    def on_month_asc_input_clicked(self):
        start_month = 0
        start_out_month = 0
        for i in range(1, 999):
            if hasattr(self, "input_year_{0}".format(i)):
                is_enable = self.__getattribute__("date_enable_{0}".format(i)).isChecked()
                if is_enable:
                    if 0 == start_month:
                        start_month = int(self.__getattribute__("input_month_{0}".format(i)).text())
                        start_out_month = int(self.__getattribute__("output_month_{0}".format(i)).text())
                    else:
                        start_month = (start_month % 12) + 1
                        start_out_month = (start_out_month % 12) + 1
                        self.__getattribute__("input_month_{0}".format(i)).setValue(start_month)
                        self.__getattribute__("output_month_{0}".format(i)).setValue(start_out_month)
            else:
                break

    @Slot()
    def on_enable_all_date_clicked(self):
        for i in range(1, 999):
            if hasattr(self, "input_year_{0}".format(i)):
                is_enable = self.__getattribute__("date_enable_{0}".format(i)).isChecked()
                if is_enable:
                    self.__getattribute__("date_enable_{0}".format(i)).setChecked(False)
                else:
                    self.__getattribute__("date_enable_{0}".format(i)).setChecked(True)
            else:
                break

    @Slot()
    def on_exec_all_btn_clicked(self):
        wdg_lsg = [1, 3, 2, 0]
        func_map = [
            self.do_file_attr_modify,
            self.do_file_content_modify,
            self.do_gen_root_tree,
            self.do_file_content_date_modify]
        cnt = 0

        def _t():
            nonlocal cnt, wdg_lsg
            is_open = self.open_when_fin_btn.isChecked()
            self.open_when_fin_btn.setChecked(False)
            for index in wdg_lsg:
                try:
                    self.__task_lock.acquire()
                    self.tabWidget.setCurrentIndex(index)

                    self.__task_lock.release()
                    self.log_msg("****************************************************")
                    cnt += 1
                    self.log_msg("【执行任务】[{0}]:{1} ".format(cnt, self.tabWidget.tabText(index)))
                    self.log_msg("****************************************************")
                    func_map[index]()
                    time.sleep(2)
                    self.__task_lock.acquire()
                    self.log_msg("****************************************************")
                    self.log_msg("【完成任务】[{0}]:{1} !".format(cnt, self.tabWidget.tabText(index)))
                    self.log_msg("****************************************************")
                    self.__task_lock.release()
                except Exception as e:
                    self.log_msg(str(e))
                    self.__task_lock.release()
                    raise e
            if is_open:
                self.open_when_fin_btn.setChecked(True)
                self.open_explore(self.sel_input_lbl.text().strip())

        t = threading.Thread(target=_t)
        t.setDaemon(True)
        t.start()

    def _routine(self):
        while True:
            # self.__task_lock.acquire()
            for item in list(self.__task_map.items()):
                pass
                # if item is not None:
                #     k, v = item
                # else:
                #     continue
                # if not isinstance(v, Task):
                #     continue
                #
                # raw_pct = v.get_progress()
                # raw_pct = raw_pct if raw_pct < 1 else 1
                # pct = int(35 * raw_pct)
                # p = "▋" * pct + " " * (35 - pct)
                # self.log_msg("{0} 分析进度:{1} \t{2:.2f}%\n".format(k, p, 100*pct/35.0),
                #              mv_end=True, replace_pattern="{0} 分析进度".format(k))
                # if v.is_done():
                #     if not v.fault_msg:
                #         self.log_msg("{0} 分析完成，写入行数:{1}".format(k, v.get_write_len()))
                #     else:
                #         self.log_msg("发生错误 " + v.fault_msg + "\n")
                #     self.__task_map.pop(k)

            # self.__task_lock.release()
            QCoreApplication.processEvents(QEventLoop.AllEvents, 100)
            time.sleep(1)

    def start(self):
        t = threading.Thread(target=self._routine)
        t.setDaemon(True)
        t.start()


if __name__ == "__main__":
    uiapp = QApplication([])
    a = UI()
    a.show()
    sys.exit(uiapp.exec_())
