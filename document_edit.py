#!/usr/bin/env python
# -*- coding:utf-8 -*-
import os
import time

import win32com
from win32com.client import Dispatch
from win32com import client as wc
import pythoncom
import openpyxl
import re
# import pptx
import docx
from pptx import Presentation
from typing import List, Dict, Tuple
import datetime


class EditDocx:
    def __init__(self, filename=None):
        self.doc = None
        self.filename = None
        if filename:
            self.filename = filename
            if os.path.exists(self.filename):
                self.doc = docx.Document(filename)
            else:
                raise NotImplemented
        else:
            raise NotImplemented

    def __del__(self):
        if self.doc:
            pass

    @staticmethod
    def _replace_text(text_frame, replace_dct: Dict, replace_cb_lst: List = None):
        cnt = 0
        for paragraph in text_frame.paragraphs:
            val = paragraph.text
            if replace_cb_lst:
                for func in replace_cb_lst:
                    val, is_change = func(val)
                    if is_change:
                        paragraph.text = val
                        cnt += 1
            else:
                for string, new_string in replace_dct.items():
                    res = re.search(string, val)
                    if res:
                        val = re.sub(string, new_string, val)
                        paragraph.text = val
                        cnt += 1

        return cnt

    @staticmethod
    def _replace_table(table, replace_dct: Dict, replace_cb_lst: List = None):
        cnt = 0
        for row in range(table.rows.__len__()):
            for cell in table.rows[row].cells:
                val = cell.text
                print(val, replace_dct)
                if replace_cb_lst:
                    for func in replace_cb_lst:
                        val, is_change = func(val)
                        if is_change:
                            cell.text = val
                            cnt += 1
                else:
                    for string, new_string in replace_dct.items():
                        print(string, val)
                        res = re.search(string, val)
                        if res:
                            val = re.sub(string, new_string, val)
                            cell.text = val
                            cnt += 1

        return cnt

    def replace_str(self, replace_dct: Dict):
        """????????????"""
        cnt = 0
        for para in self.doc.paragraphs:
            for i in range(len(para.runs)):
                val = para.runs[i].text
                for string, new_string in replace_dct.items():
                    if string in val:
                        val = val.replace(string, new_string)
                        para.runs[i].text = val
                        cnt += 1
        # ????????????
        for i in range(self.doc.sections.__len__()):
            header = self.doc.sections[i].header
            footer = self.doc.sections[i].footer
            for para in header.paragraphs:
                for j in range(len(para.runs)):
                    val = para.runs[j].text
                    for string, new_string in replace_dct.items():
                        if string in val:
                            val.replace(string, new_string)
                            para.runs[j].text = val
                            cnt += 1
            for para in footer.paragraphs:
                for j in range(len(para.runs)):
                    val = para.runs[j].text
                    for string, new_string in replace_dct.items():
                        if string in val:
                            val.replace(string, new_string)
                            para.runs[j].text = val
                            cnt += 1
        return cnt

    def replace(self, replace_dct: Dict, replace_cb_lst: List = None):
        """???????????????????????????"""
        cnt = 0
        cnt += self._replace_text(self.doc, replace_dct, replace_cb_lst)
        for i in range(self.doc.tables.__len__()):
            cnt += self._replace_table(self.doc.tables[i], replace_dct, replace_cb_lst)

        # for para in self.doc.paragraphs:
        #     val = para.text
        #     if replace_cb_lst:
        #         for func in replace_cb_lst:
        #             val, is_change = func(val)
        #             if is_change:
        #                 para.text = val
        #                 cnt += 1
        #     else:
        #         for string, new_string in replace_dct.items():
        #             res = re.search(string, val)
        #             if res:
        #                 val = re.sub(string, new_string, val)
        #                 cnt += 1
        #             para.text = val


        # ????????????
        for i in range(self.doc.sections.__len__()):
            header = self.doc.sections[i].header
            footer = self.doc.sections[i].footer
            cnt += self._replace_text(header, replace_dct, replace_cb_lst)
            cnt += self._replace_text(footer, replace_dct, replace_cb_lst)
            for i in range(header.tables.__len__()):
                cnt += self._replace_table(header.tables[i], replace_dct, replace_cb_lst)
            for i in range(footer.tables.__len__()):
                cnt += self._replace_table(footer.tables[i], replace_dct, replace_cb_lst)
            # for para in header.paragraphs:
            #     val = para.text
            #     print(val)
            #     for j in range(len(para.runs)):
            #         val2 = para.runs[j].text
            #         print(val2)
            #     if replace_cb_lst:
            #         for func in replace_cb_lst:
            #             val, is_change = func(val)
            #             if is_change:
            #                 para.text = val
            #                 cnt += 1
            #     else:
            #         for string, new_string in replace_dct.items():
            #             res = re.search(string, val)
            #             if res:
            #                 val = re.sub(string, new_string, val)
            #                 para.text = val
            #                 cnt += 1
            #
            # for para in footer.paragraphs:
            #     val = para.text
            #     if replace_cb_lst:
            #         for func in replace_cb_lst:
            #             val, is_change = func(val)
            #             if is_change:
            #                 para.text = val
            #                 cnt += 1
            #     else:
            #         for string, new_string in replace_dct.items():
            #             res = re.search(string, val)
            #             if res:
            #                 val = re.sub(string, new_string, val)
            #                 para.text = val
            #                 cnt += 1

        return cnt

    def save(self):
        """????????????"""
        self.doc.save(self.filename)

    def save_as(self, filename):
        """???????????????"""
        self.doc.save(filename)


class EditXlsx:
    def __init__(self, filename=None):
        self.wb = None
        self.filename = None
        if filename:
            self.filename = filename
            if os.path.exists(self.filename):
                self.wb = openpyxl.load_workbook(filename, data_only=True)
            else:
                raise NotImplemented
        else:
            raise NotImplemented

    def __del__(self):
        if self.wb:
            self.wb.close()

    def replace_str(self, replace_dct: Dict):
        """????????????"""
        cnt = 0
        for item in self.wb.sheetnames:
            cur = self.wb[item]
            for i, row in enumerate(cur.iter_rows()):
                for j, cell in enumerate(row):
                    val = str(cell.value)
                    for string, new_string in replace_dct.items():
                        if string in val:
                            val = val.replace(string, new_string)
                            cell.value = val
                            cnt += 1
        return cnt

    def replace(self, replace_dct: Dict, replace_cb_lst: List = None):
        """???????????????????????????"""
        cnt = 0
        for item in self.wb.sheetnames:
            cur = self.wb[item]
            for i, row in enumerate(cur.iter_rows()):
                for j, cell in enumerate(row):
                    val = str(cell.value)
                    if replace_cb_lst:
                        for func in replace_cb_lst:
                            val, is_change = func(val)
                            if is_change:
                                cell.value = val
                                cnt += 1
                    else:
                        for string, new_string in replace_dct.items():
                            res = re.search(string, val)
                            if res:
                                val = re.sub(string, new_string, val)
                                cell.value = val
                                cnt += 1
        return cnt

    def save(self):
        """????????????"""
        self.wb.save(self.filename)

    def save_as(self, filename):
        """???????????????"""
        self.wb.save(filename)


class EditPptx:
    def __init__(self, filename=None):
        self.prs = None
        self.filename = None
        if filename:
            self.filename = filename
            if os.path.exists(self.filename):
                self.prs = Presentation(filename)
            else:
                raise NotImplemented
        else:
            raise NotImplemented

    def __del__(self):
        if self.prs:
            pass

    @staticmethod
    def _replace_text(text_frame, replace_dct: Dict, replace_cb_lst: List = None):
        cnt = 0
        for paragraph in text_frame.paragraphs:
            val = paragraph.text
            if replace_cb_lst:
                for func in replace_cb_lst:
                    val, is_change = func(val)
                    if is_change:
                        paragraph.text = val
                        cnt += 1
            else:
                for string, new_string in replace_dct.items():
                    res = re.search(string, val)
                    if res:
                        val = re.sub(string, new_string, val)
                        paragraph.text = val
                        cnt += 1

        return cnt

    def replace_str(self, replace_dct: Dict):
        """????????????"""
        return self.replace(replace_dct)

    def replace(self, replace_dct, replace_cb_lst: List = None):
        """???????????????????????????"""
        cnt = 0
        for slide in self.prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:  # ??????Shape?????????????????????
                    text_frame = shape.text_frame
                    cnt += self._replace_text(text_frame, replace_dct, replace_cb_lst)  # ??????replace_text????????????????????????
                if shape.has_table:  # ??????Shape??????????????????
                    table = shape.table
                    for cell in table.iter_cells():  # ???????????????cell
                        text_frame = cell.text_frame
                        cnt += self._replace_text(text_frame, replace_dct, replace_cb_lst)  # ??????replace_text????????????????????????
        return cnt

    def save(self):
        """????????????"""
        self.prs.save(self.filename)

    def save_as(self, filename):
        """???????????????"""
        self.prs.save(filename)


class EditText:
    def __init__(self, filename=None):
        self.filename = None
        self.content = []
        self.my_encoding = 'utf-8'
        if filename:
            self.filename = filename
            try:
                # ????????????????????????
                import chardet
                with open(self.filename, 'rb') as f:
                    self.my_encoding = chardet.detect(f.read(200))['encoding']
                with open(filename, "r+", encoding=self.my_encoding) as f:
                    self.content = f.readlines()
            except Exception as e:
                print(e)
                self.content = []
        else:
            raise NotImplemented

    def __del__(self):
        pass

    def replace_str(self, replace_dct: Dict):
        """????????????"""
        cnt = 0
        if not self.content:
            return 0
        for i in range(self.content.__len__()):
            txt = self.content[i]
            for string, new_string in replace_dct.items():
                if string in txt:
                    txt = txt.replace(str, new_string)
                    self.content[i] = txt
                    cnt += 1
        return cnt

    def replace(self, replace_dct: Dict, replace_cb_lst: List = None):
        """???????????????????????????"""
        cnt = 0
        if not self.content:
            return 0
        for i in range(self.content.__len__()):
            val = self.content[i]
            if replace_cb_lst:
                for func in replace_cb_lst:
                    val, is_change = func(val)
                    if is_change:
                        self.content[i] = val
                        cnt += 1
            else:
                for string, new_string in replace_dct.items():
                    res = re.search(string, val)
                    if res:
                        val = re.sub(string, new_string, val)
                        self.content[i] = val
                        cnt += 1
        return cnt

    def save(self):
        """????????????"""
        if self.content:
            try:
                with open(self.filename, "w+", encoding=self.my_encoding) as f:
                    f.writelines(self.content)
            except Exception as e:
                print(e)

    def save_as(self, filename):
        """???????????????"""
        if self.content:
            try:
                with open(filename, "w+", encoding=self.my_encoding) as f:
                    f.writelines(self.content)
            except Exception as e:
                print(e)


def get_edit(file_name: str):
    sub_name = os.path.splitext(file_name)[-1].lower()
    m_map = {
        ".docx": EditDocx,
        ".doc": EditDocx,
        ".xlsx": EditXlsx,
        ".xls": EditXlsx,
        ".pptx": EditPptx,
        ".ppt": EditPptx,
        ".csv": EditText,
        ".txt": EditText
    }
    return m_map.get(sub_name, EditText)


_y_map = {"0": "???", "1": "???", "2": "???", "3": "???", "4": "???",
         "5": "???", "6": "???", "7": "???", "8": "???", "9": "???"}
_y_map_2 = {"0": "???", "1": "???", "2": "???", "3": "???", "4": "???",
         "5": "???", "6": "???", "7": "???", "8": "???", "9": "???"}
_m_map = {"1": "???", "2": "???", "3": "???", "4": "???",
         "5": "???", "6": "???", "7": "???", "8": "???", "9": "???",
         "01": "???", "02": "???", "03": "???", "04": "???",
         "05": "???", "06": "???", "07": "???", "08": "???", "09": "???",
         "10": "???", "11": "??????", "12": "??????"}

_m_simple_month_days_count = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]


def replace_date(input_str: str, ori_date: List, dst_date: List,
                 fmt_index_used: Tuple[int] = (0, 1), for_excel_sp: bool = False):
    """
    TODO ??????excel????????????2021???1?????? ?????????openexls?????????value???1900???1???1?????????????????????int,
    ????????????int??????????????????????????????????????????????????????int?????????
    """
    # print("func ", ori_date, dst_date, fmt_index_used, input_str, "aaa")
    assert ori_date.__len__() == dst_date.__len__()
    ori_date_str = [str(ori_date[0]), str(ori_date[1])]
    dst_date_str = [str(dst_date[0]), str(dst_date[1])]
    y_str = "".join(_y_map[item] for item in str(ori_date[0]))
    y_str_2 = "".join(_y_map_2[item] for item in str(ori_date[0]))
    dst_m_str = _m_map[str(dst_date[1])]
    ret_str = ""

    # ????????????,excel???????????????
    if for_excel_sp and input_str.isdigit():
        src_stamp_day = int(input_str)
        cmp_stamp_day = (datetime.date(ori_date[0], ori_date[1], 1) - datetime.date(1900, 1, 1)).days + 2
        time_dlt = src_stamp_day - cmp_stamp_day
        # print("date is int ", cmp_stamp_day, int(input_str), ori_date[0], ori_date[1])
        if 0 <= time_dlt <= _m_simple_month_days_count[ori_date[1] - 1]:
            input_day = datetime.date(ori_date[0], ori_date[1], 1) + datetime.timedelta(days=time_dlt)
            ret_str = "{0}???{1}???{2}???".format(dst_date[0], dst_date[1], input_day.day)
            return ret_str, True

    if ori_date.__len__() == 2:
        mc_fmts = [
            r"(" + str(ori_date[0]) + r")" +
            r"(\s?[\.\-\s\\/???]{1,1}\s?)(" + str(ori_date[1]) + r"|" + "{0:02d}".format(ori_date[1]) +
            r")(\s?[\s\.\-\\/???]{0,1}[0-9]{0,2})",
            r"(" + y_str + r"|" + y_str_2 + r")" +
            r"(\s?[\s\.???]{1,1}\s?)(" + _m_map[str(ori_date[1])] + r")(\s?[\s\.???$\r\n]{1,1})"]
        mc_fmts = [mc_fmts[i] for i in fmt_index_used]
        found = False
        ret_str = input_str
        # if input_str and input_str != "None":
        #     print(mc_fmts, input_str)
        for fmt in mc_fmts:
            if re.search(fmt, ret_str):
                # print("FFFF ", re.split(fmt, ret_str))
                found = True
                words = re.split(fmt, ret_str)
                i = 0
                # print(words)
                while i < (words.__len__() - 3):
                    if not words[i]:
                        i += 1
                        continue
                    per_word = "".join([words[i + per] for per in range(4)])
                    if re.match(fmt, per_word):
                        # print("aaa", per_word, ori_date_str, i, words)
                        if words[i + 0] == ori_date_str[0]:  # 2021
                            if words[i + 2].__len__() == 2:  # 2021.02
                                words[i + 0] = dst_date_str[0]
                                words[i + 2] = "{0:02d}".format(dst_date[1])
                            else:  # 2021.02
                                words[i + 0] = dst_date_str[0]
                                words[i + 2] = "{0:d}".format(dst_date[1])
                        else:  # ????????????
                            if re.search(r'???', words[i + 0]):  # ????????????
                                words[i + 0] = "".join(_y_map_2[item] for item in dst_date_str[0])
                                words[i + 2] = dst_m_str
                            else:
                                words[i + 0] = "".join(_y_map[item] for item in dst_date_str[0])
                                words[i + 2] = dst_m_str
                        i += 4
                    else:
                        i += 1
                # print(words)
                ret_str = "".join(words)
                # print("Done", ret_str)
    else:
        raise NotImplementedError  # TODO
    return ret_str, found


def replace_dir(dir_path: str, replace_dct: Dict, filter_lst: List = None, logger=None, replace_cb_lst: List = None):
    if os.path.isdir(dir_path):
        tt_cnt = 0
        for item in os.listdir(dir_path):
            abs_path = os.path.join(dir_path, item)
            tt_cnt += replace_dir(abs_path, replace_dct, filter_lst, logger=logger, replace_cb_lst=replace_cb_lst)

        return tt_cnt
    elif os.path.isfile(dir_path):
        filename = os.path.split(dir_path)[-1]
        found = False
        if filter_lst:
            for item in filter_lst:
                if re.search(item, filename):
                    found = True
                    break
        if found:
            if re.search(r"~.*\..*", filename):  # ??????~???????????????
                found = False
            else:
                found = True
        if found:
            logger("??????????????????: ???{0}???".format(filename))
            cnt = 0
            try:
                inst = get_edit(file_name=filename)(dir_path)
                cnt = inst.replace(replace_dct, replace_cb_lst=replace_cb_lst)
                inst.save()
                logger("????????????????????????{0}???, ?????????{1}???".format(cnt, filename))
            except Exception as e1:
                logger("e: {0} file: {1}".format(e1, os.path.split(dir_path)[-1]))
                raise e1
            return cnt
    return 0


def office2officeX(path, logger=print):
    path = path.replace("/", os.sep).replace("\\", os.sep)
    for subPath in os.listdir(path):
        subPath = os.path.join(path, subPath)
        if os.path.isdir(subPath):
            office2officeX(subPath, logger=logger)
        elif re.search(r"~.*\..*", subPath):  # ??????~???????????????
            continue
        elif subPath.endswith('.ppt'):
            if os.path.exists("{}x".format(subPath)):
                continue
            pythoncom.CoInitialize()
            logger("??????????????????{0} to pptx...".format(os.path.split(subPath)[-1]))
            powerpoint = win32com.client.Dispatch('PowerPoint.Application')
            win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
            powerpoint.Visible = 1
            ppt = powerpoint.Presentations.Open(subPath)
            ppt.SaveAs("{}x".format(subPath))
            ppt.Close()
            try:
                powerpoint.Quit()  # ????????????
                os.remove(subPath)
            except Exception as e:
                raise e
        elif subPath.endswith('.doc'):
            if os.path.exists("{}x".format(subPath)):
                continue
            pythoncom.CoInitialize()
            logger("??????????????????{0} to docx...".format(os.path.split(subPath)[-1]))
            time.sleep(1)  # ?????????sleep????????????word????????? TODO
            word = wc.Dispatch("Word.Application")  # ??????word????????????
            doc = word.Documents.Open(subPath)  # ??????word??????
            doc.SaveAs("{}x".format(subPath), FileFormat=12)  # ??????????????????".docx"????????????????????????12???docx??????
            doc.Close()  # ????????????word??????
            try:
                word.Quit()
                os.remove(subPath)
            except Exception as e:
                raise e
        elif subPath.endswith('.xls'):
            if os.path.exists("{}x".format(subPath)):
                continue
            pythoncom.CoInitialize()
            logger("??????????????????{0} to xlsx...".format(os.path.split(subPath)[-1]))
            excel = wc.Dispatch("Excel.Application")  # ??????word????????????
            wb = excel.Workbooks.Open(subPath)
            wb.SaveAs("{}x".format(subPath), FileFormat=51)  # ??????????????????".docx"????????????????????????12???docx??????
            wb.Close()
            try:
                excel.Quit()
                os.remove(subPath)
            except Exception as e:
                raise e


# if __name__ == "__main__":
    # office2officeX(r'D:\share_dir\product_env\01.SVN\01.local_git\DocTreat\YYY')
    # doc2docx(r'D:\share_dir\product_env\01.SVN\01.local_git\DocTreat\XXX\aa.doc')
    # e = EditXlsx(r'D:\share_dir\product_env\01.SVN\01.local_git\DocTreat\YYY\AAA.xlsx')
    # e = EditDocx(r'D:\share_dir\product_env\01.SVN\01.local_git\DocTreat\YYY\AAA.docx')
    # e = EditPptx(r'D:\share_dir\product_env\01.SVN\01.local_git\DocTreat\YYY\AAA.pptx')
    # e = EditText(r'D:\share_dir\product_env\01.SVN\01.local_git\DocTreat\YYY\AAA.csv')
    # # e.replace_str("HAHA", "sss")
    # e.replace({r'([0-9].*)UUU[A-Za-z].*': 'HELLO', 'D': 'HELLO'})
    # e.save()

    # powerpoint = win32com.client.Dispatch('PowerPoint.Application')
    # powerpoint.Visible = 1
    # subPath = "D:/share_dir/product_env/01.SVN/01.local_git/DocTreat/YYY/AAA - ?????? (2).ppt"
    # subPath = subPath.replace("/",os.sep)
    # print(subPath)
    # ppt = powerpoint.Presentations.Open(subPath)

# msg, is_change = replace_date("\n2021???11???11???\n", (2021, 11), (2021, 8))
# print(msg, is_change)

