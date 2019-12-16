# -*-coding:utf-8-*-

import re
import pythoncom
import win32com.client as win32


class WordOperator:
    def __init__(self, file_path=''):
        self.file_path = file_path

    def __enter__(self):
        pythoncom.CoInitialize()
        self.doc, self.word = self._init_doc()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.doc.Close(True)
        self.word.Application.Quit()
        pythoncom.CoUninitialize()

    def _init_doc(self):
        try:
            word = win32.gencache.EnsureDispatch("Word.Application")
        except:
            word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(self.file_path)
        return doc, word

    def _convert_regex_pos_to_doc_pos(self, pos):
        '''
        由于某些原因， 正则得到的字符位置和 word 文档中实际的字符位置有偏差，
        这个方法用于将正则得到的位置转化到 word 文档中的对应位置
        '''
        return self.doc.Characters.Item(pos).Start + 1

    def get_full_text(self):
        return self.doc.Content.Text.replace('\x07', '')\
                                    .replace('\r', '\n')\
                                    .replace('\xa0', ' ')\
                                    .replace('\x0c', ' ')\
                                    .replace('\t', ' ')

    def get_partial_text(self, start_index, end_index):
        return self.doc.Range(start_index, end_index).Text.replace('\x07', '')\
                                                        .replace('\r', '\n')\
                                                        .replace('\xa0', ' ')\
                                                        .replace('\x0c', ' ')\
                                                        .replace('\t', ' ')

    def search_text(self, match_text):
        myRange = self.doc.Content
        myRange.Find.Execute(FindText=match_text, Forward=True)
        if myRange.Find.Found:
            return myRange.Start, myRange.End

    def insert_by_match(self, match_text, insert_text, before=False, underline=False, offset=0):
        # myRange = self.doc.Content
        # myRange.Find.Execute(FindText=match_text, Forward=True)
        # if myRange.Find.Found:
        #     if before:
        #         pos = myRange.Start - offset
        #     else:
        #         pos = myRange.End + offset
        #     self.doc.Range(pos, pos).InsertAfter(insert_text)
        #     self.doc.Range(pos, pos+len(insert_text)).Font.Underline = underline
        #     return True
        start = 0
        index = self.get_full_text().find(match_text, start)
        while index >= 0:
            pos = self._convert_regex_pos_to_doc_pos(index)
            if before:
                pos = pos - offset
            else:
                pos = pos + len(match_text) + offset
            self.doc.Range(pos, pos).InsertAfter(insert_text)
            self.doc.Range(pos, pos + len(insert_text)).Font.Underline = underline

            start = index + len(match_text) + len(insert_text) + 1
            index = self.get_full_text().find(match_text, start)



    def insert_by_regex(self, pattern, insert_text, before=False, underline=False, offset=0, flag=0):
        regex = re.compile(pattern, flags=flag)
        start = 0
        regex_result = regex.search(self.get_full_text(), pos=start)
        while regex_result:
            if before:
                pos = self._convert_regex_pos_to_doc_pos(regex_result.start()) - offset
                # pos = regex_result.start() - offset
            else:
                pos = self._convert_regex_pos_to_doc_pos(regex_result.end()) + offset
                # pos = regex_result.end() + offset
            self.doc.Range(pos, pos).InsertAfter(insert_text)
            self.doc.Range(pos, pos+len(insert_text)).Font.Underline = underline
            start = regex_result.end() + len(insert_text) + 1
            regex_result = regex.search(self.get_full_text(), pos=start)


    def replace_by_match(self, match_text, replace_text, underline=False):
        # myRange = self.doc.Content
        # myRange.Find.Execute(FindText=match_text, Forward=True)
        # if myRange.Find.Found:
        #     self.doc.Range(myRange.Start, myRange.End).Text = replace_text
        #     self.doc.Range(myRange.Start, myRange.Start+len(replace_text)).Font.Underline = underline
        #     return True
        start = 0
        index = self.get_full_text().find(match_text, start)
        while index >= 0:
            pos = self._convert_regex_pos_to_doc_pos(index)
            self.doc.Range(pos, pos+len(match_text)).Text = replace_text
            self.doc.Range(pos, pos+len(replace_text)).Font.Underline = underline
            start = index + len(replace_text) + 1
            index = self.get_full_text().find(match_text, start)


    def replace_by_regex(self, pattern, replace_text, underline=False, flag=0):
        '''
        替换的部分使用括号
        '''
        regex = re.compile(pattern, flags=flag)
        start = 0
        regex_result = regex.search(self.get_full_text(), pos=start)
        while regex_result:
            # start, end = regex_result.start(1), regex_result.end(1)
            start = self._convert_regex_pos_to_doc_pos(regex_result.start(1))
            end = self._convert_regex_pos_to_doc_pos(regex_result.end(1))
            self.doc.Range(start, end).Text = replace_text
            self.doc.Range(start, start+len(replace_text)).Font.Underline = underline
            start = regex_result.start(0) + len(replace_text) + 1
            regex_result = regex.search(self.get_full_text(), pos=start)

    def get_table_cell_content(self, table_index, row, column):
        tbl = self.doc.Tables.Item(table_index)
        return tbl.Cell(row, column).Range.Text.replace('\r', ' ').replace('\x07', ' ')

    def insert_into_table_cell(self, table_index, row, column, insert_text, replace=True, underline=False):
        tbl = self.doc.Tables.Item(table_index)
        cell_range = tbl.Cell(row, column).Range
        if replace:
            cell_range.Text = insert_text
        else:
            cell_range.InsertAfter(insert_text)
        cell_range.Font.Underline = underline
        return True

    def make_choice(self, pattern, options, check_item):
        '''
        pattern: regex string
        options: dict {item : offset}
        check_item: str
        '''
        # myRange = self.doc.Content
        # myRange.Find.Execute(FindText=pattern, Forward=True)
        # if myRange.Find.Found:
        #     offset = options.get(check_item)
        #     if offset != None:
        #         pos = myRange.End + offset
        #         self.doc.Range(pos, pos+1).Text = u''
        #         self.doc.Range(pos, pos+1).Font.Name = 'wingdings 2'
        #         return True
        regex = re.compile(pattern)
        start = 0
        regex_result = regex.search(self.get_full_text(), pos=start)
        while regex_result:
            pos = self._convert_regex_pos_to_doc_pos(regex_result.end(0))
            offset = options.get(check_item)
            if offset != None:
                pos = pos + offset
                self.doc.Range(pos, pos+1).Text = u''
                self.doc.Range(pos, pos+1).Font.Name = 'wingdings 2'
                start = regex_result.end(0) + 1
                regex_result = regex.search(self.get_full_text(), pos=start)
            else:
                break

if __name__ == '__main__':
    with WordOperator('C:\\Users\\xuhuan\\Desktop\\fill_doc\\123.docx') as wd:
        # wd.insert_by_regex(r'平方米，其中地上(?:.*?)平方米，地下', 'hello, world', underline=True)
        # start, end = wd.search_text(u'增值税普通发票')
        # print(wd.get_partial_text(1316, 1331))
        # wd.make_choice(u'增值税普通发票', options={'16%': 2, '10%': 7}, check_item='10%')
        # wd.replace_by_regex(r'发包方(以下简称甲方)：', u'深圳', True)
        # print(wd.doc.Content.End)
        # print(wd.doc.Characters.Count)
        print(len(wd.get_full_text()))
        regex = re.compile(r'若在(.*?)项目')
        s = regex.search(wd.get_full_text())
        print(s)
