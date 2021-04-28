# coding=utf-8
import os
import re
import docx
import json
import jieba
import pathlib
from win32com import client as wc


class PaperTextChecker(object):

    def __init__(self,
                 file_path,
                 par_chinese_threshold=0.2,
                 white_tables=None,
                 from_json_file=True,
                 config_file_path='.\\config.json'):
        """
        :param file_path: 字符串类型或pathlib.Path类, 表示要读取的文件路径
        :param par_chinese_threshold: 浮点型, 其值介于[0,1]之间, 默认0.2
                                    表示在一个段落中中文占比比例如果小于这个阈值,
                                    则筛掉这个段落, 主要是为了过滤代码段
        :param white_tables: 列表或元组类型, 默认为None
                            白名单, 在一句话中可以单独出现的词组, 此时视为正常
        :param from_json_file: 是否从配置文件读取信息, 默认为True
                                如果为False, 则config_file_path参数不起作用
        :param config_file_path: 字符串类型或pathlib.Path类, 配置文件所在路径
        """
        if not isinstance(file_path, (str, pathlib.Path)):
            raise ValueError('参数file_path类型错误')
        if not isinstance(from_json_file, bool):
            raise ValueError('参数from_json_file类型错误')
        if not from_json_file:
            config_file_path = None
        if config_file_path is not None:
            if not isinstance(config_file_path, (str, pathlib.Path)):
                raise ValueError('参数config_file_path类型错误')
            elif isinstance(config_file_path, pathlib.Path):
                config_file_path = str(config_file_path)
        if isinstance(file_path, pathlib.Path):
            file_path = str(file_path)
        file_path = os.path.abspath(file_path)
        config_file_path = os.path.abspath(config_file_path)
        now_folder = os.path.abspath(os.path.dirname('.\\report\\'))
        base_name = os.path.basename(file_path)
        md_file_name = ''.join(base_name.split('.')[:-1]) + "_论文检测报告.md"
        file_path = PaperTextChecker.doc2doc_x(file_path)
        self.file = docx.Document(file_path)
        self.valid_paragraph = list()
        if not config_file_path:
            self.config = dict()
        else:
            self.config = json.load(open(config_file_path, 'r', encoding='utf-8'))
        if 'par_chinese_threshold' in self.config.keys():
            self.par_chinese_threshold = self.config['par_chinese_threshold']
        else:
            self.par_chinese_threshold = par_chinese_threshold
        if not isinstance(par_chinese_threshold, (int, float)):
            raise ValueError('参数par_chinese_threshold只能是介于0-1之间的数')
        if not 0 <= par_chinese_threshold <= 1:
            raise ValueError('参数par_chinese_threshold的值只能介于0-1之间')
        self.is_main_text = False
        self.has_appear_key_word = False
        if white_tables is None:
            self.white_tables = list()
        else:
            self.white_tables = list(white_tables)
        if 'white_tables' in self.config.keys():
            self.white_tables = list(self.config['white_tables'])
        if white_tables is not None and not isinstance(white_tables, (list, tuple)):
            raise ValueError('参数white_tables类型错误')
        self.error_paragraph = list()
        self.check_report_file_path = os.path.join(now_folder, md_file_name)
        self.check_report_file = open(self.check_report_file_path, 'w', encoding='utf-8')

    def is_chinese_paragraph(self, paragraph, par_chinese_threshold=None):
        """
        :param paragraph: 要筛选的段落
        :param par_chinese_threshold: 最低中文比例筛选阈值, 等同于类的同名参数
        :return: bool类型, 是否是中文段落
        """
        if par_chinese_threshold is None:
            par_chinese_threshold = self.par_chinese_threshold
        chinese_len = len(list(filter(lambda ch: '\u4e00' <= ch <= '\u9fff', paragraph)))
        if chinese_len / len(paragraph) < par_chinese_threshold:
            return False
        return True

    def has_end_punctuation(self, paragraph):
        """
        主要是为了筛选掉标题等非段落的paragraph, 如果段落中出现'。：！？；......'中
        之一才被识别为正常段落
        :param paragraph: 要筛选的段落
        :return: bool类型, 是否是正常段落
        """
        if 'end_punctuations' in self.config.keys():
            end_punctuations = self.config['end_punctuations']
        else:
            end_punctuations = ['。', '：', '！', '？', '；', '......']
        for punctuation in end_punctuations:
            if paragraph.find(punctuation) != -1:
                return True
        return False

    def is_valid_paragraph(self, paragraph):
        """
        多条件判断是否是合法段落
        :param paragraph: 要筛选的段落
        :return: bool类型, 是否是合法段落
        """
        paragraph_text = paragraph.text.strip('\n\t ')
        if paragraph_text.startswith('关键词') or paragraph_text.startswith('关键字') \
                and self.has_appear_key_word is False:
            self.is_main_text = True
            self.has_appear_key_word = True
            return False
        if paragraph_text == '参考文献':
            self.is_main_text = False
            return False
        if self.is_main_text is False:
            return False
        if paragraph.paragraph_format.first_line_indent is None:
            return False
        if not isinstance(paragraph.paragraph_format.first_line_indent, int):
            return False
        paragraph_text = paragraph.text.replace(' ', '').replace('\n', '').replace('\t', '')
        if paragraph_text == "":
            return False
        if not self.is_chinese_paragraph(paragraph_text):
            return False
        if not self.has_end_punctuation(paragraph_text):
            return False
        return True

    def get_all_valid_paragraph(self):
        for paragraph in self.file.paragraphs:
            if self.is_valid_paragraph(paragraph):
                self.valid_paragraph.append(paragraph.text.strip('\n\t .'))

    def pre_process_paragraph(self, paragraph):
        """
        预处理段落, 过滤掉'、“”（）—《》~·‘’，：'字符以及非字母、数字的ascii字符
        :param paragraph:要进行预处理的段落
        :return:去除过滤字符后的索引列表和字符串
        """
        if 'ignore_chinese_chars' in self.config.keys():
            ignore_chinese_chars = self.config['ignore_chinese_chars']
        else:
            ignore_chinese_chars = "、“”（）—《》~·‘’，："

        def filter_valid_word(index):
            word = paragraph[index]
            if word in ignore_chinese_chars:
                return False
            if ord(word) < 128 and not word.isalnum():
                    return False
            return True
        if 'split_punctuations' in self.config.keys():
            split_punctuations = self.config['split_punctuations']
        else:
            split_punctuations = ['！', '。', '？', '；']
        filter_list, filter_paragraph = list(), list()
        for idx in range(len(paragraph)):
            if filter_valid_word(idx):
                filter_list.append(idx)
                if paragraph[idx] in split_punctuations:
                    filter_paragraph.append(' ')
                else:
                    filter_paragraph.append(paragraph[idx])
        return filter_list, ''.join(filter_paragraph)

    def check_all_paragraph(self):
        """
        检查所有的段落
        """
        self.get_all_valid_paragraph()
        for paragraph in self.valid_paragraph:
            self.check_single_paragraph(paragraph)
        self.write_report()

    @staticmethod
    def add_color(comment, color='red'):
        """
        为文本添加颜色打印结果
        :param comment: 要添加颜色的字符串
        :param color: 颜色值, 默认为红色
        :return: 添加颜色后的字符串
        """
        return f"<font color={color}>{comment}</font>"

    def write_report(self):
        """
        将检测信息写入md文件中
        :return: None
        """
        if not self.error_paragraph:
            self.check_report_file.write('未检测出错误信息!完美!')
        else:
            for paragraph in self.error_paragraph:
                self.check_report_file.write(f'\n\n{PaperTextChecker.add_color("段落出错:")}\n$\\qquad${paragraph}\n')
        self.check_report_file.close()

    def check_single_paragraph(self, paragraph: str):
        """
        检测单个段落
        :param paragraph: 要检测的段落
        :return: None
        """
        filter_index_list, filter_paragraph = self.pre_process_paragraph(paragraph)
        split_list = filter_paragraph.split()
        error_segment = list()
        start = 0
        for segment in split_list:
            segment_split_list = list(jieba.cut(segment))
            if len(segment_split_list) == 1 and segment_split_list[0] not in self.white_tables \
                    and re.search(r'[\u4e00-\u9fa5]', segment_split_list[0]):
                error_segment.append((filter_index_list[start], filter_index_list[start + len(segment) - 1]))
            start += 1 + len(segment)
        if error_segment:
            error_paragraph, start = "", 0
            for i in range(len(error_segment)):
                error_paragraph += paragraph[start: error_segment[i][0]] + \
                                   PaperTextChecker.add_color(paragraph[error_segment[i][0]: error_segment[i][1] + 1])
                start = error_segment[i][1] + 1
            error_paragraph += paragraph[start:]
            self.error_paragraph.append(error_paragraph)

    @staticmethod
    def doc2doc_x(file_path: str):
        """
        将后缀为doc的文件转换为后缀为docx的文件
        :param file_path: 文件名
        :return: 转换后的文件名
        """
        if file_path.endswith('.docx'):
            return file_path
        if file_path.endswith('.doc'):
            print('正在将.doc文件转换为.docx文件')
            word = wc.Dispatch('Word.Application')
            doc = word.Documents.Open(file_path)
            doc.SaveAs(file_path + 'x', 12, False, "", True, "", False, False, False, False)
            doc.Close()
            word.Quit()
            print('转换完成')
            return file_path + 'x'
        postfix = file_path.split('.')[-1]
        raise ValueError(f"暂不支持 {postfix} 格式的文件, 请选择 doc, docx 格式")


if __name__ == '__main__':
    try:
        checker = PaperTextChecker('C:\\Users\\25783\\Desktop\\需要检查的论文\\硕士学位论文正文_2.docx')
        checker.check_all_paragraph()
    except Exception as error:
        print(str(error))
