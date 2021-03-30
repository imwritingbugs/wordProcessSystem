#!/usr/bin/env python
# coding=utf-8
from zipfile import ZipFile
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
import string
import itertools
import xml.etree.ElementTree as ET


def read_file(filename):
    artical = ""
    document = ZipFile(filename)
    xml = document.read("word/document.xml")
    wordObj = BeautifulSoup(xml.decode("utf-8"), features="html.parser")
    texts = wordObj.findAll("w:t")
    # print(texts)
    for text in texts:
        artical = artical + text.string
    return artical


def get_comment(filename):
    comment = []
    document = ZipFile(filename)
    xml = document.read("word/comments.xml")
    wordObj = BeautifulSoup(xml.decode("utf-8"), features="html.parser")

    texts = wordObj.findAll("w:t")
    for text in texts:
        comment.append(text.string)
    return comment


def neighborhood(iterable):
    iterator = iter(iterable)
    prev_item = None
    current_item = next(iterator)  # throws StopIteration if empty.
    for next_item in iterator:
        yield (prev_item, current_item, next_item)
        prev_item = current_item
        current_item = next_item
    yield (prev_item, current_item, None)


def str_count(artical):
    count_en = count_dg = count_sp = count_zh = count_pu = count_dg_num = count_en_num = 0  # 统一将0赋值给这5个变量
    s_len = len(artical)
    for prev, item, next in neighborhood(artical):
        # 统计英文
        if item in string.ascii_letters:
            count_en += 1
            # 统计数字的个数
            if prev in string.ascii_letters:
                pass
            else:
                count_en_num += 1
        # 统计数字的位数
        elif item.isdigit():
            count_dg += 1
            # 统计数字的个数
            if prev.isdigit():
                pass
            else:
                count_dg_num += 1
        # 统计空格
        elif item.isspace():
            count_sp += 1
        # 统计中文
        elif item.isalpha():
            count_zh += 1
        # 统计特殊字符
        else:
            count_pu += 1
    total_chars = count_zh + count_en + count_sp + count_dg + count_pu
    if total_chars == s_len:
        return "字数: {7}\n字符数：{6}(不计空格)\n字符数：{0}(计空格)\n中文字符：{1}\n英文字符：{2}\n空格：{3}\n数字个数：{4}\n标点符号：{5}\n".format(
            s_len,
            count_zh,
            count_en,
            count_sp,
            count_dg,
            count_pu,
            s_len - count_sp,
            count_zh + count_en_num + count_pu + count_dg_num,
        )


def find_color(filename):
    redText = []
    greenBgText = []
    yellowBgText = []
    jump = []
    redStr = ""
    doc = Document(filename)
    # TODO:完成统一段落中的句子连接
    for p in doc.paragraphs:
        # print(p.text)
        for r in p.runs:
            print(r.text)
            if r.font.color.rgb == RGBColor(255, 0, 0):
                redStr += r.text
        redText.append(redStr)
        redStr = ""
        # if r.font.highlight_color == WD_COLOR_INDEX.YELLOW:
        #     jump.append(r.text)
        #     # if r.font.highlight_color == None:
        #     #     print("dd")
        #     try:
        #         # print(r.text, r.font, r.font.color.rgb, r.font.highlight_color)
        #     except:
        #         if r.font.color.rgb == RGBColor(255, 0, 0):
        #             redText.append(r.text)
        #     else:
        #         if r.font.color.rgb == RGBColor(255, 0, 0):
        #             redText.append(r.text)
        #         if r.font.highlight_color == WD_COLOR_INDEX.BRIGHT_GREEN:
        #             greenBgText.append(r.text)
        #         elif r.font.highlight_color == WD_COLOR_INDEX.YELLOW:
        #             yellowBgText.append(r.text)
    print(redStr)
    return redText, greenBgText, yellowBgText, jump


filename = "./sample.docx"
artical = read_file(filename)
comment = get_comment(filename)
cnt_result = str_count(artical)
# print(artical)
print(comment)
# print(cnt_result)
# red, green, yellow, jump = find_color(filename)
# print(red)
# print(jump)
# print(yellow)
# print("Error: 没有找到文件或读取文件失败")

