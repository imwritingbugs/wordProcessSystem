#!/usr/bin/env python
# coding=utf-8
import os
import re
import string
import xml.etree.ElementTree as ET
import zipfile
from bs4 import BeautifulSoup

info = []
err = []
namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def unzip_file(filename):
    global err
    global info
    # print("unzip", filename)
    r = zipfile.is_zipfile(filename)
    dir_path, file_path = os.path.split(filename)
    exist_comments = False
    if r:
        fz = zipfile.ZipFile(filename, "r")
        for file in fz.namelist():
            if file == "word/document.xml":
                fz.extract(file, dir_path + "./doctmp")
            if file == "word/comments.xml":
                exist_comments = True
                fz.extract(file, dir_path + "./doctmp")
    else:
        err.append("解压错误：该文件无法解压，请使用.docx后缀的文件")
        return "error", "error"
    document_path = dir_path + "/doctmp/word/document.xml"
    if exist_comments:
        comments_path = dir_path + "/doctmp/word/comments.xml"
    else:
        comments_path = "notexist"
    return document_path, comments_path


# 解压文件，返回纯文字部分
def read_file(filename):
    artical = ""
    document = zipfile.ZipFile(filename)
    xml = document.read("word/document.xml")
    wordObj = BeautifulSoup(xml.decode("utf-8"), features="html.parser")
    texts = wordObj.findAll("w:t")
    # print(texts)
    for text in texts:
        artical = artical + text.string
    return artical


# 从xml中解析得到评论的内容列表
def get_comment(filename):
    tree = ET.parse(filename)
    root = tree.getroot()
    content_list = []
    for comment in root:
        content = ""
        for para in comment:
            for run in para.findall(namespace + "r"):
                for text in run.findall(namespace + "t"):
                    content += text.text
        content_list.append(content)
    # print(content_list)
    return content_list


# 获取三种颜色的字体，已完成
def get_color(xml_file):
    global err
    global info
    tree = ET.parse(xml_file)
    root = tree.getroot()
    document = root[0]
    red_list = []
    yellow_list = []
    green_list = []
    red_content = ""
    yellow_content = ""
    green_content = ""
    highlight_err_content = ""
    err_content_2 = ""
    id = 1
    for para in document.findall(namespace + "p"):
        red_content = ""
        yellow_content = ""
        green_content = ""
        highlight_err_content = ""
        font_err_content = ""
        for r in para.findall(namespace + "r"):
            # print(r.tag, r.attrib, r.text)
            rpr = r[0]
            color = rpr.findall(namespace + "color")
            highlight = rpr.findall(namespace + "highlight")
            # 如果不存在颜色标签，则设置为auto
            if len(color) != 0:
                color = color[0].get(namespace + "val")
            else:
                color = "auto"

            # 同理处理高亮
            if len(highlight) != 0:
                # 存在高亮，判断为黄或者绿色
                highlight = highlight[0].get(namespace + "val")
                # print(highlight)
                if highlight == "yellow":
                    for t in r.findall(namespace + "t"):
                        # print(t.text)
                        yellow_list.append(t.text)
                elif highlight == "green":
                    for t in r.findall(namespace + "t"):
                        green_content += t.text
                elif highlight == "none":
                    pass
                else:
                    # 标错颜色了
                    for t in r.findall(namespace + "t"):
                        highlight_err_content += t.text

            if color == "FF0000" or "ff0000":
                # 是红色，则提取所有字
                for t in r.findall(namespace + "t"):
                    # print(t.text)
                    red_content += t.text
            elif color == "auto" or color == "000000":
                # 是黑色，不管
                pass
            else:
                # 是其他颜色，提醒
                # print(color)
                for t in r.findall(namespace + "t"):
                    # print(t.text)
                    font_err_content += t.text

        if len(highlight_err_content) != 0:
            err.append("高亮颜色错误: " + highlight_err_content + '"高亮颜色不符合要求！')
        if len(font_err_content) != 0:
            err.append("字体颜色错误: " + font_err_content + '"字体颜色不符合要求！')
        red_list.append(red_content)
        green_list.append(green_content)
        id += 1
    return red_list, yellow_list, green_list


def parse_red(red_list):
    global err
    global info
    no_error = True
    id = 1
    for line in red_list:
        # print(line)
        if line.endswith("。") or line.endswith("！") or line.endswith("？"):
            # 查找句子结束标志
            cnt = line.count("。") + line.count("！") + line.count("？")
            if cnt > 3:
                # 过于冗长
                no_error = False
                # print(line)
                err.append("标注错误: 第" + str(id) + "段重点部分过长，请勿超过三句话。定位：" + line[0:3])
        else:
            # 会议纪要最后未以。结尾
            cnt = line.count("。") + line.count("！") + line.count("？") + 1
            # print(cnt)
            if cnt > 3:
                # 过于冗长
                no_error = False
                err.append("标注错误: 第" + str(id) + "段重点部分过长，请勿超过三句话。定位：" + line[0:3])
        id += 1


# 获得评论位置里的内容
def get_comment_location(xml_file, green_list):
    global err
    global info
    f = open(xml_file, encoding="utf-8")
    xml_str = ""
    while True:
        line = f.readline()
        xml_str += line
        if not line:
            break
    # 查找所有范围标注
    range_location_str = ''
    range_pattern = "(commentRangeStart)([\s\S]*?)(commentRangeEnd)"
    range_res = re.findall(range_pattern, xml_str)
    comment_range_start_num = len(range_res)
    # 以下为之前的，查找到范围标注就报错的代码
    for comment in range_res:
        comment = list(comment)
        # print(comment)
        # print(type(comment))
        for tmp in comment:
            comment_pattern = "(<w:t>)([\s\S]*?)(</w:t>)"
            res2 = re.findall(comment_pattern, tmp)
            if res2:
                for comment2 in res2:
                    comment2 = list(comment2)[1]
                    range_location_str += comment2
    # 查看有没有非范围标注的标注
    single_pattern = "(commentReference)"
    single_res = re.findall(single_pattern, xml_str)
    comment_reference_num = len(single_res)

    # print(comment_range_start_num, comment_reference_num)
    if comment_range_start_num != comment_reference_num:
        if len(range_location_str) != 0:
            err.append("批注错误：存在非范围批注，请检查！范围标注有：" + range_location_str)
        else:
            err.append("批注错误：存在非范围批注，请检查！")
    # print("location", location_str)
    # print(location_str)
    # # 以下为判断范围标注的序号
    # location_str_list = location_str.split("【")
    # # print(location_str_list)
    # new_location_list = []
    # for lo in location_str_list:
    #     if len(lo) != 0:
    #         lo = '【' + lo
    #         new_location_list.append(lo)
    # print(new_location_list)

    # info.append("标注所在位置为：" + location_str)
    green = list(set(green_list))
    # print("green", green)
    f.close()
    # return new_location_list
    # print("green", green)
    # for green in green_list:

    # 判断第二次出现时是否有批注

    # if matchObj:
    #     print("matchObj.group() : ", matchObj.group())
    # else:
    #     print("No match!!")
    # # tree = ET.parse(xml_file)
    # root = tree.getroot()
    # body = root[0]
    # print(body.tag)
    # for p in body:
    #     for comment in p.findall(namespace + "commentRangeStart"):
    #         print(comment.tag, comment.attrib, comment.text)
    #     for comment in p.findall(namespace + "commentRangeEnd"):
    #         print(comment.tag, comment.attrib, comment.text)
    # # word
    # document = Document("sample.docx")
    # print(document)
    # print(document.core_properties.comments)


def parse_green(green_list):
    global err
    global info
    # print(green_list)
    no_error = True
    full_str = ""
    stack = []
    source = []
    for line in green_list:
        if len(line) != 0:
            # 【】在每个地方只能出现偶数次
            cnt = line.count("【") + line.count("】")
            if cnt == 0:
                no_error = False
                err.append('序号错误: "' + line + '"不是序号，请使用<x.x>格式标注')
            elif cnt % 2 != 0:
                no_error = False
                err.append('序号错误: "' + line + '"标注不完整')
            else:
                # 分割长度为4的
                a = line.split("】")
                a.pop()
                a = [i + "】" for i in a]
                source += a

    # print("sources", source)
    for item in source:
        if len(stack) == 0:
            stack.append(item)
        elif item != stack[-1]:
            stack.append(item)
        else:
            stack.pop()
    if len(stack) == 0:
        return 0
    else:
        # 找到孤立的那个数
        orphan = ""
        for item in stack:
            cnt = stack.count(item)
            if cnt == 1:
                orphan += item
        no_error = False
        err.append('序号错误: 存在孤立序号: \"' + orphan + '\"未被绿色高亮，请检查')

    # print(full_str)
    # for str in full_str:
    #     if str = "【":
    #     Stack.append()


def parse_yellow(yellow_list, red_list):
    # print(yellow_list)
    global err
    global info
    # print(yellow_list)
    red_str = ''
    for red_sentence in red_list:
        red_str += red_sentence
    # print(red_str)
    for yellow_word in yellow_list:
        if len(yellow_word) != 0:
            # print(line)
            # 检查黄色是否只标注了词语或短语
            if "。" in yellow_word:  # or "？" in line or "！" in line:
                full = yellow_word.index("。")
                # print(line.index("。"))
                err.append("标注错误: 黄色只能标注短语或词语，\"" + yellow_word[full - 3:full] + '。\"处"。"被标注。')
            elif "？" in yellow_word:
                question = yellow_word.index("？")
                err.append("标注错误: 黄色只能标注短语或词语，\"" + yellow_word[question - 3:question] + '？\"处"？"被标注。')
            elif "！" in yellow_word:
                exclamatory = yellow_word.index("！")
                err.append("标注错误: 黄色只能标注短语或词语，\"" + yellow_word[exclamatory - 3:exclamatory] + '！\"处"！"被标注。')
            else:
                # 检查黄色是否只标注在红色上
                if yellow_word not in red_str:
                    # print(yellow_word)
                    err.append("标注错误：短语\"" + yellow_word + "\"未在红色重点上标注，请检查！")


# 对内容列表处理,已完成
def parse_comment(content_list, filename):
    global err
    global info
    idx = 1
    no_error = True
    for comment in content_list:
        if comment.find("小标题：") == -1:
            # 没有小标题
            no_error = False
            err.append("批注错误: 第" + str(idx) + "个批注没有'小标题：'或不完整，注意：使用中文字符")
        elif comment.find("会议纪要：") == -1:
            # 没有会议纪要
            no_error = False
            err.append("批注错误: 第" + str(idx) + "个批注没有'会议纪要：'或不完整，注意：使用中文字符")
        else:
            info = comment.split("会议纪要：")
            # print(info)
            if len(info[1]) < 5:
                # 太短了
                no_error = False
                err.append("批注错误: 第" + str(idx) + "个批注长度过短")
            # 会议纪要最后未以。结尾
            if info[1].endswith("。"):
                # 查找句子结束标志
                cnt = info[1].count("。") + info[1].count("！") + info[1].count("？")
                # print(cnt)
                if cnt > 3:
                    # 过于冗长
                    no_error = False
                    err.append("批注错误: 第" + str(idx) + "个批注会议纪要过长，请勿超过三句话。")
            else:
                # 会议纪要最后未以。结尾
                cnt = info[1].count("。") + info[1].count("！") + info[1].count("？") + 1
                # print(cnt)
                if cnt > 3:
                    # 过于冗长
                    no_error = False
                    err.append("批注错误: 第" + str(idx) + "个批注会议纪要过长，请勿超过三句话。")
        idx += 1
    #  未出错，写入txt
    if no_error:
        dir_path, file_path = os.path.split(filename)
        # print(filename)
        txt_name = filename.split("/")[-1].split(".")[0]
        # print("批注文件", filename)
        # print("dir", dir_path)
        # print("txtname", txtname)
        f = open(dir_path + "/批注信息_" + txt_name + ".txt", "w", encoding='utf-8')
        i = 0
        # print(location_list, len(location_list))
        # print(content_list, len(content_list))
        for line in content_list:
            title, brief = line.split("会议纪要")
            brief = "会议纪要" + brief
            # f.write(location_list[i] + '\n')
            f.write(title + "\n")
            f.write(brief + "\n\n")
            i += 1
        f.close()


def neighborhood(iterable):
    try:
        iterator = iter(iterable)
        prev_item = " "
        current_item = next(iterator)
        # throws StopIteration if empty.
        for next_item in iterator:
            yield prev_item, current_item, next_item
            prev_item = current_item
            current_item = next_item
        yield prev_item, current_item, None
    except StopIteration:
        pass


def str_count(artical):
    count_en = count_dg = count_sp = count_zh = count_pu = count_dg_num = count_en_num = 0  # 统一将0赋值给这5个变量
    s_len = len(artical)
    ch_pu = "，。！？【】（）"
    en_pu = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~"
    for prev, item, next in neighborhood(artical):
        # 统计英文
        if item in string.ascii_letters:
            count_en += 1
            # 统计数字的个数
            # 英文前面如果有英文标点，英文字符，数字都算一个字
            if prev in string.ascii_letters:
                pass
            elif prev.isdigit():
                pass
            elif prev in en_pu:
                pass
            else:
                count_en_num += 1
        # 统计数字的位数
        elif item.isdigit():
            count_dg += 1
            # 统计数字的个数
            # 数字前面如果有英文标点，英文字符都算一个字
            if prev.isdigit():
                pass
            elif prev in en_pu:
                pass
            elif prev in string.ascii_letters:
                pass
            else:
                count_dg_num += 1
        # 统计空格
        elif item.isspace():
            count_sp += 1
        # 统计中文
        elif item.isalpha():
            count_zh += 1
        # 最特殊的一种，在=======<B>这种情况下可以计数
        elif item == '<' and prev == '=':
            count_pu += 1
        # 统计英文标点
        elif item in en_pu:
            if prev.isdigit():
                pass
            elif prev in en_pu:
                pass
            elif prev in string.ascii_letters:
                pass
            else:
                count_pu += 1
        # 统计特殊字符
        else:
            count_pu += 1
    total_num = count_zh + count_en_num + count_pu + count_dg_num
    # print(
    #     "字数: {7}\n字符数：{6}(不计空格)\n字符数：{0}(计空格)\n中文字符：{1}\n英文字符：{2}\n空格：{3}\n数字个数：{4}\n标点符号：{5}\n".format(
    #         s_len, count_zh, count_en, count_sp, count_dg, count_pu, s_len - count_sp, total_num
    #     )
    # )
    return total_num


def complete_count(xml_file):
    f = open(xml_file, encoding="utf-8")
    xml_str = ""
    while True:
        line = f.readline()
        xml_str += line
        if not line:
            break
    # print(xml_str)
    # 正则表达式，有两种情况。如果有一部分文字被标上注释了，就会触发后一种
    pattern = r"(<w:t>)([\s\S]*?)(</w:t>)|(<w:t xml:space=\"preserve\">)([\s\S]*?)(</w:t>)"
    para_list = xml_str.split(":tab/>")
    para_content_list = []
    # print(para_list)
    for para in para_list:
        para_content = ''
        res = re.findall(pattern, para)
        if res:
            for text in res:
                # 获取起始位置
                text = list(text)
                start = 0
                if "<w:t>" in text:
                    start = text.index("<w:t>")
                elif "<w:t xml:space=\"preserve\">" in text:
                    start = text.index("<w:t xml:space=\"preserve\">")

                content = text[start + 1]
                # print(content)
                para_content += content
                para_content = para_content.replace("&lt;", "<").replace("&gt;", ">")

        para_content_list.append(para_content)
    total_cnt = 0
    for para_content in para_content_list:
        total_cnt += str_count(para_content)
    # print(total_cnt)
    f.close()
    return total_cnt


def change_file_name(filename, cnt):
    # print(filename)
    prefix = filename.split(".docx")[0]
    name_no_num = prefix.split("_字数")[0]
    try:
        os.rename(filename, name_no_num + "_字数" + str(cnt) + ".docx")
    except:
        err.append("改名错误：该文件正在使用中，无法修改")


def parse_file(filename):
    global err
    global info
    err = []
    info = []
    document_path, comments_path = unzip_file(filename)
    if document_path == "error":
        return -1
    else:
        # print(document_path)
        red, yellow, green = get_color(document_path)
        # parse_red(red)
        parse_green(green)
        parse_yellow(yellow, red)
        cnt_result = complete_count(document_path)
        # print(cnt_result)
    if comments_path != "notexist":
        get_comment_location(document_path, green)
        comment = get_comment(comments_path)
        parse_comment(comment, filename)
    change_file_name(filename, cnt_result)
    return err, info, cnt_result

# parse_file("./sample.docx")
# change_file_name("zishu.docx", 123)
# artical = read_file(filename)
# # comment = get_comment("test.xml")
# # comment_parse_res = parse_comment(comment)
# cnt_result = str_count(artical)
# print(red)
# print(artical)
# print(comment)
# print(cnt_result)
# red, green, yellow, jump = find_color(filename)
# print(red)
# print(jump)
# print(yellow)
# print("Error: 没有找到文件或读取文件失败")
