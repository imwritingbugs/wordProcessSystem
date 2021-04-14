import re
import wordProcess as WP


def mytest(xml_file):
    print(xml_file)
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
        print(para_content, WP.str_count(para_content))

    total_cnt = 0
    for para_content in para_content_list:
        total_cnt += WP.str_count(para_content)
    print(total_cnt)
    red, yellow, green = WP.get_color(xml_file)
    for red_content in red:
        print(red_content)


mytest("C:\\Users\\zhouyan\\Desktop\\新建文件夹\\word\\document.xml")
