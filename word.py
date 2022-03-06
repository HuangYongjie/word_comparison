from docx import Document
from PyQt5.QtWidgets import QTextEdit
import re, sys, datetime


def getText(wordname):
    d = Document(wordname)
    texts = []
    for para in d.paragraphs:
        texts.append(para.text)
    return texts


def is_Chinese(word):
    for ch in word:
        if '\u4e00' <= ch <= '\u9fff':
            return True
    return False


def msplit(s, seperators=',|\.|\?|，|。|？|！'):
    return re.split(seperators, s)


def readDocx(t, docfile):
    myprint(t, '*' * 80)
    myprint(t, '文件'+str(docfile)+'加载中……')
    t1 = datetime.datetime.now()
    paras = getText(docfile)
    segs = []
    for p in paras:
        temp = []
        for s in msplit(p):
            if len(s) > 2:
                temp.append(s.replace(' ', ""))
        if len(temp) > 0:
            segs.append(temp)
    t2 = datetime.datetime.now()
    myprint(t, '加载完成，用时: '+str(t2 - t1))
    showInfo(t, segs, docfile)
    return segs


def showInfo(t, doc, filename='filename'):
    chars = 0
    segs = 0
    for p in doc:
        for s in p:
            segs = segs + 1
            chars = chars + len(s)
    myprint(t, '段落数: {0:>8d} 个。'.format(len(doc)))
    myprint(t, '短句数: {0:>8d} 句。'.format(segs))
    myprint(t, '字符数: {0:>8d} 个。'.format(chars))


def compareParagraph(t, doc1, i, doc2, j, min_segment=5):
    """
    功能为比较两个段落的相似度，返回结果为两个段落中相同字符的长度与较短段落长度的比值。
    :param p1: 行
    :param p2: 列
    :param min_segment = 5: 最小段的长度
    """
    p1 = doc1[i]
    p2 = doc2[j]
    len1 = sum([len(s) for s in p1])
    len2 = sum([len(s) for s in p2])
    if len1 < 10 or len2 < 10:
        return []

    list = []
    for s1 in p1:
        if len(s1) < min_segment:
            continue;
        for s2 in p2:
            if len(s2) < min_segment:
                continue;
            if s2 in s1:
                list.append(s2)
            elif s1 in s2:
                list.append(s1)

    # 取两个字符串的最短的一个进行比值计算
    count = sum([len(s) for s in list])
    ratio = float(count) / min(len1, len2)
    if count > 10 and ratio > 0.1:
        myprint(t, ' 发现相同内容 '.center(80, '*'))
        myprint(t, '文件1第{0:0>4d}段内容：{1}'.format(i + 1, p1))
        myprint(t, '文件2第{0:0>4d}段内容：{1}'.format(j + 1, p2))
        myprint(t, '相同内容：'+str(list))
        myprint(t, '相同字符比：{1:.2f}%\n相同字符数： {0}\n'.format(count, ratio * 100))
    return list

def myprint(testSqare, ss):
    testSqare.append(str(ss)+'\n')

def wordrun(file1, file2, t):
    doc1 = readDocx(t, file1)
    doc2 = readDocx(t, file2)

    myprint(t, '开始比对...'.center(80, '*'))
    t1 = datetime.datetime.now()
    for i in range(len(doc1)):
        if i % 100 == 0:
            myprint(t, '处理进行中，已处理段落 {0:>4d} (总数 {1:0>4d} ） '.format(i, len(doc1)))
        for j in range(len(doc2)):
            compareParagraph(t, doc1, i, doc2, j,)

    t2 = datetime.datetime.now()
    myprint(t, '\n比对完成，总用时: '+str(t2 - t1))