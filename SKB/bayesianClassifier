from konlpy.tag import Twitter
from openpyxl import load_workbook
from openpyxl import Workbook
from konlpy.utils import pprint
from math import log

twit = Twitter()

def readDic(textfile,wordSet):
    f_dic = open(textfile, 'r', encoding='UTF8')
    lines = f_dic.readlines()
    for line in lines:
        pos_items = twit.pos(line)
        for pos_item in pos_items:
            tmp = '/'.join(pos_item)
            if tmp not in wordSet:
                wordSet.add(tmp)
    return wordSet

def makePos(sentence):
    wordSet = set()
    pos_items = twit.pos(sentence)
    for pos_item in pos_items:
        tmp = '/'.join(pos_item)
        if tmp not in wordSet:
            wordSet.add(tmp)
    return wordSet

def getProb(sentence,posList):
    prob, listCnt = 0, len(posList)
    items = makePos(sentence)
    for item in items:
        cnt = 1/countAll
        if item in posList:
            cnt += listCnt/countAll
        prob += log(cnt/(listCnt+countAll))
    return prob

def findMaxProb(pos, neg, mid):
    list = [pos, neg, mid]
    if max(list) == pos:
        print('긍정입니다')
        return '긍정'
    elif max(list) == neg:
        print('부정입니다')
        return '부정'
    else:
        print('중립입니다')
        return '중립'

def bayesian(sheetName, cnt):
    load_sheet = load_excel[sheetName]
    for i in range(3,cnt):
         target = load_sheet.cell(row=i, column=3).value + ' ' \
                  + str(load_sheet.cell(row=i, column=4).value) + ' ' \
                  + str(load_sheet.cell(row=i, column=7).value)


    load_excel.save()

emptySet = set()
f_pos = readDic('pos_pol_word.txt',emptySet)
countPos = len(f_pos)

emptySet = set()
f_neg = readDic('neg_pol_word.txt',emptySet)
countNeg = len(f_neg)

emptySet = set()
f_mid = readDic('obj_unknown_pol_word.txt',emptySet)
countMid = len(f_mid)


countAll = 7649

load_excel = load_workbook(filename='output.xlsx', data_only=True)


List = ['SK', 'KT', 'LG', '통합']

for idx in range(4):
    load_sheet = load_excel[List[idx]]
    if idx == 0:
        k = 745
    elif idx == 1:
        k = 770
    elif idx == 2:
        k = 765
    else:
        k = 2274
    for i in range(3,k):
        print(i-2,'번째 글입니다')
        sentence = load_sheet.cell(row=i, column=3).value + ' ' + str(load_sheet.cell(row=i, column=4).value) + ' ' + str(load_sheet.cell(row=i, column=7).value)

        posProb = getProb(sentence,f_pos)
        negProb = getProb(sentence,f_neg)
        midProb = getProb(sentence,f_mid)

        load_sheet.cell(row=i, column=8, value=findMaxProb(posProb,negProb,midProb))
        load_excel.save(filename='output.xlsx')







