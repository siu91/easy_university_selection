# -*- coding: utf-8 -*-

import sys
import os
import urllib2
import xml.sax
from xml.dom.minidom import parse
import xml.dom.minidom
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


class ScoreLine:
    def __init__(self):
        pass

    year = ''
    region = ''
    subject = ''
    tier = ''
    score = 0


class ProvinceScore:
    def __init__(self):
        pass

    year = 0
    region = ''
    school = ''
    subject = ''
    maxScore = 0
    avgScore = 0
    minScore = 0
    tier = ''
    hope = 0
    hot = 0


class UniversityInfo:
    def __init__(self):
        pass

    longitude = ''
    latitude = ''
    name = ''
    region = ''
    regionCode = ''
    level = ''
    hot = 0
    classes = ''
    classRank = 0
    web = ''
    code = ''


# 载入所有年份的分数线
def load_score_line(code, region):
    files = os.listdir('./resource/score_line/' + code)
    sls = {}  # year:region:subject:tier score
    for file in files:
        f = open('./resource/score_line/' + code + '/' + file)
        iter_f = iter(f);  # 创建迭代器
        for line in iter_f:
            line = ''.join(line.split())
            arr = line.split(",")
            if len(arr) < 4: continue
            sl = ScoreLine()
            sl.year = arr[0]
            sl.region = region[arr[1]]
            if arr[2] == '理科':
                sl.subject = '10035'
            elif arr[2] == '文科':
                sl.subject = '10034'
            else:
                continue

            if '一' in arr[3]:
                sl.tier = '10036'
            elif '二' in arr[3]:
                sl.tier = '10037'
            elif '三' in arr[3]:
                sl.tier = '10038'
            elif '专科' or '高职' in arr[3]:
                sl.tier = '10148'
            else:
                continue

            # 专为福建没有三本特殊处理
            if code == '10024':
                if sl.tier == '10148':
                    sl.tier = '10038'

            sl.score = int(arr[4])
            sls[sl.year + ',' + sl.region + ',' + sl.subject + ',' + sl.tier] = sl
    return sls


def load_prince_score_by_file(path):
    pss = {}
    f = open(path)  # 打开文件
    iter_f = iter(f)  # 创建迭代器
    for line in iter_f:  # 遍历文件，一行行遍历，读取文本
        line = ''.join(line.split())
        ps = ProvinceScore()
        arr = line.split(',')
        ps.year = int(arr[1])
        ps.maxScore = int(arr[5])
        ps.minScore = int(arr[6])
        ps.avgScore = int(arr[7])
        ps.tier = arr[4]
        ps.region = arr[2]
        ps.school = arr[0]
        ps.subject = arr[3]
        # 学校 年份 福建 文科 批次 = 清华大学2016年在福建地区文科第一批次招生分数线
        key = arr[0] + ',' + arr[1] + ',' + arr[2] + ',' + arr[3] + ',' + arr[4]

        pss[key] = ps

    return pss


def load_prince_score(regionCode, artsOrScienceCode):
    scorePath = './resource/spider_files/' + regionCode + '_' + artsOrScienceCode + '.score'
    scoreFile = ''
    if os.path.exists(scorePath): return load_prince_score_by_file(scorePath)
    paths = os.listdir('./resource/spider_files/' + regionCode + '/')
    pss = {}
    for path in paths:
        files = os.listdir('./resource/spider_files/' + regionCode + '/' + path)
        for file in files:
            if not os.path.isdir(file):
                if not artsOrScienceCode in file: continue
                print file
                dom = xml.dom.minidom.parse('./resource/spider_files/' + regionCode + '/' + path + '/' + file)
                root = dom.documentElement
                scores = root.getElementsByTagName("score")
                for score in scores:
                    # print score.nodeName
                    # print score.toxml()
                    yearNode = score.getElementsByTagName("year")[0]
                    year = ''
                    if len(yearNode.childNodes) > 0: year = yearNode.childNodes[0].nodeValue
                    # print (yearNode.childNodes)
                    maxScoreNode = score.getElementsByTagName("maxScore")[0]
                    maxScore = ''
                    if len(maxScoreNode.childNodes) > 0: maxScore = maxScoreNode.childNodes[0].nodeValue
                    minScoreNode = score.getElementsByTagName("minScore")[0]
                    minScore = ''
                    if len(minScoreNode.childNodes) > 0: minScore = minScoreNode.childNodes[0].nodeValue
                    avgScoreNode = score.getElementsByTagName("avgScore")[0]
                    avgScore = ''
                    if len(avgScoreNode.childNodes) > 0: avgScore = avgScoreNode.childNodes[0].nodeValue
                    tierNode = score.getElementsByTagName("rb")[0]
                    tier = ''
                    if len(tierNode.childNodes) > 0: tier = tierNode.childNodes[0].nodeValue
                    ps = ProvinceScore()
                    if not ('--' == year or '' == year): ps.year = int(year)
                    if not ('--' == maxScore or '' == maxScore): ps.maxScore = int(maxScore[0:3])
                    if not ('--' == minScore or '' == minScore): ps.minScore = int(minScore[0:3])
                    if not ('--' == avgScore or '' == avgScore): ps.avgScore = int(avgScore[0:3])
                    tier = tier.encode('utf-8')

                    tierCode = ''
                    if '一' in tier:
                        tierCode = '10036'
                    elif '二' in tier:
                        tierCode = '10037'
                    elif '三' in tier:
                        tierCode = '10038'
                    elif '专' in tier:
                        tierCode = '10148'
                    else:
                        continue

                    ps.tier = tierCode
                    ps.region = regionCode
                    ps.school = path
                    ps.subject = artsOrScienceCode

                    # 学校 年份 福建 文科 批次 = 清华大学2016年在福建地区文科第一批次招生分数线
                    key = path + ',' + year + ',' + regionCode + ',' + artsOrScienceCode + ',' + tierCode
                    scoreFile = scoreFile + key + ',' + str(ps.maxScore) + ',' + str(ps.minScore) + ',' + str(
                        ps.avgScore) + '\n'
                    pss[key] = ps

    f = open(scorePath, 'w')  # 文件句柄（放到了内存什么位置）
    f.write(scoreFile.encode('utf-8'))  # 写入内容，如果没有该文件就自动创建
    f.close()  # (关闭文件)
    return pss


def load_university_info(region):
    universityDict = {}
    f = open('./resource/university_info.csv')
    iter_f = iter(f)  # 创建迭代器
    for line in iter_f:
        if line.startswith('#'): continue

        line = ''.join(line.split())
        university = UniversityInfo()
        arr = line.split(',')
        university.longitude = arr[0]
        university.latitude = arr[1]
        university.name = arr[2]
        university.region = arr[3]
        university.regionCode = region[arr[3]]
        university.level = arr[4]
        university.hot = arr[5]
        university.classes = arr[6]
        university.classRank = arr[7]
        university.web = arr[8]
        university.code = arr[9]
        universityDict[arr[9]] = university

    return universityDict


def init_region_code():
    regionCode = {}
    f = open('./resource/region_code.csv')
    iter_f = iter(f)  # 创建迭代器
    for line in iter_f:
        line = ''.join(line.split())
        arr = line.split(',')
        regionCode[arr[0]] = arr[1]

    return regionCode


def init_code_region():
    regionCode = {}
    f = open('./resource/region_code.csv')
    iter_f = iter(f)  # 创建迭代器
    for line in iter_f:
        line = ''.join(line.split())
        arr = line.split(',')
        regionCode[arr[1]] = arr[0]

    return regionCode


def init_spider(path):
    urlSet = set()
    f = open(path)
    iter_f = iter(f)  # 创建迭代器
    for line in iter_f:
        line = ''.join(line.split())
        urlSet.add(line)

    return urlSet


def spider_university_score_line(universityInfo, regionCode, subject, tier):
    url404 = init_spider('./resource/spider_files/' + regionCode + '_404.url')
    hasSpider = init_spider('./resource/spider_files/' + regionCode + '_spider.url')
    url404Size = len(url404)
    hasSpiderSize = len(hasSpider)
    urlBase = 'http://gkcx.eol.cn/schoolhtm/scores/provinceScores[university_code]_' + regionCode + '_' + subject + '_' + tier + '.xml'
    for k in universityInfo:
        url = urlBase.replace('[university_code]', universityInfo[k].code)
        if url in url404: continue
        if url in hasSpider: continue
        print url
        req = urllib2.Request(url)
        res_data = urllib2.urlopen(req)
        # print res_data.url
        if "http://gkcx.eol.cn/404.htm" == res_data.url:
            url404.add(url)
            continue
        hasSpider.add(url)
        res = res_data.read()
        path = './resource/spider_files/' + regionCode + '/' + universityInfo[k].code
        if not os.path.exists(path): os.makedirs(path)
        xmlFile = open(
            path + '/provinceScores' + universityInfo[
                k].code + '_' + regionCode + '_' + subject + '_' + tier + '.xml', 'w')
        xmlFile.write(res)
        xmlFile.close()
        # print res

    if not len(url404) == url404Size:
        wr = ''
        for u in url404:
            wr = wr + u + '\n'
        f = open('./resource/spider_files/' + regionCode + '_404.url', 'w')  # 文件句柄（放到了内存什么位置）
        f.write(str(wr))  # 写入内容，如果没有该文件就自动创建
        f.close()  # (关闭文件)

    if not len(hasSpider) == hasSpiderSize:
        wr = ''
        for u in hasSpider:
            wr = wr + u + '\n'
        f = open('./resource/spider_files/' + regionCode + '_spider.url', 'w')  # 文件句柄（放到了内存什么位置）
        f.write(str(wr))  # 写入内容，如果没有该文件就自动创建
        f.close()  # (关闭文件)


def filterUniversity(year, region, subject, score, scoreLines, provinceScores, universityInfoDict):
    # 换算该学生在往年的分数
    # 粗暴的算法
    # 计算 考生的分数在今年高考划线的比值
    # 如 该生文科393 ，2017 划线 489,380,300 ，比值分别是393/489=0.803,393/380=1.034,393/300=1.31
    # 2016 年 划线 501,403,319,评估得分为(501*0.803+403*1.034+319*1.31)/3=412.3,评估该生在2016分数为412.3
    year = int(year)
    rate1 = score / float(scoreLines[str(year) + ',' + region + ',' + subject + ',10036'].score)
    rate2 = score / float(scoreLines[str(year) + ',' + region + ',' + subject + ',10037'].score)
    rate3 = score / float(scoreLines[str(year) + ',' + region + ',' + subject + ',10038'].score)  # 福建地区没有三本划线，取专科划线

    last1 = (scoreLines[str(year - 1) + ',' + region + ',' + subject + ',10036'].score * rate1 + scoreLines[
        str(year - 1) + ',' + region + ',' + subject + ',10037'].score * rate2 + scoreLines[
                 str(year - 1) + ',' + region + ',' + subject + ',10038'].score * rate3) / 3
    print '评估' + str(year - 1) + '分数为：' + str(last1)
    last2 = (scoreLines[str(year - 2) + ',' + region + ',' + subject + ',10036'].score * rate1 + scoreLines[
        str(year - 2) + ',' + region + ',' + subject + ',10037'].score * rate2 + scoreLines[
                 str(year - 2) + ',' + region + ',' + subject + ',10038'].score * rate3) / 3
    print '评估' + str(year - 2) + '分数为：' + str(last2)
    last3 = (scoreLines[str(year - 3) + ',' + region + ',' + subject + ',10036'].score * rate1 + scoreLines[
        str(year - 3) + ',' + region + ',' + subject + ',10037'].score * rate2 + scoreLines[
                 str(year - 3) + ',' + region + ',' + subject + ',10038'].score * rate3) / 3
    print '评估' + str(year - 3) + '分数为：' + str(last3)
    print '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%'

    result = []
    schoolSet = set()
    for k in provinceScores:
        ps = provinceScores[k]
        isFilter = True
        if ps.school in schoolSet: continue
        if not ps.subject == subject: continue
        if ps.minScore == 0 and ps.avgScore == 0 and ps.maxScore == 0: continue
        hope = 0
        if not ps.minScore == 0:
            if ps.year == year - 1:
                if last1 > ps.minScore:
                    isFilter = False
                    hope = 3
            elif ps.year == year - 2:
                if last2 > ps.minScore:
                    isFilter = False
                    hope = 2
            elif ps.year == year - 3:
                if last3 > ps.minScore:
                    isFilter = False
                    hope = 1

        if not ps.avgScore == 0:
            if ps.year == year - 1:
                if last1 > ps.avgScore:
                    isFilter = False
                    hope = 6
            elif ps.year == year - 2:
                if last2 > ps.avgScore:
                    isFilter = False
                    hope = 5
            elif ps.year == year - 3:
                if last3 > ps.avgScore:
                    isFilter = False
                    hope = 4

        if not ps.maxScore == 0:
            if ps.year == year - 1:
                if last1 > ps.maxScore:
                    isFilter = False
                    hope = 9
            elif ps.year == year - 2:
                if last2 > ps.maxScore:
                    isFilter = False
                    hope = 8
            elif ps.year == year - 3:
                if last3 > ps.maxScore:
                    isFilter = False
                    hope = 7

        if isFilter: continue
        schoolSet.add(ps.school + str(ps.year))
        ps.hope = hope
        ps.hot = universityInfoDict[ps.school].hot
        result.append(ps)
    return result


def initCustomCode():
    code = {}
    code['10035'] = '理科'
    code['10034'] = '文科'
    code['10036'] = '一本'
    code['10037'] = '二本'
    code['10038'] = '三本'
    code['10148'] = '专科'
    return code


def help(region):
    print '输入参数：省份（行政区代码） 文理科 考生分数 年份 过滤批次'
    print '省份代码对应如下'
    for k in region:
        print k + '-' + region[k]
    print '文科-10034，理科10035'
    print '10036 一本,10037 二本,10038 三本,10148 专科'

    sys.exit(-1)


# 10035 理科
# 10034 文科
#
# 10036 一本
# 10037 二本
# 10038 三本
# 10148 专科
# 10149 提前

if __name__ == "__main__":
    regionCodeDict = init_region_code()
    codeRegionDict = init_code_region()
    customCodeDict = initCustomCode()
    # print str(len(sys.argv))
    if len(sys.argv) < 5:
        help(regionCodeDict)
    filterTier = ''
    if len(sys.argv) > 5:
        filterTier = sys.argv[5].split(',')

    regionCode = sys.argv[1]  # region_code 行政区代码
    artsOrScienceCode = sys.argv[2]  # arts_or_science_code 文理科
    score = int(sys.argv[3])  # 考生分数
    year = sys.argv[4]  # 年份
    # tierCode = sys.argv[3]  # tier_code 本科层次（一本，二本，三本，提前，专科）
    print '年份：' + year
    print '地区：' + codeRegionDict[regionCode]
    print '分数：' + sys.argv[3] + ' ' + customCodeDict[artsOrScienceCode]
    if not filterTier == '':
        for f in filterTier:
            print '过滤：' + customCodeDict[f]
    print '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%'
    regionCodeDict = init_region_code()
    codeRegionDict = init_code_region()
    customCodeDict = initCustomCode()
    universityInfoDict = load_university_info(regionCodeDict)
    print '抓取高校库中所有高校在[' + codeRegionDict[regionCode] + ']地区[' + customCodeDict[artsOrScienceCode] + ']招生分数线'
    spider_university_score_line(universityInfoDict, regionCode, artsOrScienceCode, '10036')
    print '本一批次抓取完成'
    spider_university_score_line(universityInfoDict, regionCode, artsOrScienceCode, '10037')
    print '本二批次抓取完成'
    spider_university_score_line(universityInfoDict, regionCode, artsOrScienceCode, '10038')
    print '本三批次抓取完成'
    spider_university_score_line(universityInfoDict, regionCode, artsOrScienceCode, '10148')
    print '高职专科批次抓取完成'
    # spider_university_score_line(universityInfoDict, '10024', '10034', '10036')

    # 历年分数线
    print '载入[' + codeRegionDict[regionCode] + ']地区' + '历年高考分数线'
    scoreLines = load_score_line(regionCode, regionCodeDict)
    # 各学校入取分数
    print '载入全国高校在[' + codeRegionDict[regionCode] + ']地区[' + customCodeDict[artsOrScienceCode] + ']历年录取分数线'
    provinceScores = load_prince_score(regionCode, artsOrScienceCode)
    print '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%'
    universityList = filterUniversity(year, regionCode, artsOrScienceCode, score, scoreLines, provinceScores,
                                      universityInfoDict)

    print '筛选结果如下，结果将保存到./resource/result.xlsx'
    print '-----------------------------------------------------------------------------------------------------'
    title = '学校\t地区\t类别\t类别排名\t热度排名\t入取成功预测值（1-9）\t 最高分\t最低分\t平均分\t年份'
    print title
    print '-----------------------------------------------------------------------------------------------------'
    # 筛选结果保存到xls
    # 在内存中创建一个workbook对象，而且会至少创建一个 worksheet
    wb = Workbook()
    # 获取当前活跃的worksheet,默认就是第一个worksheet
    ws = wb.active
    titles = title.split('\t')
    for i in range(1, len(titles) + 1):
        ws.cell(row=1, column=i).value = titles[i - 1]

    row = 2
    for u in universityList:
        if u.tier in filterTier: continue
        colContent = universityInfoDict[u.school].name + '\t' + universityInfoDict[u.school].region + '\t' + \
                     universityInfoDict[u.school].classes + '\t' + str(
            universityInfoDict[u.school].classRank) + '\t' + str(u.hot) + '\t' + str(u.hope) + '\t' + str(
            u.maxScore) + '\t' + str(u.minScore) + '\t' + str(u.avgScore) + '\t' + str(u.year)
        print colContent
        colContents = colContent.split('\t')
        for col in range(1, len(titles) + 1):
            ws.cell(row=row, column=col).value = colContents[col - 1]
        row = row + 1
        # 保存
    wb.save(filename="./resource/result.xlsx")
