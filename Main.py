# encoding=utf-8
import urllib2, re, string, time, datetime, xlwt, os, sys, xlrd, xlutils
from xlutils.copy import copy

reload(sys)
sys.setdefaultencoding('gbk')
CurrentDir = os.getcwd()

def search(name, key, dir):
    CurrentDir = dir
    keyword = key

    url = "http://index.sogou.com/sidx?type=0&query=" + keyword.decode("gbk").encode("gb2312") +"&newstype=-1"
    req = urllib2.urlopen(url)
    # url2 = req.geturl()
    # req = urllib2.Request(url2)
    # doc = urllib2.urlopen(req)

    doc = req.read()
    # print doc

    PeriodMatch = re.search(ur'"\d{4}\W\d{1,2}\W\d{1,2}\|\d{4}\W\d{1,2}\W\d{1,2}"',doc)
    UserIndexMatch = re.search(ur'"userIndexes":"((\d)+,)+(\d)+',doc)
    MediaIndexMatch = re.search(ur"tmp.mediaIndexes =  '((\d)+,)+(\d)+'",doc)
    # print PeriodMatch.group()
    # print UserIndexMatch.group()
    # print UserIndexMatch.group()[15:]
    # print MediaIndexMatch.group()
    # print MediaIndexMatch.group()[21:-1]
    if UserIndexMatch == None or MediaIndexMatch == None:
        print ("|-------------------------------------------------------------------------")
        print ("|Keyword: " + keyword)
        print ("|Sorry, but this keyword hasn't been included by SogouIndex.")

    else:
        UserIndex = UserIndexMatch.group()[15:].split(",")
        MediaIndex = MediaIndexMatch.group()[21:-1].split(",")
        # print UserIndex
        # print MediaIndex
        # print len(UserIndex)
        # print len(MediaIndex)

        print ("|Data received.")


        # print PeriodMatch.group()[1:11]
        Start = time.mktime(time.strptime(PeriodMatch.group()[1:11],'%Y-%m-%d'))
        # print Start
        # print PeriodMatch.group()[12:-1]
        End = time.mktime(time.strptime(PeriodMatch.group()[12:-1],'%Y-%m-%d'))
        # print End
        DayLength = (End-Start)/3600/24 + 1

        if len(UserIndex)==DayLength and len(MediaIndex)==DayLength:
            print ("|Check: The data length is correct.")
        else:
            print ("|Check: Error in data length.")

        Datetime = datetime.datetime.strptime(PeriodMatch.group()[1:11],'%Y-%m-%d')
        # print StartDatetime
        RecordDate = []
        for i in range(0, int(DayLength)):
            RecordDate.append(Datetime.strftime('%Y-%m-%d'))
            Datetime = Datetime + datetime.timedelta(days=1)

        # print RecordDate
        if os.path.isfile(CurrentDir+'\\'+name+'.xlsx') and (keyword!=name):
            print(CurrentDir+'\\'+name+'.xlsx')
            WorkbookTemp = xlrd.open_workbook(CurrentDir+'\\'+name+'.xlsx', on_demand=True, formatting_info=True)
            WorkbookTemp.release_resources()
            print WorkbookTemp.get_sheet(0).cell(0,0).value
            Workbook = copy(WorkbookTemp)
            Worksheet = Workbook.get_sheet(0)
            #print Worksheet.cell(0,0),value

        else:
            Workbook = xlwt.Workbook(encoding = 'gbk')
            Worksheet = Workbook.add_sheet(keyword,cell_overwrite_ok=True)
            k=0

        Worksheet.write(1,0+k,label=keyword)
        Worksheet.write(1,1+k,label=keyword)
        Worksheet.write(1,2+k,label=keyword)
        Worksheet.write(1,0+k,label='Date')
        Worksheet.write(1,1+k,label='UserIndex')
        Worksheet.write(1,2+k,label='MediaIndex')
        for i in range(0+k,2+k):
            for j in range(2,int(DayLength)+2):
                Worksheet.write(j,0,label=RecordDate[j-2])
                Worksheet.write(j,1,label=int(UserIndex[j-2]))
                Worksheet.write(j,2,label=int(MediaIndex[j-2]))

        Workbook.save(CurrentDir+'\\'+keyword+'.xlsx')

        print ("|Data saved")
        print ("|-------------------------------------------------------------------------")
        print ("|Keyword: " + keyword)
        if (keyword!=name):
            print ("|Saving Structure: multi-keywords")
        else:
            print ("|Saving Structure: single-keywords")
        print ("|Time span: "+ str(int(DayLength)) +' days')
        print ("|Excel File saved at "+CurrentDir+'\\'+keyword+'.xlsx')

if __name__=="__main__":
    isLocal= " "
    while not (isLocal=="Y" or isLocal=="N"):
        print isLocal
        isLocal = raw_input("Read from local files?(Y/N): ")

    if isLocal=="Y":
        print ("=========================================================================")
        KeywordList = raw_input("Input the name of file: ")
        while not os.path.isfile(KeywordList+".txt"):
            KeywordList = raw_input("File ["+KeywordList+".txt] does not exist, please input the name of file again:")

        KeywordText = open(KeywordList+".txt").read()
        print re.split('[,]*[\s]*[\r]*[\n]*[\t]*',KeywordText)
        KeywordText = KeywordText.replace('\r',' ').replace('\n', ' ').replace('\t', ' ').strip()
        print (KeywordText)
        keywords = re.split('[\s]*[,]*[\s]*[\r]*[\n]*[\t]*',KeywordText)
        print keywords
        for element in keywords:
            search(element, element, os.getcwd())
        keyword = raw_input("Press Enter to exit...")
    else:
        while 1:
            print ("=========================================================================")
            keyword = raw_input("Input the keyword (Press Enter to exit): ")
            if len(keyword) == 0:
                break
            #等有机会写一个把多个变量加载到一个表格里的
            # elif keyword.find(",")!=-1:
            #     keywords = keyword.split(",")
            #    # os.remove(keywords[0]+'.xlsx')
            #     for element in keywords:
            #         search(keywords[0], element,os.getcwd())
            # elif keyword.find(" ")!=-1:
            #     keywords = keyword.split(" ")
            #     print keywords
            #     for element in keywords:
            #         search(keywords[0], element, os.getcwd())
            else:
                search(keyword, keyword,os.getcwd())