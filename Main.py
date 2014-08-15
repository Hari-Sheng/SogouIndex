# encoding=utf-8
import urllib2, re, string, time, datetime, os, sys, openpyxl
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook



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
        if os.path.isfile(CurrentDir+'\\'+name+'.xls') and (keyword!=name):
            print ""
            # print(CurrentDir+'\\'+name+'.xls')
            # WorkbookTemp = xlrd.open_workbook(CurrentDir+'\\'+name+'.xls', on_demand=True, formatting_info=True)
            # print WorkbookTemp.get_sheet(0).cell(0,0).value
            # Workbook = copy(WorkbookTemp)
            # Worksheet = Workbook.get_sheet(0)
            #print Worksheet.cell(0,0),value

        else:
            wb = Workbook( )
            ws = wb.create_sheet()
            ws.title = 'SogouIndex'
            k=0

        ws.cell(row=0,column=0+k).set_value_explicit(value=keyword)
        ws.cell(row=0,column=1+k).set_value_explicit(value=keyword)
        ws.cell(row=0,column=2+k).set_value_explicit(value=keyword)
        ws.cell(row=1,column=0+k).set_value_explicit(value='Date')
        ws.cell(row=1,column=1+k).set_value_explicit(value='UserIndex')
        ws.cell(row=1,column=2+k).set_value_explicit(value='MediaIndex')

        for i in range(0+k,2+k):
            for j in range(2,int(DayLength)+2):
                ws.cell(row=j,column=0+k).set_value_explicit(value=RecordDate[j-2])
                ws.cell(row=j,column=1+k).set_value_explicit(value=UserIndex[j-2])
                ws.cell(row=j,column=2+k).set_value_explicit(value=MediaIndex[j-2])

        wb.save(CurrentDir+'\\'+name+'.xlsx')

        print ("|Data saved")
        print ("|-------------------------------------------------------------------------")
        print ("|Keyword: " + keyword)
        if (keyword!=name):
            print ("|Saving Structure: multi-keywords")
        else:
            print ("|Saving Structure: single-keywords")
        print ("|Time span: "+ str(int(DayLength)) +' days')
        print ("|Excel File saved at "+CurrentDir+'\\'+name+'.xlsx')

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
        #print re.split('[,]*[\s]*[\r]*[\n]*[\t]*',KeywordText)
        #KeywordText = KeywordText.replace('\r',' ').replace('\n', ' ').replace('\t', ' ').strip()
        #print (KeywordText)
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

            # try to merge multi-variable into one excel file
            # elif keyword.find(",")!=-1:
            #     keywords = keyword.split(",")
            #    # os.remove(keywords[0]+'.xls')
            #     for element in keywords:
            #         search(keywords[0], element,os.getcwd())
            # elif keyword.find(" ")!=-1:
            #     keywords = keyword.split(" ")
            #     print keywords
            #     for element in keywords:
            #         search(keywords[0], element, os.getcwd())
            else:
                search(keyword, keyword,os.getcwd())