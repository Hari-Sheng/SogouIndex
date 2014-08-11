# encoding=gbk
import urllib2, re, string, time, datetime, xlwt, os, sys
reload(sys)
sys.setdefaultencoding('gbk')
CurrentDir = os.getcwd()

while 1:
    print ("=========================================================================")
    keyword = raw_input("Input the keyword (Press Enter to exit):")
    if len(keyword) == 0:
        break
    else:
        # keyword = u"小米"
        # keyword = u"UCSD"
        print ("|Connecting...")

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
        if UserIndexMatch == None:
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

            Workbook = xlwt.Workbook(encoding = 'gbk')
            Worksheet = Workbook.add_sheet(keyword,cell_overwrite_ok=True)
            Worksheet.write(0,0,label='Date')
            Worksheet.write(0,1,label='UserIndex')
            Worksheet.write(0,2,label='MediaIndex')
            for i in range(0,2):
                for j in range(1,int(DayLength)+1):
                    Worksheet.write(j,0,label=RecordDate[j-1])
                    Worksheet.write(j,1,label=int(UserIndex[j-1]))
                    Worksheet.write(j,2,label=int(MediaIndex[j-1]))

            Workbook.save(CurrentDir+'\\'+keyword+'.xls')

            print ("|Data saved")
            print ("|-------------------------------------------------------------------------")
            print ("|Keyword: " + keyword)
            print ("|Time span: "+ str(int(DayLength)) +' days')
            print ("|Excel File saved at "+CurrentDir+'\\'+keyword+'.xls')


