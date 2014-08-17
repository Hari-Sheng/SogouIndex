SogouIndex Fetcher
==========

##About
This program will download SogouIndex automatically for given keywords, and stored them in .xlsx files. 


##How to use
###1.Input keywords from keyboard
```
Read from local files?(Y/N): 
>>>N
Input the keyword (Press Enter to exit):
>>>GitHub
|Data received.
|Check: The data length is correct.
|Data saved
|-------------------------------------------------------------------------
|Keyword: GitHub
|Saving Structure: single-keywords
|Time span: 685 days
|Excel File saved at D:\GitHub\SogouIndex\GitHub.xlsx
=========================================================================
Input the keyword (Press Enter to exit):
```

The Excel file *GitHub.xlsx* looks like something following:

     | GitHub | GitHub |
------------- | ------ | ------ |
Date | UserIndex | MediaIndex |
2012-10-01    | 0 | 0 |
...  | ... | ... |
2014-08-14 | 102 | 1
2014-08-15 | 150 | 0
2014-08-16 | 93 | 0


###2.Read from local files
Text file *test.txt* contains keywords split by `,`, `blank` or `\n`: 
```
iphone,  BAT ,, 雷军
  小米手环
```

Now we read it:

```
Read from local files?(Y/N):
>>>Y
=========================================================================
Input the name of file:
>>>test
|Data received.
|Check: The data length is correct.
|Data saved
|-------------------------------------------------------------------------
|Keyword: iphone
|Saving Structure: multi-keywords
|Time span: 685 days
|Excel File saved at D:\GitHub\SogouIndex\test.xlsx
|-------------------------------------------------------------------------
|Keyword: BAT
|Sorry, but this keyword hasn't been included by SogouIndex.
|Data received.
|Check: The data length is correct.
|Data saved
|-------------------------------------------------------------------------
|Keyword: 雷军
|Saving Structure: multi-keywords
|Time span: 685 days
|Excel File saved at D:\GitHub\SogouIndex\test.xlsx
|Data received.
|Check: The data length is correct.
|Data saved
|-------------------------------------------------------------------------
|Keyword: 小米手环
|Saving Structure: multi-keywords
|Time span: 685 days
|Excel File saved at D:\GitHub\SogouIndex\test.xlsx
Press Enter to exit...
```
And one of the records is like:

         | iphone | iphone | 雷军 | 雷军 | 小米手环 | 小米手环 |
-------- | ------ | ------ | ---- | ---- | ------ | -------  |
Date | UserIndex | MediaIndex | UserIndex | MediaIndex | UserIndex | MediaIndex |
...  | ... | ... | ... | ... |  ... | ... |
2014-08-15 | 2996 | 1271 | 619 | 17 | 923 | 13 |



##SogouIndex
[Sogou](http://www.sogou.com/) is a popular SE(search engine) in China (also not that much as Baidu). Just as [Google Trends](http://www.google.com/trends/) and [BaiduIndex](http://index.baidu.com/), SogouIndex tries to measure popularity of given keywords based on the searching behavior of its users.


##Motivation
This toy program is done during my internship at [T.H.Capital](http://thcapital-china.com/), Jul 2014. We are using searching indexes to predict the sale volume of upcoming cellphone, the numbers of visitors to Macau, etc.





Contact Information
----------
+ Email: hsheng@pku.edu.cn
+ Weibo: [@盛浩Nihilist](http://weibo.com/u/1870340245)

