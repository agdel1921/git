I\import sys
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from PyQt4.QtWebKit import *
from lxml import html
import xlwt

class Render(QWebPage):  
  def __init__(self, urls, cb):
    self.app = QApplication(sys.argv)  
    QWebPage.__init__(self)  
    self.loadFinished.connect(self._loadFinished)  
    self.urls = urls  
    self.cb = cb
    self.crawl()  
    self.app.exec_()  
      
  def crawl(self):  
    if self.urls:  
      url = self.urls.pop(0)  
      print 'Downloading', url  
      self.mainFrame().load(QUrl(url))  
    else:  
      self.app.quit()  
        
  def _loadFinished(self, result):  
    frame = self.mainFrame()  
    url = str(frame.url().toString())  
    html = frame.toHtml()  
    self.cb(url, html)
    self.crawl()   

def scrape(url, htm):
     #html is returned as r.frame.toHTML() -i.e. QString
     formatted_result = str(htm.toAscii())
     
     #print formatted_result
     tree = html.fromstring(formatted_result)

     table = tree.xpath('//tbody/tr/td')
     mrt = tree.xpath('//td/a')

     bookname = (url[51:].replace('-',''))

     excel = book.add_sheet(url[51:].replace('-',''))
     cols = ['MRT','S-Condo','S-Landed','S-HDB','R-Condo','R-Landed','R-HDB']
     
     for c in range(len(cols)): excel.write(0,c,cols[c])
     #write the station names into col of excel
     for row, stn in enumerate(mrt):
          excel.write(row+1,0,stn.text)

     row=1
     col=1
     for x in range(4, len(table)):
          if len(table[x]) == 0:
               price = table[x].text
               print row,col,price
               excel.write(row,col,price)
               col = col + 1
               if (col%7 == 0):
                    col = 1
                    row = row + 1
     
url = 'http://www.srx.com.sg/mrt-home-prices/'

mrt_lines=[
'price-of-home-east-west-line',
'price-of-home-downtown-line',
'price-of-home-north-south-line',
'price-of-home-north-east-line',
'price-of-home-circle-line',
'price-of-home-thomson-line']

urls=[]
for x in mrt_lines: urls.append(url+x)

#set up excel for writing
book = xlwt.Workbook(encoding="utf-8")
r = Render(urls, cb=scrape)
book.save('MRT_Prices.xls')
