#读取XML文件范例
library(XML)
#-------------------------------------------------------
#读取xml格式gpx文件数据并解析
gpxxmlstr=xmlParse(gpxfile<-file.choose(),encoding="UTF-8") 
class(gpxxmlstr)
gpxdf<-xmlToDataFrame(gpxxmlstr) #将XML文件转换成Data.Frame
#-------------------------------------------------------
#读取XML格式KML轨迹文件数据并解析
kmlxmlstr=xmlParse(kmlfile<-file.choose(),encoding="UTF-8") 
class(kmlxmlstr)
kmldf<-xmlToDataFrame(kmlxmlstr)
#-------------------------------------------------------
  #形成根目录列表数据
xmltop = xmlRoot(xmlfile) 
class(xmltop) #查看类
xmlName(xmltop) #查看根目录名
xmlSize(xmltop) #查看根目录总数
xmlName(xmltop[[1]]) #查看子目录名
#-------------------------------------------------------

#-------------------------------------------------------