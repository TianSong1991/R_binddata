#####合并表#######
rm(list=ls())
library(readr)
library(rpivotTable)
library(lubridate)
library(readxl)
library(dplyr)
library(tidyr)

#---------------------合并索引---------------------
a1= list.files("E:/input")#list.files命令将input文件夹下所有文件名输入a

setwd("E:/input/")

b1= list.files(a1[1],pattern = "索引.")
n=length(a1) 

for (i in 2:n){
  m1= list.files(a1[i],pattern = "索引.")
  
  b1<-rbind(b1,m1)
}
setwd("E:/")  
dir = paste("./input/",a1,"/",b1,sep="") #用paste命令构建路径变量dir
n = length(dir) #读取dir长度，也就是文件夹下的文件个数
merge.data1 = read_excel(path= dir[1],col_names = T)#读入第一个文件内容（可以不用先读一个，但是为了简单，省去定义data.frame的时间，我选择先读入一个文件。
merge.data<-merge.data1%>%select(CUST_NUM,CUST_ID,AC_NUM,CARDHOLD,委托年月,COLL,催收记录,催记样式,外访单,外访照片,客户签署文件)

for (i in 2:n){
  
  new.data= read_excel(path= dir[i],col_names = T)
  new.data1<-new.data%>%select(CUST_NUM,CUST_ID,AC_NUM,CARDHOLD,委托年月,COLL,催收记录,催记样式,外访单,外访照片,客户签署文件)
  merge.data = rbind(merge.data,new.data1)
}

#循环从第二个文件开始读入所有文件，并组合到merge.data变量中
write.csv(merge.data,file = "./output/suoyin.csv",row.names=F)  #输出组合后的文件merge.csv到input文件
#---------------------统计个索引数-----------------
count.data1 = read_excel(path= dir[1],col_names = T)
count<-nrow(count.data1)

for (i in 2:n){
  new.data= read_excel(path= dir[i],col_names = T)
  count1<-nrow(new.data)
  count<-rbind(count,count1)
}

cc<-paste("E:/input",a1[1],"******",sep = "/")
c1= list.files(cc)
lc1=length(c1)
for (i in 2:n){
  cc1<-paste("E:/input",a1[i],"******",sep = "/")
  cm= list.files(cc1)
  lcm<-length(cm)
  lc1<-rbind(lc1,lcm)
}

counttest<-data.frame(a1,count,lc1)
names(counttest)<-c("分公司","索引量","催记量")

write.csv(counttest,file = "./output/count.csv",row.names=F)
#-----------------------外访名单------------------------
setwd("E:/input/")
wb1= list.files(a1[1],pattern = "外访名单.")

for (i in 2:n){
  wm1= list.files(a1[i],pattern = "外访名单.")
  wb1<-rbind(wb1,wm1)
}
setwd("E:/")  
wdir = paste("./input/",a1,"/",wb1,sep="") #用paste命令构建路径变量dir
wn = length(wdir) #读取dir长度，也就是文件夹下的文件个数
wmerge.data1 = read_excel(path= wdir[1],col_names = T)#读入第一个文件内容（可以不用先读一个，但是为了简单，省去定义data.frame的时间，我选择先读入一个文件。
wmerge.data<-wmerge.data1%>%select(序号,客户ID,客户姓名,客户地址,客户情况说明,客户所取得的效果,催收员,日期)

for (i in 2:n){
  
  wnew.data= read_excel(path= wdir[i],col_names = T)
  wnew.data1<-wnew.data%>%select(序号,客户ID,客户姓名,客户地址,客户情况说明,客户所取得的效果,催收员,日期)
  wmerge.data = rbind(wmerge.data,wnew.data1)
}

#循环从第二个文件开始读入所有文件，并组合到merge.data变量中
write.csv(wmerge.data,file = "./output/wf.csv",row.names=F) 