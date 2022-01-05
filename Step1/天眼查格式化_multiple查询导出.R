library(openxlsx)
library(stringr)
library(readxl)
filelist=list.files()
filelist2=filelist[str_detect(filelist,'批量查询导出数据')]

resource1<-read_excel(paste0(getwd(),'/',filelist2))
resource1$`联系电话`<-paste0(resource1$`联系电话`,';',resource1$`其他电话`)
resource1$`所属省份`<-paste0(resource1$`所属省份`,'-',resource1$`所属城市`)
resource1$`所属省份`<-str_replace_all(resource1$`所属省份`,'-','')
resource2<-resource1[,c(seq(1,3),5,9,13,19,20,seq(22,25))]
names(resource2)<-c('公司注册名称','姓名','注册资本','成立时间','城市','工商行业','参保人数','手机号','地址','官网','邮箱','经营范围')
resource1<-resource2

write.xlsx(resource1,'MPB_output.xlsx',na='',row.names = FALSE)
