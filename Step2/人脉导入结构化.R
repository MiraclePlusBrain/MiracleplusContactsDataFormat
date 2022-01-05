library(readxl)
library(stringr)
library(openxlsx)
template<-read_excel('模板放这里（一般不用变化）/template.xlsx')

filelist<-list.files('要格式化的文档放这里')
addinfo<-read.csv('额外添加的信息.csv',header=FALSE)
allframe2<-data.frame()
#进行格式化
for(i in 1:length(filelist)){
  resource1<-read_excel(paste0('要格式化的文档放这里/',filelist[i]))
  short<-as.data.frame(array(rep(NA,nrow(resource1)*length(names(template))),dim=c(nrow(resource1),length(names(template)))))
  names(short)<-names(template)
  note<-rep('',nrow(resource1))
  for (j in 1:ncol(resource1)){
    if (names(resource1)[j]%in% names(template)){
      short[,which(names(template)==names(resource1)[j])]<-resource1[,j]
    }
    else{
      for (k in 1:nrow(resource1)){
        note[k]<-paste0(note[k],names(resource1)[j],': ',resource1[k,j],'; ')
      }}
  }
  short$`备注`<-note
  allframe2<-rbind(allframe2,short)
}

allframe2$`MPB 备注`<-addinfo[3,2]
allframe2$`来源渠道备注`<-addinfo[2,2]
allframe2$`来源渠道`<-addinfo[1,2]

# 下面做简单清洗
allframe2$`成立时间`<-as.Date(allframe2$`成立时间`)
allframe2$`成立时间`<-as.character(allframe2$`成立时间`)
allframe2$`首次融资时间`<-as.Date(allframe2$`首次融资时间`)
allframe2$`首次融资时间`<-as.character(allframe2$`首次融资时间`)
allframe2$上一轮融资至今时长<-as.Date(allframe2$上一轮融资至今时长)
allframe2$上一轮融资至今时长<-as.character(allframe2$上一轮融资至今时长)

HH<-allframe2
# 做轮次清洗
HH$`当前轮次`<-as.character(HH$`当前轮次`)
HH$`当前轮次`<-str_trim(HH$`当前轮次`)
for (i in which(is.na(HH$`当前轮次`)==FALSE)){
  if (str_detect(HH$`当前轮次`[i],'种子')){
    HH$`当前轮次`[i]<-'种子'
  }
  if (str_detect(HH$`当前轮次`[i],'天使')){
    HH$`当前轮次`[i]<-'天使'
  }
  if (str_detect(HH$`当前轮次`[i],'pre-A|Pre-A')){
    HH$`当前轮次`[i]<-'Pre-A'
  }
  if (str_detect(HH$`当前轮次`[i],'A\\+')){
    HH$`当前轮次`[i]<-'A+'
  }
  if (str_detect(HH$`当前轮次`[i],'A')&(!str_detect(HH$`当前轮次`[i],'pre|Pre'))&(!str_detect(HH$`当前轮次`[i],'\\+'))){
    HH$`当前轮次`[i]<-'A'
  }
  if (str_detect(HH$`当前轮次`[i],'战略融资|并购|股权融资|定向增发|股权转让|官方披露|债权融资|保密|未知|补助|债务融资|可转债|私募股权|非股权援助')){
    HH$`当前轮次`[i]<-'不明'
  }
  if (str_detect(HH$`当前轮次`[i],'pre-B|Pre-B')){
    HH$`当前轮次`[i]<-'Pre-B'
  }
  if (str_detect(HH$`当前轮次`[i],'B')&(!str_detect(HH$`当前轮次`[i],'pre|Pre'))&(!str_detect(HH$`当前轮次`[i],'\\+'))){
    HH$`当前轮次`[i]<-'B'
  }
}
HH$`当前轮次`[HH$`当前轮次`==''|is.na(HH$`当前轮次`)]<-'不明'
HH<-HH[!str_detect(HH$`当前轮次`,'C|上市|D|IPO|B|拟收购'),]

#手机号处理
HH$`手机号`<-as.character(HH$`手机号`)
HH$`手机号`<-str_trim(HH$`手机号`)
HH$`手机号`<-str_replace_all(HH$`手机号`,';-','')
HH$`手机号`<-str_replace_all(HH$`手机号`,' ','')
HH$`手机号`[HH$`手机号`==''|HH$`手机号`=='-']<-NA
if(addinfo[4,2]=='是'){
#下面对没有加上手机号的数据统一加识别码
HH$`手机号`[which(is.na(HH$手机号))]=paste0('无手机号',strftime(Sys.time(),'%m%d%H%M%S'),1:length(which(is.na(HH$手机号))))
}else{HH<-HH[!is.na(HH$`手机号`),]}

#对没有姓名的做填充
HH$`姓名`[which(is.na(HH$`姓名`)==TRUE)]<-'未知'

#批量变更奇绩负责人
if(addinfo[5,2]!='否'){
  HH$`奇绩负责人`=addinfo[5,2]}

#做关键词筛选
if (addinfo$V2[6]=='是'){
library(rjson)
keyword=fromJSON(file='key_word.json')
word=paste(c(keyword$unuseful_key_words$行业,keyword$unuseful_key_words$公司),collapse='|')
HH$公司注册名称[is.na(HH$公司注册名称)]=''
HH2<-HH[!str_detect(HH$公司注册名称,word),]
HH2$一句话介绍[is.na(HH2$一句话介绍)]=''
HH2<-HH2[!str_detect(HH2$一句话介绍,word),]
HH2$行业赛道[is.na(HH2$行业赛道)]=''
HH2<-HH2[!str_detect(HH2$行业赛道,word),]
#输出通过筛选的
HH2$来源渠道备注<-paste0('pass_',HH2$来源渠道备注)
HH3=HH[is.na(match(HH$公司注册名称,HH2$公司注册名称)),]
HH3$来源渠道备注<-paste0('fail_',HH3$来源渠道备注)
#导出数据
today<-strftime(Sys.time(),'%m月%d日%H点')
len1<-nrow(HH2)
write.xlsx(HH2,paste0('MPB_',today,'_pass筛选_',len1,'.xlsx'),na='',sep = ',',row.names=FALSE)
len2<-nrow(HH3)
write.xlsx(HH3,paste0('MPB_',today,'_fail筛选_',len2,'.xlsx'),na='',sep = ',',row.names=FALSE)
}else{
today<-strftime(Sys.time(),'%m月%d日%H点')
len<-nrow(HH)
write.xlsx(HH,paste0('MPB_',today,'_',len,'.xlsx'),na='',sep = ',',row.names=FALSE)
}
