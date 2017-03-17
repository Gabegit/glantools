#' 试卷分析
#' @author lgm
#' @param file xlsx格式的成绩表，第一列为学号或姓名，中间列为个小题得分，最后一列为总分；每行为一个学生分数；文件第一行为各列名字（取什么名字无影响）；
#' @param ques_cols 小题放置的列数，如2:5，表示从第2列到5列；
#' @param fenZhi 小题按顺序的总分放在 c()中，如,c(10,20,30,40)
#' @param sheetname 工作表序号，默认为第1个。如果数字房子第一工作表中，不要设置。
#' @return results 各种结果和图表
#' @export
#' @examples
#' data = file.path(system.file(package = "glantools"),"test.xlsx")
#' shijuan_fenxi(data,ques_cols=2:5,fenZhi=c(10,20,30,40))
#' data2 = file.path(system.file(package = "glantools"),"test2.xlsx")
#' shijuan_fenxi(data2,ques_cols=2:5,fenZhi=c(30,15,30,25))
#'


shijuan_fenxi <- function(file,ques_cols=2:5,fenZhi,sheetname=1){
  #install neccesary packages
  if (!"ggplot2" %in% installed.packages()) install.packages("ggplot2")
  if (!"tidyr" %in% installed.packages()) install.packages("tidyr")
  if (!"ggrepel" %in% installed.packages()) install.packages("ggrepel")
  if (!"readxl" %in% installed.packages()) install.packages("readxl")

  suppressWarnings(suppressPackageStartupMessages({
    library(ggplot2)
    library(tidyr)
    library(ggrepel)
    library(readxl)
  }))

  # read data
  ##data <- xlsx::read.xlsx(file,sheetname,startRow = 2,header = FALSE)
  data <- read_excel(file,sheetname,col_names = FALSE,skip=1)
  data <- as.data.frame(data)
  tcol <- dim(data)[2]

  data_cut <- cut(as.numeric(data[,tcol]),breaks = c(0,59,69,79,89,100),labels=c("0-59","60-69","70-79","80-89","90-100"))
  numbers <-table(data_cut)
  options(digits = 3)
  percentage <- round(table(data_cut)/nrow(data)*100,2)
  df <- as.data.frame(cbind(numbers,percentage))
  df$names <- row.names(df) # rownames as a column

  df1 = as.data.frame(table(data_cut))

  pl <- ggplot(df,aes(x=names,y=percentage)) +
    geom_bar(aes(fill="pink"),stat="identity") +
    geom_text(vjust = -0.2,aes(label=paste(percentage,"%")))+
    theme(text=element_text(family="SimSun",size=10))+
    labs(x="全部成绩各分数段情况",y="人数百分占比")

  cat("全部成绩各分数段统计图\n")
  print(pl)


  ## 总体分析
  ## 标准差

  sd <- sd(data[,tcol])

  ## 阿尔法信度
  ## alhpa = 题数/(题数-1)*(1-各题方差之和/总分方差)
  varcovar <- var(data[,ques_cols])
  #var_sum <- sum(var(data$X2),var(data$X3),var(data$X4),var(data$X5))
  var_sum <- sum(diag(varcovar))
  k <- ncol(data) - 2
  alpha <- k/(k-1)*(1-var_sum/var(data[,tcol]))

  ##斯皮尔曼-布朗公式计算信度
  ##半个测试信度=（奇数题分数列和偶数列分数的相关系数）
  ##整个测试的信度=2*半个测试的信度/(1+半个测试的信度）
  odd <-  rowSums(data[,ques_cols[seq(1,length(ques_cols),by=2)]])
  even <- rowSums(data[,ques_cols[seq(2,length(ques_cols),by=2)]])
  half_rebi <- cor(odd,even)
  sd_re <- 2*half_rebi/(1+half_rebi)
  cat("--------------------------------------\n\n")
  cat("*****总体分析****\n")
  cat(paste(" 考试总人数：",dim(data)[1],"人;\n","全班平均分数：",round(mean(data[,tcol]),3)),"分; \n","标准差：",round(sd,3),"\n","阿尔法信度:",round(alpha,3),"\n","斯皮尔曼-布朗信度:",round(sd_re,3),"\n\n")
  cat("--------------------------------------\n")
  cat("*****小题分析****\n")

  ## 小题分析
  question <- function(ques,fenZhi){
    mean <- mean(ques)
    sd_q <- sd(ques)
    diffi <- mean/fenZhi
    ##区分度（discrimination index）= 27%高分组的中值-27%低分组中值)/总分数
    ques_sort <- sort(ques)
    stu <- length(ques)
    twenty7 <- trunc(stu*0.27)
    discr <- (median(ques_sort[(stu-twenty7):stu]) - median(ques_sort[1:twenty7]))/fenZhi
    return(list(平均分 = round(mean,2), 标准差 = round(sd_q,2), 难度系数 = round(diffi,2),区分度 = round(discr,2)))

  }
  df_timu <- data[,ques_cols]
  names(df_timu) <- paste0("第",ques_cols-1,"题")
  dq <- as.data.frame(mapply(question,df_timu,fenZhi))
  dq$names <- row.names(dq)
  row.names(dq)<- NULL
  fz <- as.data.frame(matrix(c(fenZhi,"分值"),nrow =1),stringsAsFactors = FALSE)
  names(fz) <- names(dq)
  dat <- rbind(fz,dq)
  res <- dat[,c(length(ques_cols)+1,1:(length(ques_cols)))]
  print(res)

  ## plot res
  res_long <- gather(res,questions,value,-names)
  res_long$value <- as.numeric(res_long$value)
  res_long$names <- as.factor(res_long$names)

  pl_index <- ggplot(res_long,aes(x=names,y=value,color=questions))+
    geom_point() +
    geom_text_repel(aes(label=value))+
    theme(text=element_text(family="SimSun",size=12))+
    labs(x="",y="系数值")

  cat("各小题系数比较图\n")
  print(pl_index)

  return("well done!")

}
