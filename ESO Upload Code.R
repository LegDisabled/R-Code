#Load this library before running code
library(openxlsx)

#Change file name as needed. The option statistics need to be in their own
#Excel workbook
x<-read.xlsx("Articulating 652901.3104 Option Statistics.xlsx")

#If you have renamed the template, change that here
y<-read.xlsx("ESO Template.xlsx")


#Variables a-h represent values needed for certain columns in the template
#Change them to fit the particular exam

#Analysis Name
a<-"ArticulatingBoomLoaderCertificationExamination2022September"

#Exam Name
b<-"Articulating Boom Loader Certification Examination "

#Form Name
d<-"652901.3104"

#Analysis Date
e<-as.Date("09/08/2022", format="%m/%d/%Y")

#Delivery Window Start
f<-as.Date("11/18/2021", format="%m/%d/%Y")

#Delivery Window End
g<-as.Date("9/07/2022", format="%m/%d/%Y")

#Candidates
h<-683

#Function that fills in the template with the provided data. Do not edit unless 
#columns have changed in the Option Statistics file
eso<-function(x,y,a,b,d,e,f,g,h){
  n<-nrow(x)
  for (i in 1:n){
    if (is.na(x[i,2])==TRUE){
      y[i,]<-NA
    }else if (x[i,2]=="A"){
      y[i,13]<-x[i,4]
      y[i,14]<-x[i,8]
      y[i,15]<-x[i,5]
      y[i,16]<-x[i,6]
      y[i,17]<-x[i,7]
    }else if (x[i,2]=="B"){
      y[i-1,18]<-x[i,4]
      y[i-1,19]<-x[i,8]
      y[i-1,20]<-x[i,5]
      y[i-1,21]<-x[i,6]
      y[i-1,22]<-x[i,7]
    }else if (x[i,2]=="C"){
      y[i-2,23]<-x[i,4]
      y[i-2,24]<-x[i,8]
      y[i-2,25]<-x[i,5]
      y[i-2,26]<-x[i,6]
      y[i-2,27]<-x[i,7]
    }else if (x[i,2]=="D"){
      y[i-3,28]<-x[i,4]
      y[i-3,29]<-x[i,8]
      y[i-3,30]<-x[i,5]
      y[i-3,31]<-x[i,6]
      y[i-3,32]<-x[i,7]
    }else{
      y[i,]<-NA
    }
    if (is.na(x[i,3])==TRUE){
      y[i,12]<-NA
    }else if(x[i,2]=="A"& (x[i,3]==1|x[i,3]==2)){
      y[i,12]<-"A"
    }else if(x[i,2]=="B"& (x[i,3]==1|x[i,3]==2)){
      y[i-1,12]<-"B"
    }else if(x[i,2]=="C"& (x[i,3]==1|x[i,3]==2)){
      y[i-2,12]<-"C"
    }else if(x[i,2]=="D"& (x[i,3]==1|x[i,3]==2)){
      y[i-3,12]<-"D"
    }
  }
  
  z<-is.na(y[,12])
  y_2<-y[!z,]
  
  n_2<-nrow(y_2)
  
  for (i in 1:n_2){
    if (y_2[i,12]=="A"){
      y_2[i,9]<-y_2[i,15]
      y_2[i,10]<-y_2[i,16]
      y_2[i,11]<-y_2[i,17] 
    }else if (y_2[i,12]=="B"){
      y_2[i,9]<-y_2[i,20]
      y_2[i,10]<-y_2[i,21]
      y_2[i,11]<-y_2[i,22] 
    }else if (y_2[i,12]=="C"){
      y_2[i,9]<-y_2[i,25]
      y_2[i,10]<-y_2[i,26]
      y_2[i,11]<-y_2[i,27] 
    }else if (y_2[i,12]=="D"){
      y_2[i,9]<-y_2[i,30]
      y_2[i,10]<-y_2[i,31]
      y_2[i,11]<-y_2[i,32]
    }
  }
  
  y_2$Analysis.Name<-a
  y_2$Exam.Name<-b
  y_2$Form.Name<-d
  y_2$Analysis.Date<-e
  y_2$Delivery.Window.Start<-f
  y_2$Delivery.Window.End<-g
  y_2$Candidates.Total<-h
  
  names(y_2)[names(y_2)=="Analysis.Name"]<-"Analysis Name"
  names(y_2)[names(y_2)=="Exam.Name"]<-"Exam Name"
  names(y_2)[names(y_2)=="Form.Name"]<-"Form Name"
  names(y_2)[names(y_2)=="Item.Name"]<-"Item Name"
  names(y_2)[names(y_2)=="Analysis.Date"]<-"Analysis Date"
  names(y_2)[names(y_2)=="Delivery.Window.Start"]<-"Delivery Window Start"
  names(y_2)[names(y_2)=="Delivery.Window.End"]<-"Delivery Window End"
  names(y_2)[names(y_2)=="Candidates.Total"]<-"Candidates Total"
  write.csv(y_2,"ESO Filled Template.csv", row.names = FALSE)
  return(print("A CSV file called ESO Filled Template has been saved to the working directory. Item Names still need to be filled in."))
}

#Unless variable names prior to the function changed, just run this
eso(x,y,a,b,d,e,f,g,h)
