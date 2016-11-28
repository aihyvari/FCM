options(warn = -1)
rm(list=ls())

library(FCMapper)
library(openxlsx)

###################################################################################
#YOU DON'T HAVE TO CHANGE ANYTHING DOWN BELOW  
#Tolerance for convergence
tol<-0.01
raw<-read.xlsx(xlsxFile="Basic.xlsx",sheet="BasicModel", startRow=1, colNames=TRUE,
               detectDates=FALSE, skipEmptyRows = FALSE)

#POIMITAAN HALUTUT PARAMETRIT
lambda<-as.numeric(raw[1,3])
maxit<-as.numeric(raw[3,3])
set1<-as.numeric(raw[6,3])
set2<-as.numeric(raw[7,3])
#set3<-as.numeric(raw[8,3])
picW<-as.numeric(raw[9,3])
picH<-as.numeric(raw[10,3])

init<-raw[14,3:ncol(raw)]
init<-init[,!is.na(init)]
init<-as.numeric(init)

d<-raw[20:42,3:ncol(raw)]
d<-data.matrix(d)
d<-matrix(d[!is.na(d)], ncol=length(init))

colnames(d)<-raw[20:(19+length(init)),2]
rownames(d)<-colnames(d)

#STUDY CONCEPT MATRIX USING FCMapper
mi<-matrix.indices(d)
ci<-concept.indices(d, colnames(d))
#PICTURE
##########################################################################################
simpleg<-function (matrix, concept.sizes, concept.names) 
{
  matrix.plot = igraph::graph.adjacency(matrix, mode = "directed", 
                                        weighted = T)
  igraph::V(matrix.plot)$size = concept.sizes * 25
  igraph::E(matrix.plot)$color = ifelse(igraph::E(matrix.plot)$weight < 
                                          0, "red", "black")
  edge.labels = ifelse(igraph::E(matrix.plot)$weight < 0, "-", 
                       "+")
  edge.curved = rep(0, length(igraph::E(matrix.plot)))
  i = 1
  for (x in 1:length(matrix[1, ])) {
    for (y in 1:length(matrix[1, ])) {
      if (matrix[x, y] != 0 & matrix[y, x] != 0) {
        edge.curved[i] = 0.5
      }
      if (matrix[x, y] != 0) {
        i = i + 1
      }
    }
  }
  igraph::plot.igraph(matrix.plot, edge.width = abs(igraph::E(matrix.plot)$weight * 
                                                      5), vertex.color = "grey", vertex.label.color = "blue", 
                      vertex.label = concept.names, vertex.label.dist = 0, 
                      edge.curved = edge.curved)
}
#######################################################################
if(set2==1){
res1=nochanges.scenario(d,iter=maxit+5,colnames(d))
dev.off()
png(filename = "Graafi.png", width = 1, height = 1, units="cm", res=100)
devAskNewPage(ask = FALSE)
dev.off()  
png(filename = "Graafi.png", width = 30, height = 30, units="cm", res=300)
  par(mar=c(1,1,1,1))
  devAskNewPage(ask = FALSE)
  try(simpleg(d, concept.sizes = res1$Equilibrium_value, colnames(d)))
  dev.off()
  #try(graph.fcm(d, concept.sizes = res1$Equilibrium_value, colnames(d)), silent=TRUE) 
}
########################################################################
#ADD 1 INTO DIAGONAL (self-loop)
d<-d+diag(x=1, nrow=nrow(d), ncol=ncol(d))
check<-capture.output(check.matrix(d))

###########################################################################

#CREATE EMPTY MATRIX TO STORE THE RESULTS
iter <- matrix(data=NA,nrow=maxit+1,ncol=ncol(d))
########
#First row contains initial vector
iter[1,]<-init

for (i in 2:nrow(iter)){
    for (j in 1:ncol(d)){
    #taytto
    iter[i,j]<-1/(1+exp(-lambda*((iter[i-1,]%*%d[,j]))))
  }
}  

colnames(iter)<-colnames(d)

################################################
##DIFFERENCES IN STEPS
#
diff <- matrix(data=NA,nrow=nrow(iter),ncol=ncol(iter)+1)
diff[,1]<- rep(0:maxit)
colnames(diff)<-c("iter", colnames(d))
#
for (i in 2:maxit){
  for (j in 1:ncol(iter)){
    #taytto
    if (abs(iter[i,j]-iter[i-1,j])<tol){
      diff[i,j+1]<-0}
    else{
      diff[i,j+1]<-1
    }
  }  
}
####################
#NUMBER OF STEPS REQUIRED FOR CONVERGENCE
ARS<-colSums(diff[,2:ncol(diff)], na.rm=TRUE)
ARS<-ifelse(apply(iter[(nrow(iter)-5):nrow(iter),], 2, sd)<tol, ARS, "NO FP?")
ARS<-matrix(ARS, nrow=1, dimnames=list("ARS",colnames(d)))
ARS_all<-round(mean(ARS, na.rm=TRUE),2)
ARS_all<-matrix(ARS_all, nrow=1, dimnames=list("ARS_all",""))
#RESULT FOR PRINTING
result<-as.data.frame(iter)
result<-round(result,2)
result<- cbind(seq(0, maxit), result)
colnames(result)<-c("iteration", colnames(d))
res2<-result[-(3:nrow(result)-1),]
last<-result[nrow(result),-1]
###############################################
#PICTURES BY COLUMN - NOTE: produces multiple pictures
#NOTE DIFFERING Y-AXIS SCALE
png(filename = "Conv.png", width = 1, height = 1, units="cm", res=100)
devAskNewPage(ask = FALSE)
dev.off()


png(filename = "Conv.png", width = picW, height = picH, units="cm", res=300)
devAskNewPage(ask = FALSE)
rivi<-ifelse(length(init)>4, 2, 1)
rivi<-ifelse(rivi==2&length(init)>8, 3, rivi)
rivi<-ifelse(rivi==3&length(init)>12, 4, rivi)

par(mfrow=c(rivi,ceiling(ncol(d)/rivi)))
for ( i in seq(1,ncol(iter),1) ){ 
plot(seq(0,maxit),iter[0:nrow(iter),i], ylab="", xlab="iteration", 
       main=substitute(paste('Iterations of ', a), 
                       list(a=colnames(iter)[i])),type="l", cex.main=2, cex.axis=1.3)
  grid(7, 7, lwd = 2)}

dev.off()
#########################################################
#PIENET KUVAT
png(filename = "pikku1.png", width = length(init)*130, height = 150)
devAskNewPage(ask = FALSE)
par(mfrow=c(1,length(init)), mar=c(1,1.5,1,1))
for ( i in seq(1,ncol(iter),1) ){ 
  plot(seq(0,maxit),iter[0:nrow(iter),i], ylab="", xlab="", type="l", ylim=c(0, 1))
grid(7, 7, lwd = 1.2)}
dev.off()
########################################################

#CREATE EXCEL 
###
wb = loadWorkbook("Basic.xlsx")
oldw <- getOption("warn")
options(warn = -1)

try(removeWorksheet(wb, sheet="DetailedResults"))
deleteData(wb, sheet=1, cols=1:100, rows=60:159, gridExpand = TRUE)
addWorksheet(wb, "DetailedResults")
try(addWorksheet(wb, "Summary"), silent=TRUE)
setColWidths(wb, "BasicModel", cols=3:25, widths = 11.95)

setColWidths(wb, "DetailedResults", cols=1, widths = 25)
setColWidths(wb, "DetailedResults", cols=2:21, widths = 11)

header1<-createStyle(fontSize=13, textDecoration="bold", halign="right",fgFill="grey", wrapText=TRUE)
addStyle(wb, sheet=1, header1, rows=c(55,60), cols=2:(3+length(init)), gridExpand = TRUE)
koro<-createStyle(fontSize=13, textDecoration="bold", halign="right", wrapText=TRUE)
addStyle(wb, sheet=1, koro, rows=63, cols=2:3, gridExpand = TRUE)

writeData(wb,"BasicModel", check, startRow=50, startCol=2)
#writeData(wb,"BasicModel", x=t(FIXP), startRow=65, startCol=2, header=TRUE, rownames=c("FIXP"))
writeData(wb,"BasicModel", x=last, startRow=17, startCol=3, colNames = FALSE)
writeData(wb,"BasicModel", x=res2, startRow=55, startCol=2)
writeData(wb,"BasicModel", x=(ARS),startRow=60, startCol=2, colNames=TRUE, rowNames=TRUE)
writeData(wb,"BasicModel", x=ARS_all,startRow=63, startCol=2, colNames=FALSE, rowNames = TRUE)
writeData(wb,"DetailedResults",x=result, startRow=2, startCol=1)
writeData(wb,"DetailedResults",x= mi,startRow=10+maxit, startCol = 1)
writeData(wb,"DetailedResults", x=ci,startRow=25+maxit, startCol=1)
####KUVAT
insertImage(wb, sheet="BasicModel", file="pikku1.png", width = length(init)*315, 
            height=270, units="px", startRow= 18, startCol=3, dpi=300)

insertImage(wb, sheet="DetailedResults", file="Conv.png", width = picW, height=picH, startRow= maxit+length(init)+30,
            startCol=1, units = "cm", dpi=300)
if(set2==1){
insertImage(wb, sheet="DetailedResults", file="Graafi.png", width = 45, height=35, startRow= maxit+length(init)+85,
            startCol=1, units = "cm", dpi=300)
}
addStyle(wb, sheet="DetailedResults", header1, rows=c(2,(10+maxit),(25+maxit)), cols=1:(1+length(init)), 
         gridExpand = TRUE)

#hea<-createStyle(numFmt="GENERAL", halign="right", wrapText=TRUE)

##############################
set3<-as.numeric(raw[6,3])
roundUp <- function(x,to=10)
{
  to*(x%/%to + as.logical(x%%to))
}
if(set1==1){
  mitta<-try(read.xlsx(xlsxFile="Basic.xlsx",sheet="Summary", startRow=1, colNames=TRUE,
                 detectDates=FALSE, skipEmptyRows = FALSE), silent=TRUE)
  if(is.null(mitta)|class(mitta)=="try-error"){
    #addStyle(wb, sheet="Summary", hea, rows=1:40, cols=2, gridExpand = TRUE)
    addStyle(wb, sheet="Summary", header1, rows=c(2:(1+length(init))), cols=2, gridExpand = TRUE)
    writeData(wb,"Summary", x=d,startRow=1, startCol=2, colNames=TRUE, rowNames=TRUE)
    writeData(wb,"Summary", x="lambda",startRow=23, startCol=2, colNames=TRUE, rowNames=TRUE)
    writeData(wb,"Summary", x=lambda,startRow=23, startCol=3, colNames=TRUE, rowNames=TRUE)
    writeData(wb,"Summary", x=res2, startRow=25, startCol=2)
    writeData(wb,"Summary", x=(ARS),startRow=30, startCol=2, colNames=TRUE, rowNames=TRUE)
    writeData(wb,"Summary", x=ARS_all,startRow=33, startCol=2, colNames=FALSE, rowNames = TRUE)
    addStyle(wb, sheet="Summary", header1, rows=c(1,25,30), cols=2:(3+length(init)), gridExpand = TRUE)
    addStyle(wb, sheet="Summary", koro, rows=c(23,33), cols=2:3, gridExpand = TRUE)
    
    insertImage(wb, sheet="Summary", file="pikku1.png", width = length(init)*315, 
                height=270, units="px", startRow= 28, startCol=length(init)+6, dpi=300)
#    if(set2==1){
#      insertImage(wb, sheet="Summary", file="Graafi.png", width = 25, height=15, startRow= 2,
#                  startCol=length(init)+6, units = "cm", dpi=300)
#    }
  }else{
 
  ajo<-roundUp(nrow(mitta))
  addStyle(wb, sheet="Summary", header1, rows=ajo-2, cols=1:40, gridExpand = TRUE)
  #addStyle(wb, sheet="Summary", hea, rows=ajo:2*ajo, cols=2, gridExpand = TRUE)
  addStyle(wb, sheet="Summary", header1, rows=c(ajo:(ajo+length(init))), cols=2, gridExpand = TRUE)
  
  writeData(wb,"Summary", x=d,startRow=ajo, startCol=2, colNames=TRUE, rowNames=TRUE)
  writeData(wb,"Summary", x=lambda,startRow=22+ajo, startCol=3, colNames=TRUE, rowNames=TRUE)
  writeData(wb,"Summary", x="lambda",startRow=22+ajo, startCol=2, colNames=TRUE, rowNames=TRUE)
  writeData(wb,"Summary", x=res2, startRow=24+ajo, startCol=2)
  writeData(wb,"Summary", x=(ARS),startRow=29+ajo, startCol=2, colNames=TRUE, rowNames=TRUE)
  writeData(wb,"Summary", x=ARS_all,startRow=32+ajo, startCol=2, colNames=FALSE, rowNames = TRUE)
  addStyle(wb, sheet="Summary", header1, rows=c(ajo), cols=1:(3+length(init)), gridExpand = TRUE)
  addStyle(wb, sheet="Summary", header1, rows=c(24+ajo,29+ajo), cols=2:(3+length(init)), gridExpand = TRUE)
  writeData(wb,"Summary", x=round((ajo/38)+1,0),startRow=ajo, startCol=1, colNames=FALSE, rowNames=FALSE)
  addStyle(wb, sheet="Summary", koro, rows=c(22+ajo,32+ajo), cols=2:3, gridExpand = TRUE)
  #if(set2==1){
  #  insertImage(wb, sheet="Summary", file="Graafi.png", width = 25, height=15, startRow= ajo,
  #              startCol=length(init)+6, units = "cm", dpi=300)
  #}
  insertImage(wb, sheet="Summary", file="pikku1.png", width = length(init)*315, 
              height=270, units="px", startRow= 27+ajo, startCol=length(init)+6, dpi=300)
  }
}

options(warn = 1)
saveWorkbook(wb, file="Basic.xlsx", overwrite=TRUE)
###

print("Done! Go check Excel output")