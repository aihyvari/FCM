#rm(list=ls())
cat("\n","I'm computing for a while, take it easy...")
#options(warn = -1)

#########
library(openxlsx)
library(GA)
library(reshape2)

nimi<-as.character("ConMatOpt.xlsx")

##############################################################################
inPut<-read.xlsx(xlsxFile=nimi,sheet=1, startRow=1, colNames=FALSE,
               detectDates=FALSE, skipEmptyRows = FALSE)

#POIMITAAN HALUTUT 
#Alkuvektori
inval<-inPut[16,3:(ncol(inPut))]
inval<-inval[,!is.na(inval)]
inval<-as.numeric(inval)
row<-col<-length(inval)
#Tavoite
goalval<-inPut[17,3:(2+col)]
goalval<-as.numeric(goalval)
#Ohjematriisi
oldmat<-inPut[22:(22+col),3:(2+col)]
oldmat<-data.matrix(oldmat)
oldmat<-oldmat[1:col,1:col]
colnames(oldmat)<-inPut[22:(21+col),2]
rownames(oldmat)<-colnames(oldmat)
#Parametrit
lambda<-as.numeric(inPut[3,3])
simnr<-as.numeric(inPut[4,3])
pop<-as.numeric(inPut[9,3])
Min<-as.numeric(inPut[10,3])
Max<-as.numeric(inPut[11,3])
roundB<-as.numeric(inPut[12,3])

############
#Kuva-asetukset
graph<-1
summa<-1
picW<-as.numeric(inPut[7,3])
picH<-as.numeric(inPut[8,3])
####################################################

k<-which(oldmat!=0, arr.ind=F) #etsi nonzero solut, vain ne optimoidaan
paramlkm=length(k)
k2<-oldmat[k]
##########################
#diag(oldmat)<-1
oldmat2<-melt(oldmat)
oldmat2<-oldmat2[,3]
#####################################################################
guess<-rep(0.5, paramlkm)
#####################################################
findmatga<-function(param){
  
  oldmat2<-replace(oldmat2, k, param)
  mat<-matrix(oldmat2, nrow=row, ncol=col)
  
  out<-1/(1+exp(-lambda*(inval%*%mat)))
  n<-1
  while(n<=simnr){
    n<-n+1
    out<-1/(1+exp(-lambda*(out%*%mat)))
  }
  out<<-out
  rmse<-sqrt(sum((goalval-out)^2)/col) #rmse
  return(rmse)
  }


#######################################################################
fitness<- function(guess) -findmatga(guess)
lb<-rep(-1, length(k))
ub<-rep(1, length(k))


GA1 <- ga(type = "real-valued", fitness = fitness, 
          min = lb, max = ub, 
          maxiter = pop, run = 100, optim = TRUE, seed=1234)


###################################################
sol<-replace(oldmat2, k, GA1@solution)
rmseUR=-GA1@fitnessValue
rmseUR<-round(rmseUR, 4)
#matparas=GA1@solution#paras alkuvect 
matparas<-matrix(sol, nrow=row, ncol=col)
matparas<-round(matparas, 3)
matparas<-ifelse(abs(matparas)<roundB, 0, matparas)
colnames(matparas)<-rownames(matparas)<-colnames(oldmat)
###################################################################
####BASIC MODEL
#CREATE EMPTY MATRIX TO STORE THE RESULTS
iter <- matrix(data=NA,nrow=simnr+1,ncol=col)
########
#First row contains initial vector
iter[1,]<-inval

for (i in 2:nrow(iter)){
  for (j in 1:col){
    #taytto
    iter[i,j]<-1/(1+exp(-lambda*((iter[i-1,]%*%matparas[,j]))))
  }
}  

colnames(iter)<-colnames(oldmat)

################################################
##DIFFERENCES IN STEPS
#

diff <- matrix(data=NA,nrow=nrow(iter),ncol=ncol(iter)+1)
diff[,1]<- rep(0:simnr)
colnames(diff)<-c("iter", colnames(iter))
#
for (i in 2:simnr){
  for (j in 1:ncol(iter)){
    #taytto
    if (abs(iter[i,j]-iter[i-1,j])<0.01){
      diff[i,j+1]<-0}
    else{
      diff[i,j+1]<-1
    }
  }  
}
####################
#NUMBER OF STEPS REQUIRED FOR CONVERGENCE
AveReqSteps<-colSums(diff[,2:ncol(diff)], na.rm=TRUE)
AveReqSteps<-ifelse(apply(iter[(nrow(iter)-3):nrow(iter),], 2, sd)<0.01, AveReqSteps, "NO FP?")
AveReqSteps<-matrix(AveReqSteps, nrow=1, dimnames=list("AveReqSteps",colnames(oldmat)))
AveReqSteps_all<-round(mean(AveReqSteps, na.rm=TRUE),2)
AveReqSteps_all<-matrix(AveReqSteps_all, nrow=1, dimnames=list("AveReqSteps_all",""))
#PARAS RMSE
rmseparas<-sqrt(sum((goalval-iter[simnr,])^2)/col) #rmse
rmsediff<-round(rmseparas-rmseUR, 4)

##############################################################
simpleg<-function (matrix, concept.names) 
{
  matrix.plot = igraph::graph.adjacency(matrix, mode = "directed", 
                                        weighted = T)
  igraph::V(matrix.plot)$size = 12
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
  igraph::plot.igraph(matrix.plot, edge.width = abs(igraph::E(matrix.plot)$weight * 3), 
                      vertex.color = "grey", vertex.label.color = "blue", 
                      vertex.label = concept.names, vertex.label.dist = 0, 
                      edge.curved = edge.curved)
}
#######################################################################

if(graph==1){
  png(filename = "Graaf.png", type="cairo", width = 30, height = 30, units="cm", res=300)
  par(mar=c(1,1,1,1))
  devAskNewPage(ask = FALSE)
  try(simpleg(matparas, colnames(matparas)))
  dev.off()
}

################################################################

names(inval)<-colnames(oldmat)
#CREATE EXCEL 
###
wb5 = loadWorkbook(nimi)

deleteData(wb5, sheet=1, cols=1:100, rows=49:199, gridExpand = TRUE)
try(addWorksheet(wb5, "Summary"), silent=TRUE)

header1<-createStyle(fontSize=13, textDecoration="bold", halign="right",fgFill="grey", wrapText=TRUE)
koro<-createStyle(fontSize=13, textDecoration="bold", halign="right", wrapText=TRUE)
negStyle<-createStyle(fontColour="#9C0006", bgFill="red")
posStyle<-createStyle(fontColour="#006100", bgFill="green")
#########################
####FRONT SHEET
#################
#Results
try(removeWorksheet(wb5, sheet="Results"))
try(addWorksheet(wb5, "Results"), silent=TRUE)
#
setColWidths(wb5, "Results", cols=1, widths = 25)
setColWidths(wb5, "Results", cols=2:31, widths = 8)
################################PARAMS
#lambda arvo
writeData(wb5,"Results", x="lambda",startRow=2, startCol=1, colNames=TRUE, rowNames=TRUE)
writeData(wb5,"Results", x=lambda,startRow=2, startCol=2, colNames=TRUE, rowNames=TRUE)
#Mallin iteraatiot lkm
writeData(wb5,"Results", x="Number of FCM model iterations",startRow=3, startCol=1)
writeData(wb5,"Results", x=simnr,startRow=3, startCol=2)
#Param lkm
writeData(wb5,"Results", x="Number of parameters (to optimize)",startRow=4, startCol=1)
writeData(wb5,"Results", x=paramlkm,startRow=4, startCol=2)
#GA iteraatiot lkm
writeData(wb5,"Results", x="Number of GA iterations",startRow=5, startCol=1)
writeData(wb5,"Results", x=pop,startRow=5, startCol=2)
#paras RMSE
writeData(wb5,"Results", x="Optimized RMSE",startRow=6, startCol=1)
writeData(wb5,"Results", x=round(rmseparas,4),startRow=6, startCol=2)
#rmse diff
writeData(wb5,"Results", x="RMSE increase due to the rounding",startRow=7, startCol=1)
writeData(wb5,"Results", x=round(rmsediff,4),startRow=7, startCol=2)

addStyle(wb5, sheet="Results", koro, rows=c(2:7,14,41,43), cols=1:3, gridExpand = TRUE)

#############################VECTORS
#Initial
writeData(wb5,"Results", x="Given INITIAL vector",startRow=10, startCol=1)
writeData(wb5,"Results",x=t(round(inval,3)), startRow=9, startCol=2, colNames=TRUE)
#GOAL
writeData(wb5,"Results", x="Given GOAL vector", startRow=11, startCol=1)
writeData(wb5,"Results",x=t(goalval), startRow=11, startCol=2, colNames=FALSE)
#PREDICTED
writeData(wb5,"Results", x="Predicted goal vector", startRow=12, startCol=1)
writeData(wb5,"Results",x=round(t(iter[nrow(iter),]),3), startRow=12, startCol=2, colNames=FALSE)
#MAT
writeData(wb5,"Results", x="Optimized concept matrix", startRow=14, startCol=1)
writeData(wb5,"Results",x=matparas, startRow=15, startCol=1, colNames=TRUE, rowNames=TRUE)
conditionalFormatting(wb5, sheet="Results", cols=(2:(2+col)), rows=(16:(16+col)), rule="<0", style=negStyle)
conditionalFormatting(wb5, sheet="Results", cols=(2:(2+col)), rows=(16:(16+col)), rule=">0", style=posStyle)
addStyle(wb5, sheet="Results", header1, rows=c(9,15,40), cols=1:(2+col), 
         gridExpand = TRUE)
addStyle(wb5, sheet="Results", header1, rows=c(15:(15+row)), cols=1, gridExpand = TRUE)

###
#AveReqSteps
writeData(wb5,"Results", x=(AveReqSteps),startRow=40, startCol=1, colNames=TRUE, rowNames=TRUE)
writeData(wb5,"Results", x=AveReqSteps_all,startRow=43, startCol=1, colNames=FALSE, rowNames = TRUE)

if(graph==1){
  writeData(wb5,"Results", x="Optimized concept matrix as a graph", startRow=45, startCol=1)
  
  insertImage(wb5, sheet="Results", file="Graaf.png", width = 30, height=25, startRow= 47,
              startCol=1, units = "cm", dpi=300)
}

##############################
roundUp <- function(x,to=10)
{
  to*(x%/%to + as.logical(x%%to))
}
if(summa==1){
  mitta<-try(read.xlsx(xlsxFile=nimi,sheet="Summary", startRow=1, colNames=TRUE,
                       detectDates=FALSE, skipEmptyRows = FALSE), silent=TRUE)
  if(is.null(mitta)|class(mitta)=="try-error"){
    setColWidths(wb5, "Summary", cols=2, widths = 25)
    setColWidths(wb5, "Summary", cols=3:60, widths = 7)
    #lambda arvo
    writeData(wb5,"Summary", x="Lambda",startRow=2, startCol=2)
    writeData(wb5,"Summary", x=lambda,startRow=2, startCol=3)
    #lkm
    writeData(wb5,"Summary", x="Number of FCM model iterations",startRow=3, startCol=2)
    writeData(wb5,"Summary", x=simnr,startRow=3, startCol=3)
    #param
    writeData(wb5,"Summary", x="Number of parameters (to optimize)",startRow=4, startCol=2)
    writeData(wb5,"Summary", x=paramlkm,startRow=4, startCol=3)
    #Number of GA iterations
    writeData(wb5,"Summary", x="Number of GA iterations",startRow=5, startCol=2)
    writeData(wb5,"Summary", x=pop, startRow=5, startCol=3)
    #paras RMSE
    writeData(wb5,"Summary", x="Optimized RMSE",startRow=6, startCol=2)
    writeData(wb5,"Summary", x=rmseparas,startRow=6, startCol=3)
    
    writeData(wb5,"Summary", x="Given INITIAL vector",startRow=9, startCol=2)
    writeData(wb5,"Summary",x=t(round(inval,3)), startRow=8, startCol=3, colNames=TRUE)
    #GOAL
    writeData(wb5,"Summary", x="Given GOAL vector", startRow=10, startCol=2)
    writeData(wb5,"Summary",x=t(goalval), startRow=10, startCol=3, colNames=FALSE)
    #PREDICTED
    writeData(wb5,"Summary", x="Predicted goal vector", startRow=11, startCol=2)
    writeData(wb5,"Summary",x=round(t(iter[nrow(iter),]),3), startRow=11, startCol=3, colNames=FALSE)
    #Matriisi
    addStyle(wb5, sheet="Summary", header1, rows=c(14:(14+row)), cols=2, gridExpand = TRUE)
    writeData(wb5,"Summary", x="Optimized concept matrix",startRow=13, startCol=1)
    writeData(wb5,"Summary", x=matparas,startRow=14, startCol=2, colNames=TRUE, rowNames=TRUE)
    conditionalFormatting(wb5, "Summary", cols=(3:(3+col)), rows=(15:(15+col)), rule="<0", style=negStyle)
    conditionalFormatting(wb5, "Summary", cols=(3:(3+col)), rows=(15:(15+col)), rule=">0", style=posStyle)
    
    addStyle(wb5, sheet="Summary", header1, rows=c(8,14), cols=2:(3+row), gridExpand = TRUE)
    addStyle(wb5, sheet="Summary", koro, rows=2:6, cols=2:3, gridExpand = TRUE)
    #writeData(wb5,"Summary", x=".",startRow=37, startCol=2)
    
  }else{
    
    ajo<-roundUp(nrow(mitta))+3
    writeData(wb5,"Summary", x=round((ajo/40)+1,0),startRow=ajo, startCol=1)
    addStyle(wb5, sheet="Summary", header1, rows=ajo, cols=1:30, gridExpand = TRUE)
    
    ###################
    #lambda arvo
    writeData(wb5,"Summary", x="Lambda",startRow=ajo+2, startCol=2)
    writeData(wb5,"Summary", x=lambda,startRow=ajo+2, startCol=3)
    #lkm
    writeData(wb5,"Summary", x="Number of FCM model iterations",startRow=ajo+3, startCol=2)
    writeData(wb5,"Summary", x=simnr,startRow=ajo+3, startCol=3)
    #param
    writeData(wb5,"Summary", x="Number of parameters (to optimize)",startRow=ajo+4, startCol=2)
    writeData(wb5,"Summary", x=paramlkm,startRow=ajo+4, startCol=3)
    #Number of GA iterations
    writeData(wb5,"Summary", x="Number of GA iterations",startRow=ajo+5, startCol=2)
    writeData(wb5,"Summary", x=pop, startRow=ajo+5, startCol=3)
    #paras RMSE
    writeData(wb5,"Summary", x="Optimized RMSE",startRow=ajo+6, startCol=2)
    writeData(wb5,"Summary", x=rmseparas,startRow=ajo+6, startCol=3)
    
    writeData(wb5,"Summary", x="Given INITIAL vector",startRow=ajo+9, startCol=2)
    writeData(wb5,"Summary",x=t(round(inval,3)), startRow=ajo+8, startCol=3, colNames=TRUE)
    #GOAL
    writeData(wb5,"Summary", x="Given GOAL vector", startRow=ajo+10, startCol=2)
    writeData(wb5,"Summary",x=t(goalval), startRow=ajo+10, startCol=3, colNames=FALSE)
    #PREDICTED
    writeData(wb5,"Summary", x="Predicted goal vector", startRow=ajo+11, startCol=2)
    writeData(wb5,"Summary",x=round(t(iter[nrow(iter),]),3), startRow=ajo+11, startCol=3, colNames=FALSE)
    #Matriisi
    addStyle(wb5, sheet="Summary", header1, rows=c((ajo+15):(15+ajo+col)), cols=2, gridExpand = TRUE)
    writeData(wb5,"Summary", x="Optimized concept matrix",startRow=ajo+13, startCol=1)
    writeData(wb5,"Summary", x=matparas,startRow=ajo+14, startCol=2, colNames=TRUE, rowNames=TRUE)
    conditionalFormatting(wb5, "Summary", cols=(3:(3+col)), rows=((ajo+15):(ajo+15+col)), rule="<0", style=negStyle)
    conditionalFormatting(wb5, "Summary", cols=(3:(3+col)), rows=((ajo+15):(ajo+15+col)), rule=">0", style=posStyle)
    
    addStyle(wb5, sheet="Summary", header1, rows=c(ajo+8,ajo+14), cols=2:(3+col), gridExpand = TRUE)
    addStyle(wb5, sheet="Summary", koro, rows=(ajo+2):(ajo+6), cols=2:3, gridExpand = TRUE)
   
  }
}

options(warn = 1)

saveWorkbook(wb5, file=nimi, overwrite=TRUE)
###
#print(summary(GA1))
cat("\n", "Done! Go check Excel output")







