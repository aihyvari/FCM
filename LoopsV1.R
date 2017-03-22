#rm(list=ls())
options(warn = -1)

#########
library(openxlsx)
library(FCMapper)
library(igraph)
##Inputs are in Loops.xlsx Excel file
nimi<-as.character("Loops.xlsx")
##############################################################################
raw<-read.xlsx(xlsxFile=nimi,sheet=1, startRow=1, colNames=FALSE,
               detectDates=FALSE, skipEmptyRows = FALSE)
#################
#POIMITAAN HALUTUT 
node<-as.numeric(raw[2,3])
pituus<-as.numeric(raw[3,3])+1
#lkm<-as.numeric(raw[5,3])
#
picW<-as.numeric(raw[5,3])
picH<-as.numeric(raw[6,3])

#Matriisi
mat<-raw[8:nrow(raw),3:ncol(raw)]
mat<-data.matrix(mat)
col<-length(na.omit(mat[,2]))
mat<-mat[1:col,1:col]
colnames(mat)<-raw[8:(7+col),2]
rownames(mat)<-colnames(mat)
####################################################
#STUDY CONCEPT MATRIX USING FCMapper
mi<-matrix.indices(mat)
ci<-concept.indices(mat, colnames(mat))
res1=nochanges.scenario(mat,iter=60,colnames(mat))
#dev.off()
lkm<-mi[1,2]*500
######################################################
k<-which(mat!=0, arr.ind=F) #etsi nonzero solut
paramlkm=length(k)
k2<-mat[k]
mat2<-matrix(0, ncol=col, nrow=col)
mat2<-replace(mat2, k, 1)
colnames(mat2)<-rownames(mat2)<-colnames(mat)
#############################################################################
#GRAPH OBJECT
g1<-graph_from_adjacency_matrix(mat2, mode = c("directed"), weighted = TRUE, diag = TRUE,
                            add.colnames = NULL, add.rownames = NA)

n1 <- neighbors(g1, node, mode=c("out"))
N<-length(n1)
if (N==0){
  print("**NO PATHS**")
}

##################################################
#ERI POLUT

apupaths<-matrix(0, ncol=pituus, nrow=lkm)
for (i in 1:lkm){
  rw<-random_walk(g1, start = node, steps = pituus, stuck = c("return"))
  length(rw)<-pituus
apupaths[i,] <- rw
} 

apupaths<-unique(apupaths[,1:ncol(apupaths)])
paths<-matrix(NA, ncol=pituus, nrow=1)

for (j in 1:nrow(apupaths)){
if ((apupaths[j,1])%in%(apupaths[j,2:ncol(apupaths)])){
  paths<-rbind(paths, apupaths[j,])
  }
}

paths<-paths[2:nrow(paths),]
###################################
paths2<-matrix(NA, ncol=pituus, nrow=nrow(paths))

for(l in 1:nrow(paths)){
ans<-rep(NA, pituus)
ans[1]<-node
  for (k in 2:pituus){
  ans[k]<-paths[l,k]
if (paths[l,k]==node){
  break
}
}
  paths2[l,]<-ans
}

paths2<-unique(paths2[,1:pituus])  
riviS<-apply(paths2[,2:ncol(paths2)], 1, table)
riviS<-lapply(riviS, max)
s<-rep(NA, length(riviS))
for (h in 1:length(riviS)){
  if (riviS[[h]][1]>1){
  s[h]<-0  
  }else{
    s[h]<-1
  }
}
paths2<-paths2[s==1,]
paths3<-as.data.frame(paths2)
for (m in 1:pituus){
paths3[,m] <- factor(paths3[,m],
                    levels = c(seq(from=1, to=col, by=1)),
                    labels = c(colnames(mat)))
}
colnames(paths3)<-paste("STEP ", seq(0,pituus-1,1))
#tkplot(g1)
#rglplot(g1, vertex.color="grey", vertex.label.dist=0.5, vertex.label.color="red",
#        edge.color="black", edge.arrow.size=1.5)

###########
#all_simple_paths(g1, node)
################################################################
simpleg<-function (matrix, concept.sizes, concept.names) 
{
  matrix.plot = igraph::graph.adjacency(matrix, mode = "directed", 
                                        weighted = T)
  igraph::V(matrix.plot)$size = concept.sizes*20
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
png(filename = "Graph.png", width = picW, height = picH, units="cm", res=300)
par(mar=c(1,1,1,1))
devAskNewPage(ask = FALSE)
try(simpleg(mat, concept.sizes = res1$Equilibrium_value, colnames(mat)))
dev.off()
########################################################################
#CREATE EXCEL 
###
wb2 = loadWorkbook(nimi)
header1<-createStyle(fontSize=13, textDecoration="bold", halign="right",fgFill="grey", wrapText=TRUE)
koro<-createStyle(fontSize=13, textDecoration="bold", halign="right", wrapText=TRUE)
#########################
####RESULTS SHEET
try(removeWorksheet(wb2, sheet="Results"), silent=TRUE)
try(addWorksheet(wb2, "Results"), silent=TRUE)
setColWidths(wb2, "Results", cols=c(1,pituus+4), widths = 15)
setColWidths(wb2, "Results", cols=c(2:pituus+2), widths = 13)

writeData(wb2,sheet="Results", x="All unique paths", startRow=2, startCol=1)
writeData(wb2,sheet="Results", x=paths3,startRow=3, startCol=2, colNames=TRUE, rowNames=TRUE)

addStyle(wb2, sheet="Results", koro, rows=c(1:(2+nrow(paths3))), cols=1:2, gridExpand = TRUE)
addStyle(wb2, sheet="Results", header1, rows=3, cols=2:(2+pituus), gridExpand = TRUE)
addStyle(wb2, sheet="Results", koro, rows=2, cols=pituus+4, gridExpand = TRUE)

##############################################
#if(graph==1){
writeData(wb2,sheet="Results", x="Concept matrix as a graph", startRow=2, startCol=pituus+4)
writeData(wb2,sheet="Results", x="Positive conn. = Black arrow, Negative = Red arrow", startRow=2, 
          startCol=pituus+6)

  insertImage(wb2, sheet="Results", file="Graph.png", width = picW, height=picH, startRow= 3,
              startCol=pituus+4, units = "cm", dpi=300)
#}

options(warn = 1)
saveWorkbook(wb2, file=nimi, overwrite=TRUE)
###
cat("\n", "Done! Go check Excel output")






