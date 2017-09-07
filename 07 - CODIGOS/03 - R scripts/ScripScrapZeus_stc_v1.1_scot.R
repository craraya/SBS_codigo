
## -------------------------------------------------------------------
## -------------------------------------------------------------------
## -------------------------------------------------------------------
## <a rel="license" href="http://creativecommons.org/licenses/by-nc-nd/4.0/">
## <img alt="Licencia Creative Commons" style="border-width:0" src="https://i.creativecommons.org/l/by-nc-nd/4.0/88x31.png" />
## </a><br />Esta obra está bajo una <a rel="license" href="http://creativecommons.org/licenses/by-nc-nd/4.0/">
## Licencia Creative Commons Atribución-NoComercial-SinDerivar 4.0 Internacional</a>.

library(RCurl)
library(XML)
library(RODBC)

## Formato de entrada
## rut  dv
## 12345456  4
## 45612379  5
## 4512  6

chn <- odbcConnect("RT_ANALYTICS_ORI")

Drut <- sqlQuery(chn, "
select rut, digito
from [BD_Campañas].[dbo].[Campaña_201707]	
where campaña like '%pros%' and campaña not like '%preaprob%'
and rut not in (select distinct Rut from [RT_ANALYTICS_ORI].[dbo].[CA_SII_Consolidado]) -- 152.719
")

#Drut <- read.table("E:/CARAYA/07 - CODIGOS/03 - R scripts/rutero_para_sii.txt"
#           , header = T
#           , sep="\t")

# Drut <- Drut[1:20000,] ## solo 10 mil

Drut.dim <- dim(Drut)
Drut[,2] <- as.character(Drut[,2])
Drut[,1] <- as.character(Drut[,1])
for(i in 1:Drut.dim[1]){
  if(Drut[i,2]=="k") {
    Drut[i,2] <- "K"
  }
}

# 13745363 es EMPRESA DE MENOR TAMAÃ'O PRO-PYME: NO
init <- i

for(i in init:Drut.dim[1]){
  rut <- Drut[i,1]
  dv  <- Drut[i,2]
  url_01 <- "https://zeus.sii.cl/cvc_cgi/stc/getstc?RUT="
  url_02 <- "&txt_captcha=bUc1Rm5JaHpZYW%20syMDE0MTAxNjE1MzMyMjlBcERZY0hpd2h3MjQyNFZ5b1ZrSktn%20VDhjMDBoSWlsdHhrZ1FqLlFVSk5PR1ZPY1ZGWVl5NUlXUT09em%20RNOVdXWmNVY1E%3D&txt_code=2424&PRG=STC&OPC=NOR"
  url_F <- paste(url_01,rut,"&DV=",dv,url_02, sep="")
  docXml <- getURL(url_F)
  XmlDocTree <- htmlTreeParse(docXml)
  class(XmlDocTree)
  #htmlDocTree$children
  XmlDocTop <- xmlRoot(XmlDocTree)
  # print(XmlDocTop)
  
  # En caso de no exitir la información el XML se corta
  # Salimos sino la siguiente rutina se caera

  # xmlValue(XmlDocTop[[2]][[1]][[3]][1]$strong)
  
  options(show.error.messages = FALSE)
  err01 <- try(XmlDocTop[[2]][[1]][[3]][1])
  options(show.error.messages = TRUE)
  if(substring(err01[1],1,5) == "Error") next
  
  # Existe en la base de SII
  # Si existe, no habra table y dara NULL
  # Si no existe la tabla rendrá el texto "No ha sido posible completar su solicitud"
  if(!is.null(XmlDocTop[[2]][[1]][[3]][1]$table)){
    valRut <- xmlSApply(XmlDocTop[[2]][[1]][[3]][1]$table, function(x) xmlSApply(x, xmlValue))
    valRut <- substring(as.character(valRut)[1],1,41)
  } else valRut <- "Ok"
  if(valRut == "No ha sido posible completar su solicitud") next
  
  # Nombre
  nombre <- xmlValue(XmlDocTop[[2]][[1]][[4]][1]$text)
  if(nombre == "**") {nombre = "-"} 
  # Rut
  # XmlDocTop[[2]][[1]][[7]][1]
  
  # Presenta Inicio de Actividades
  Ini_Act_Factor <- substring(xmlValue(XmlDocTop[[2]][[1]][[12]][1]$text),47,48)
  
  if(Ini_Act_Factor == "SI"){
    Tabla01Xml <- XmlDocTop[[2]][[1]][[25]]
    Tabla01val <- xmlSApply(Tabla01Xml, function(x) xmlSApply(x, xmlValue))
    Tabla01Fin<- data.frame(t(Tabla01val),row.names=NULL)
    
    Tabla01Fin<-Tabla01Fin[-1,]
    #names(Tabla01Fin) <- c("Actividades","Código","Categoría","Afecta IVA")
    
  } else Tabla01Fin <- matrix(c("No Tiene","","",""),ncol=4)
  
  if(dim(Tabla01Fin)[2]==0) {Tabla01Fin <- matrix(c("No Tiene","","",""),ncol=4) }
  
  Ttemp <-  cbind(rep(rut,dim(Tabla01Fin)[1])
                 ,rep(dv,dim(Tabla01Fin)[1])
                 ,rep(nombre,dim(Tabla01Fin)[1])
                 ,Tabla01Fin)
  Ttemp <- as.data.frame(Ttemp)
  names(Ttemp) <- c("Rut","Dv","Nombre","Actividades","Código","Categoría","Afecta IVA")
  
  options(show.error.messages = FALSE)
  err02 <- try(dim(TFinal))
  options(show.error.messages = TRUE)
  if(substring(err02[1],1,5) == "Error"){ # si no existe la crea sino adjunta
    TFinal <- Ttemp
  } else TFinal <- rbind(TFinal,Ttemp)
}

View(TFinal)
tail(TFinal)


# DResultado <- TFinal

# Escribe en un txt como respaldo
#write.table(TFinal,"E:/CARAYA/07 - CODIGOS/03 - R scripts/rutero_para_sii_salida.txt"
#            , quote = FALSE
#            , sep=";")

# Escribimos en una tabla de servidor.

sqlQuery(chn, "drop table CA_AUX_20170529_SII")
sqlSave(chn,TFinal, tablename = "CA_AUX_20170529_SII")

sqlQuery(chn,"
--insert into [RT_ANALYTICS_ORI].[dbo].[CA_SII_Consolidado]
select cast(GETDATE() as date) as fecha
, cast(rut as int) as Rut, cast(Dv as nvarchar(1)) as Dv, Nombre, Actividades, cast(Código as int) as Codigo
, cast(Categoría as nvarchar(50)) as Categoria, cast(AfectaIVA as nvarchar(5)) as AfectaIva
from [RT_ANALYTICS_ORI].dbo.CA_AUX_20170529_SII
where cast(Rut as int) not in (select Rut from [RT_ANALYTICS_ORI].[dbo].[CA_SII_Consolidado])")

close(chn)

