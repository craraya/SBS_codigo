#### This script uses RCurl and RJSONIO to download data from Google's API:
#### Latitude, longitude, location type (see explanation at the end), formatted address
#### Notice ther is a limit of 2,500 calls per day

library(RCurl)
library(RJSONIO)
library(plyr)

url <- function(address, return.call = "json", sensor = "false") {
  root <- "http://maps.google.com/maps/api/geocode/"
  u <- paste(root, return.call, "?address=", address, "&sensor=", sensor, sep = "")
  return(URLencode(u))
}

geoCode <- function(address,verbose=FALSE) {
  if(verbose) cat(address,"\n")
  u <- url(address)
  doc <- getURL(u)
  x <- fromJSON(doc,simplify = FALSE)
  if(x$status=="OK") {
    lat <- x$results[[1]]$geometry$location$lat
    lng <- x$results[[1]]$geometry$location$lng
    location_type <- x$results[[1]]$geometry$location_type
    formatted_address <- x$results[[1]]$formatted_address
    return(c(lat, lng, location_type, formatted_address))
  } else {
    return(c(NA,NA,NA, NA))
  }
}

dist_euc<-function(x1,y1,x2,y2){
  r<-sqrt((x1-x2)^2 + (y1-y2)^2)
  return(r)
}

centro_emp<-read.delim("U:/01 - ANALISIS VARIOS/2014-11-17 - Georeferenciacion/Direcciones.txt",encoding="Latin-1")
oficinas<-read.delim("U:/01 - ANALISIS VARIOS/2014-11-17 - Georeferenciacion/Oficinas.txt",encoding="Latin-1")
attach(oficinas)
N<-nrow(oficinas)
names(oficinas)

lat<-rep(0,N)
lng<-rep(0,N)
for(i in 1:N){
  lat[i]<-as.numeric(geoCode(Direccion_2.1[i])[1])
  lng[i]<-as.numeric(geoCode(Direccion_2.1[i])[2])
}
centro_2<-cbind(oficinas,as.numeric(lat),as.numeric(lng))
#ncol(centro_2)
#colnames(centro_2)
#geoCode(centro_emp[1,6])[2]
direccion<-"centenario de san miguel 1056, san miguel, Santiago, Chile"
#direccion<-Dirección.2[63]
d1_lat<-as.numeric(geoCode(direccion)[1])
d1_lng<-as.numeric(geoCode(direccion)[2])
d1_lat;d1_lng
Dirección.2[63]

dist<-rep(999,N)
for(i in 1:N){
  if(!is.na(centro_2[i,13])) 
  {dist[i]<-dist_euc(d1_lat,d1_lng,centro_2[i,13],centro_2[i,14])}
}
dist
minimo<-min(dist)
for(i in 1:N){
  if(dist[i]==minimo) {
    centro<-centro_2[i,2]
    ID<-centro_2[i,1]
  }
}
dist[ID]
centro
ID

plot(centro_2[,13],centro_2[,12])
#lines(d1_lng,d1_lat,type="p",col="Red")
points(d1_lng,d1_lat,type="p",col="Red")
points(centro_2[ID,8],centro_2[ID,7],type="p",col="Blue")


