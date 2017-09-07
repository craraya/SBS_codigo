#### This script uses RCurl and RJSONIO to download data from Google's API:
#### Latitude, longitude, location type (see explanation at the end), formatted address
#### Notice ther is a limit of 2,500 calls per day

library(RCurl)
library(RJSONIO)
library(leaflet)

url <- function(address, return.call = "json", sensor = "false") {
  root <- "http://maps.google.com/maps/api/geocode/"
  u <- paste(root, return.call, "?address=", address, "&sensor=", sensor, sep = "")
  return(URLencode(u))
}

### Documentacion: https://developers.google.com/maps/documentation/distance-matrix/intro?hl=es-419
url_d <- function(origen_lat, origen_lng, return.call = "json"){
  origen_lat=33
  origen_lng =-70
  dest_lat=32
  dest_lng=-71
  root <- "http://maps.google.com/maps/api/distancematrix/"
  u <- paste(root, return.call, "?origins=" , origen_lat,"," , origen_lng
          , "&destinations=", dest_lat,",", dest_lng
          #, "&key"
          , "&sensor=false"
          , sep = "")
  return(URLencode(u))
}

getURL(URLencode(u)) %>%
fromJSON(simplify=FALSE)  


getURL(URLencode("https://maps.googleapis.com/maps/api/geocode/json?address=1600+Amphitheatre+Parkway,+Mountain+View,+CA&sensor=false")) %>%
fromJSON(simplify=FALSE)  

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

direccion<-"Duble Almeyda 3272, Santiago, Chile"
url(direccion)
d1_lat<-as.numeric(geoCode(direccion)[1])
d1_lng<-as.numeric(geoCode(direccion)[2])
d1_lat;d1_lng

my_map <- leaflet() %>%
  addTiles() %>%
  addMarkers(lat=d1_lat, lng=d1_lng,
             popup="Mi Casa")
my_map
  


