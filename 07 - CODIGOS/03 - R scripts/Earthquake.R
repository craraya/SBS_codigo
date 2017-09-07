

library(rvest)      # no se si se ocupa, se hace todo con XML
library(RCurl)      # para el formato html
library(XML)
library(xml2)       # no se si se ocupa
library(leaflet)    # para los mapas
library(lubridate)  # para las fechas

#u <- "https://earthquake.usgs.gov/fdsnws/event/1/query?format=xml&starttime=2017-04-28&endtime=2017-04-28&minmagnitude=5"

# last year
u <- "https://earthquake.usgs.gov/fdsnws/event/1/query?format=xml&starttime=2017-04-27&endtime=2017-04-29&minmagnitude=4"
xmlfile <- URLencode(u) %>%
getURL() %>%
read_xml() %>%
xmlParse()

class(xmlfile)

## Gives content of root
xmltop = xmlRoot(xmlfile) 

nam <- names(xmltop[['eventParameters']])
n <- length(nam[nam=='event'])

#class(xmltop)             #"XMLInternalElementNode" "XMLInternalNode" "XMLAbstractNode"
#xmlName(xmltop)           #give name of node, PubmedArticleSet
#xmlSize(xmltop)           #how many children in node, 19
#xmlName(xmltop[[1]])      #name of root's children

#xmlSize(xmltop[[1]])
#xmlSApply(xmltop[[1]], xmlName)   #name(s)
#xmlSApply(xmltop[[1]], xmlAttrs)  #attribute(s)
#xmlSApply(xmltop[[1]], xmlSize)   #size

#xmlSApply(xmltop[[1]][[1]], xmlName)
names(xmltop[[1]][[1]])

##Descripción

nombre <- c()
localidad <- c()
fecha <- c()
lng <- c()
lat <- c()
depth <- c()
mag <- c()

for(i in 1:n){
  nombre[i]    <- xmlValue(xmltop[[1]][[i]][['description']][['type']][1]$text)
  localidad[i] <- xmlValue(xmltop[[1]][[i]][['description']][['text']][1]$text)
  fecha[i]     <- xmlValue(xmltop[[1]][[i]][['origin']][['time']][['value']][1]$text)
  lng[i]       <- xmlValue(xmltop[[1]][[i]][['origin']][['longitude']][['value']][1]$text)
  lat[i]       <- xmlValue(xmltop[[1]][[i]][['origin']][['latitude']][['value']][1]$text)
  depth[i]     <- xmlValue(xmltop[[1]][[i]][['origin']][['depth']][['value']][1]$text)
  mag[i]       <- xmlValue(xmltop[[1]][[i]][['magnitude']][['mag']][['value']][1]$text)

}

lng <- as.numeric(lng)
lat <- as.numeric(lat)
depth <- as.numeric(depth)
mag <- as.numeric(mag)

## Mapeamos

leaflet() %>%
  addMarkers(lng=lng,lat=lat, label = 'Earthquake', popup=paste(localidad,mag,sep=', ')) %>%
  #addMarkers(lng=end_loc()[1],lat=end_loc()[2], label = 'To', popup=end_address()) %>%
  #addPolylines(lng=ruta()$lng, lat=ruta()$lat, color='cyan',opacity = 1,group = "Line") %>%
  addProviderTiles("OpenStreetMap.Mapnik", group = "Road map") %>%
  addProviderTiles(providers$CartoDB.Positron, group = "Carto") %>%
  #addProviderTiles(providers$Stamen.Toner, group = "Toner") %>%
  #addProviderTiles("Esri.WorldImagery", group = "Satellite") %>%
  addLayersControl(position = 'topright',
                   #baseGroups = c("Topographical", "Road map", "Satellite"),
                   baseGroups = c("Road map","Carto"),
                   #overlayGroups = c("Hiking routes", "Photo markers"),
                   options = layersControlOptions(collapsed = FALSE))

earthTable <- as.data.frame(cbind(nombre
                    ,localidad
                    ,fecha
                    ,lng
                    ,lat
                    ,depth
                    ,mag))

