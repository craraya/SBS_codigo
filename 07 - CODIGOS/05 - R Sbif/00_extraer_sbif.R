rm(list = ls())
library("dplyr")
library("RODBC")
library("stringr")
library("lubridate")

periodos <- seq(ymd("20070101"), # ESTO NO SE MODIFICA! NADA SE MODIFICA!!. ATTE JK
                ymd(format(Sys.Date(), "%Y%m01")),
                by = "1 month") 
periodos <- format(periodos, "%Y%m")

chn <- odbcConnect("maca")

for (per in periodos) {
  # per <- sample(periodos, size = 1)
  message(per)
  
  archivo <- paste("data/sbif_", per, ".rds", sep="")
  
  Startprocess <- Sys.time()
  
  if(!file.exists(archivo)) {
    
    query <- paste('select CONVERT(varchar(6),fecha,112) as periodo,[rut]
                 ,[m_deudadirectavigente]
                 ,[m_deudadirectamorosa]
                 ,[m_deudadirectavencida]
                 ,[m_deudadirectaifinanciera]
                 ,[m_deudadirectaoperacionespactadas]
                 ,[m_deudaindirectavigente]
                 ,[m_deudaindirectavencida]
                 ,[m_deudacomercial]
                 ,[m_deudacreditoconsumo]
                 ,[n_institucionescondeuda]
                 ,[m_creditohipotecario]
                 ,[m_castigosdirectos]
                 ,[m_castigosindirectos]
                 ,[m_cupolineacredito]
                 ,[m_deudacomercialvigente]
                 ,[m_deudacomercialvencida]
                 ,[m_deudacreditocomerciales]
                 ,[m_deudaleasing]
                 ,[m_deudamorosaleasing]
                 from ods.dbo.fact_librodeudores
                 where CONVERT(varchar(6),fecha,112) = ',per, sep = "")
    
    sbif <- sqlQuery(chn,query)
    
    if(nrow(sbif) > 0) {
      saveRDS(sbif, file = archivo)  
    }
    
    rm(sbif)
    gc()
    
  }
  
  Exectime <- Sys.time() - Startprocess
  message(per, " ejectuado en:")
  print(Exectime)
}

# rm(list = ls())




