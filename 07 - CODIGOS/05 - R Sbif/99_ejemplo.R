rm(list = ls())
source("code/01_download_factlibrodeudores.R")
library(dplyr)

#load("inputs/GH_DIFERENCIAS_BORRAR.RData")


df <- tbl_df(datos_cat_6m)
head(df)

df <- df %>% 
  mutate(periodo=201601)

# download_factlibrodeudores(df$rut, df$periodo, past = 38, futr = 12)
download_factlibrodeudores(
  ruts = df$rut,
  pers = df$periodo,
  past = 38,
  futr = 2,
  variables = c("periodo", "rut","m_deudadirectavigente"
                ,"m_deudadirectamorosa"
                ,"m_deudadirectavencida"
                ,"m_deudadirectaifinanciera"
                ,"m_deudadirectaoperacionespactadas"
                ,"m_deudaindirectavigente"
                ,"m_deudaindirectavencida"
                ,"m_deudacomercial"
                ,"m_deudacreditoconsumo"
                ,"n_institucionescondeuda"
                ,"m_creditohipotecario"
                ,"m_castigosdirectos"
                ,"m_castigosindirectos"
                ,"m_cupolineacredito"
                ,"m_deudacomercialvigente"
                ,"m_deudacomercialvencida"
                ,"m_deudacreditocomerciales"
                ,"m_deudaleasing"
                ,"m_deudamorosaleasing"))



