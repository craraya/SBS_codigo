# rm(list = ls())
# load("data2/df.RData")
# ruts <- df$rut
# pers <- df$periodo

download_factlibrodeudores <- function(ruts = sample(1000000:50000000, size = 100),
                                       pers = sample(c(201003, 201012), size = 100, replace = TRUE),
                                       past = 38,
                                       futr = 12,
                                       variables = c("periodo", "rut", "m_deudadirectavigente", "m_deudadirectamorosa", 
                                                     "m_deudadirectavencida", "m_deudadirectaifinanciera", "m_deudadirectaoperacionespactadas", 
                                                     "m_deudaindirectavigente", "m_deudaindirectavencida", "m_deudacomercial", 
                                                     "m_deudacreditoconsumo", "n_institucionescondeuda", "m_creditohipotecario", 
                                                     "m_castigosdirectos", "m_castigosindirectos", "m_cupolineacredito", 
                                                     "m_deudacomercialvigente", "m_deudacomercialvencida", "m_deudacreditocomerciales", 
                                                     "m_deudaleasing", "m_deudamorosaleasing")){
  
  library("dplyr")
  library("readr")
  library("tidyr")
  library("purrr")
  library("lubridate")
  
  ori_vars <- c("periodo", "rut", "m_deudadirectavigente", "m_deudadirectamorosa", 
            "m_deudadirectavencida", "m_deudadirectaifinanciera", "m_deudadirectaoperacionespactadas", 
            "m_deudaindirectavigente", "m_deudaindirectavencida", "m_deudacomercial", 
            "m_deudacreditoconsumo", "n_institucionescondeuda", "m_creditohipotecario", 
            "m_castigosdirectos", "m_castigosindirectos", "m_cupolineacredito", 
            "m_deudacomercialvigente", "m_deudacomercialvencida", "m_deudacreditocomerciales", 
            "m_deudaleasing", "m_deudamorosaleasing")
  
  # no existen ciertas columnas
  stopifnot(! length(setdiff(tolower(variables), ori_vars)) > 0)
  
  
  dfbig <- data_frame(rut = ruts, per = pers) %>% 
    group_by(rut, per) %>% 
    do(diff = seq(-past, futr)) %>% 
    unnest() %>% 
    mutate(fecha = paste0(per, "01"),
           fecha = ymd(fecha),
           fecha2 = fecha + months(diff),
           per2 = format(fecha2, "%Y%m")) %>% 
    distinct(rut, per2)
  
  perstoload <- dfbig$per2 %>% 
    unique() %>% 
    sort()
  
  sprintf("Proceso descargará %s periodos desde %s hasta %s (%s registros app)",
          length(perstoload), min(perstoload), max(perstoload), nrow(dfbig)) %>%
    message()
  
  fname <- sprintf("outputs/sbif_%s_%s_%s.txt",
                   Sys.info()[["user"]], Sys.Date(), length(ruts))
  
  map(perstoload, function(p){ # p <- sample(perstoload, size = 1)
    
    prcnt <- scales::percent(which(p == perstoload)/length(perstoload))
    message("Descargando periodo ", p, " - ", prcnt)
    
    pfile <-  paste("Z:/proceso - fact_librodeudores/data/sbif_", p, ".rds", sep = "")
    
    if(file.exists(pfile)) {
      
      ruts_to_filter  <- dfbig %>% filter(per2 == p) %>% .$rut
      dfaux <- readRDS(pfile)
      dfaux <- tbl_df(dfaux)
      dfaux <- filter(dfaux, rut %in% ruts_to_filter)
      dfaux <- select_(dfaux, .dots = variables)
      
      write_tsv(format(dfaux,digits=NULL), path = fname, append = file.exists(fname))
      
    }
    
  })

  message("Terminado Muahahaha :D")
    
  invisible()
  
}
