# For this version I'll use the published cv. Making a mark and extracting the tables I'll
# working with the same function used in cuenta_pintados_00.

# To kick off I need the routine to run the Excel macro. Depending on the case we could use 
# the local outputs or the published data. This is the variable "ruta_cv".

# To properly make the process its mandatory have locally downloaded the files

#-------------------------------------------------------------------------------
# Into an overall way I put my key word on each tabulation for each section

library(dplyr)
library(RDCOMClient)
#---- Functions ----------------------------------------------------------------
cuenta_pintados <- function(cv){
  #cv <- tabulado[[2]]
  
  colunas <- purrr::map(.x = 1:ncol(cv), .f =~ na.omit(cv[,.x]))
  xx <- 0
  
  for (i in 2:length(colunas)){
    if (all(!colunas[[i]][1:length(colunas[[i]])]%in%NA) ) {
      xx <- c(xx, colunas[[i]])
    }else{
      
    }
  }
  
  # In the ENVE's case the accuracy level is given by: 
  # high (0<x<20), medium (20<= x <30) and low (30<= x) 
  # So I´ll sort the xx vector and use this thresholds to make the count.
  xx <- as.data.frame(sort(as.numeric(as.character(xx)))[-1])
  tamanio <- nrow(xx)
  resumen <- xx %>% mutate(nivel = case_when(xx[,1]<20 ~ "Alto",
                                             (20 <=xx[,1]) & (xx[,1] < 30) ~ "Medio",
                                             30 <= xx[,1] ~ "Bajo")
  ) %>%group_by(nivel) %>% count() %>% 
    mutate(porcentaje = round( (100 * n)/tamanio, 1))
  
  return(resumen)
  
}
llenado <- function(hoja, resumen){
  
  salida <- as.data.frame(hoja)
  salida <- cbind(salida, NA, NA, NA)
  
  if(nrow(resumen)==1){
    if(resumen[,1]=="Alto"){
      salida[1,2] <- paste0(resumen[,3], "%")
      salida[1,3] <- paste0(0.0, "%")
      salida[1,4] <- paste0(0.0, "%")
    }else if(resumen[,1]=="Bajo"){
      salida[1,2] <- paste0(0.0, "%")
      salida[1,3] <- paste0(resumen[,3], "%")
      salida[1,4] <- paste0(0.0, "%")
    }else if(resumen[,1]=="Medio"){
      salida[1,2] <- paste0(0.0, "%")
      salida[1,3] <- paste0(0.0, "%")
      salida[1,4] <- paste0(resumen[,3], "%")
    }else{
      print("Warning!!!!")
    }
  }else if(nrow(resumen)==2){
    if(resumen[1,1]=="Alto" & resumen[2,1]=="Bajo"){
      salida[1,2] <- paste0(resumen[1,3], "%")
      salida[1,3] <- paste0(0.0, "%")
      salida[1,4] <- paste0(resumen[2,3], "%")
    }else if(resumen[1,1]=="Alto" & resumen[2,1]=="Medio"){
      salida[1,2] <- paste0(resumen[1,3], "%")
      salida[1,3] <- paste0(resumen[2,3], "%")
      salida[1,4] <- paste0(0.0, "%")
    }else if(resumen[1,1]=="Bajo" & resumen[2,1]=="Medio"){
      salida[1,2] <- paste0(0.0, "%")  
      salida[1,3] <- paste0(resumen[1,3], "%")
      salida[1,4] <- paste0(resumen[2,3], "%")
    }else{
      "Warnig!!!"
    }
  }else{
    salida[1,2] <- paste0(resumen[1,3], "%")  
    salida[1,3] <- paste0(resumen[3,3], "%")
    salida[1,4] <- paste0(resumen[2,3], "%") 
  }
  
  return(salida)
}
tab_porcentajes <- function(hojas, lista){
  
  salida <- purrr::map(.x = 2:length(hojas), .f =~ llenado(hoja = hojas[.x], resumen = lista[[.x-1]]))
  salida <- do.call(rbind, salida)
  return(salida)
}
#---- placing the key word ------------------------------------------------------
# This process is carry out for each section. Thus, I think that is a good idea
# to have only one list with all the paths in order to get a shorter routine

con_cvs <- which(stringr::str_detect(list.files("D:/ENVE/ENVE 2024/Codigos_2022/cuenta_pintados/"), "cv"))
files_de_cvs <- list.files("D:/ENVE/ENVE 2024/Codigos_2022/cuenta_pintados/")[con_cvs]
ruta_cv <- paste0("D:/ENVE/ENVE 2024/Codigos_2022/cuenta_pintados/", files_de_cvs)

# Path of the excel macros.
source("D:/ENVE/ENVE 2024/Codigos_2022/cuenta_pintados/macro_tanos.R")
ruta_m <- "D:/Varios/correr_macro_R/correr_macro_R/macros_R/tanos.xlsm"

#ruta_cv <- "D:/ENVE/ENVE 2024/Codigos_2022/cuenta_pintados/enve2022_cv_I_nivel_victimizacion_delincuencia.xlsx"
for (i in 1:length(ruta_cv)) {
macro_tanos(ruta_m = ruta_m, lib_cv = ruta_cv[i] )
}

#---- Making the count ---------------------------------------------------------
# First I load the output workbook

calidad <- openxlsx::loadWorkbook("D:/ENVE/ENVE 2024/Codigos_2022/cuenta_pintados/tab_cuenta_pintados.xlsx")

# Now I load the workbook with the new mark and run the previous functions. 
# I already have the corresponding paths of the files
for(i in 1:length(ruta_cv)){
mi_wb <- openxlsx::loadWorkbook(ruta_cv[i])

hojas <- openxlsx::sheets(mi_wb)
tabulados <- purrr::map(.x = 2:length(hojas), .f =~ openxlsx::read.xlsx(mi_wb, .x))

tabulados <- lapply( 1:length(tabulados), function(i){
                   inicio <- which(tabulados[[i]][,1]%in%"Estados Unidos Mexicanos")
                   fin <- which(tabulados[[i]][,1]%in%"Tanos")-1
                   recortado <- tabulados[[i]][inicio:fin,]
 
                   return(recortado)
})
resumen <- purrr::map(.x = 1:length(tabulados), .f =~ cuenta_pintados(cv = tabulados[[.x]]))

openxlsx::writeData(calidad, x=tab_porcentajes(hojas=hojas, lista = resumen), sheet = paste0("Secc", i), startRow = 10, startCol = 1,colNames = FALSE)
  
}


# Finally I save the workbook with the new tables.
openxlsx::saveWorkbook(calidad, "D:/ENVE/ENVE 2024/Codigos_2022/cuenta_pintados/tab_cuenta_pintados.xlsx", overwrite = T)
