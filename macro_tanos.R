#Se requiere la ruta de la macro que pinta que debe estar habilitada para ejecutarse así chido
# como pueden tenerse varios libros en un mismo directorio se pide por separado el directorio
# y el nombre de los cv entre comillas.

macro_tanos <- function(ruta_m, lib_cv){

#dir = "D:/Varios/correr_macro_R/correr_macro_R/libros/" O sea donde están tus libros
#setwd(dir)
#################################LEER ARCHIVOS EN EL DIRECTORIO
#a = list.files(dir)
#############################################################
xlApp <- COMCreate("Excel.Application")
#####################################
#Corresponde al libro de cv#
xlWbk <- xlApp$Workbooks()$Open(lib_cv)
xlApp[['Visible']] <- T
###################################
xlWbk <- xlApp$Workbooks()$Open(ruta_m)
xlApp[['Visible']] <- F
##CORRER LA MACRO################
xlApp$Run("tanos")
# Close the workbook and quit the app:
xlWbk$Close(FALSE)
xlApp$Quit()
#######################################
}
