# Seguimiento de Empresas
#
# Objetivo: Base para construir el seguimiento de coyuntura, adem�s,  
#           para env�ar informaci�n al Banco Central sobre gesti�n.
#
# Equipo IVS
#
# Versi�n 22-06-2021
#
##################### Previo #######################################
getwd()
setwd("C:/Users/1andr/OneDrive/INE/Gesti�n BCCh/R")

rm(list = ls()) 
#library(tidyverse)
if (!require("readxl")) install.packages("readxl"); library(readxl)
if (!require("dplyr")) install.packages("dplyr"); library(dplyr)
if (!require("openxlsx")) install.packages("openxlsx"); library(openxlsx)
if (!require("zoo")) install.packages("zoo"); library(zoo)
if (!require("ggplot2")) install.packages("ggplot2"); library(ggplot2)
if (!require("scales")) install.packages("scales"); library(scales)
if (!require("reshape2")) install.packages("reshape2"); library(reshape2)

options(scipen=999)

##################### Actualizando Ponderador  #################################
# Ponderador Base 2020
# ponderador = suma_ventas_i_2020/suma_ventas_2020
# con i = empresa1, empresa2, ...., empresa1057

# Cargando base de calculo
# Ordenar por Categor�a, Rol, A�o, Mes y creando variable periodo
calculo <- read_excel("05_Base Calculo.xlsx") %>%
  arrange(Categoria, RolGenerico, A�o, Mes)   %>%
  mutate(Periodo = as.Date(paste0(1,"/",Mes,"/",A�o), '%d/%m/%Y')
         ,Margen = Ventas - CostoVenta
         ,MargenExp = Margen * FactorExpansion)

# Quitando las RRAA
RRAA    <- read_excel("06_Objetos.xlsx", sheet = "RRAA")
calculo <- left_join(calculo, RRAA, by = c("RolGenerico")) %>% 
  filter(is.na(RRAA))

# Cargando Directorio para Seleccionar solo "Ambito de Estudio"
directorio <- read_excel("01_Directorio de Seguimiento.xlsx", range = "A3:AV1060") %>%
  rename("RolGenerico" = "Rol Generico")

calculo    <- left_join(calculo, directorio, by = c("RolGenerico")) %>% 
  filter(`Estado Verificacion` == '�mbito de estudio') %>%
  rename("A�o" = "A�o.x")

# �Cu�ntas empresas hay?
calculo %>% summarise(n_distinct(RolGenerico))

# Margen Expandido de empresas por a�o
sum_ventas_i <- calculo %>% group_by(A�o, RolGenerico, `Razon Social Unidad Estadistica`) %>% 
  summarise(sum_margen_i = sum(MargenExp))

sum_ventas <- calculo %>% group_by(A�o) %>%
  summarise(sum_margen = sum(MargenExp))

ponderador <- inner_join(sum_ventas_i, sum_ventas, by = "A�o") %>%
  mutate(Ponderador = sum_margen_i/sum_margen)

rm(directorio, RRAA, sum_ventas, sum_ventas_i)

# Ponderador para a�o 2020
ponderador <- ponderador %>% 
  filter(A�o == 2020) %>% 
  select(RolGenerico, `Razon Social Unidad Estadistica`,Ponderador)

#################### Comparando con Ponderador antiguo #########################
# Comparando ponderadores
ponderador_new <- ponderador
ponderador_old <- read_excel("06_Objetos.xlsx", sheet = "Ponderador_old") %>% 
  rename("Ponderador_old" = "Ponderador")

# Juntando ponderador antiguo y nuevo
comp_pondera  <- left_join(ponderador_new, ponderador_old, by = c("RolGenerico"))

# 1� Contraste
# Contrastar las empresas mas grandes seg�n el ponderador anterior, 
# y ver como cambian sus ponderadores nuevos.
contrast1 <- comp_pondera %>% 
  arrange(desc(Ponderador_old)) %>%
  select(RolGenerico, `Razon Social Unidad Estadistica`, Ponderador_old) 
contrast1$Ranking_viejo <- seq(1,931,1)

# 2� Contraste
# Contrastar las empresas mas grandes seg�n el ponderador nuevo 
# y ver como cambian respecto al ponderador anterior.
contrast2 <- comp_pondera %>% 
  arrange(desc(Ponderador)) %>%
  select(RolGenerico, `Razon Social Unidad Estadistica`, Ponderador)
contrast2$Ranking_nuevo <- seq(1,931,1)

# 3� Contraste
# Ver los casos de mayor diferencia de ponderadores. 
contrast3 <- comp_pondera %>% 
  mutate(dif = Ponderador - Ponderador_old
         ,abs_dif = abs(dif)) %>% 
  arrange(desc(abs_dif)) %>%
  select(-c(abs_dif, Ponderador, Ponderador_old))
contrast3 <- contrast3[1:20,]
contrast3$Ranking_dif <- seq(1,20,1)

# Comparaci�n
contrast <- left_join(contrast2, contrast1, by = c("A�o","RolGenerico", "Razon Social Unidad Estadistica")) %>%
  mutate(dif_rank = Ranking_nuevo - Ranking_viejo
         ,dif_pond = Ponderador - Ponderador_old) %>%
  select(A�o, Ranking_nuevo, Ranking_viejo, dif_rank, RolGenerico, `Razon Social Unidad Estadistica`
         ,Ponderador, Ponderador_old, dif_pond)
contrast <- contrast[1:20,]


# Guardando informaci�n
wb<-createWorkbook()
addWorksheet(wb, sheet = "Ponderador")
addWorksheet(wb, sheet = "Contraste")
addWorksheet(wb, sheet = "Contraste3")
writeData(wb, "Ponderador", ponderador)
writeData(wb, "Contraste", contrast)
writeData(wb, "Contraste3", contrast3)
saveWorkbook(wb, "Ponderador 2020.xlsx",overwrite = TRUE)

rm(comp_pondera, contrast1, contrast2, contrast3
   , ponderador_new, ponderador_old, wb)

##################### Directorio y Seguimiento #################################
rm(list = ls()) 
  
## Bases de datos
directorio <- read_excel("01_Directorio de Seguimiento.xlsx", range = "A3:AV1060") %>%
  rename("RolGenerico" = "Rol Generico"
         ,"EstadoVerificacion" = "Estado Verificacion")
RRAA       <- read_excel("06_Objetos.xlsx", sheet = "RRAA")

# Ponderador Nuevo con base 2020
ponderador <- read_excel("Ponderador 2020.xlsx", sheet = "Ponderador") %>%
  select(RolGenerico ,Ponderador)

## Juntando Directorio con RRAA y con ponderador x RolGenerico
directorio <- left_join(directorio, RRAA, by = "RolGenerico")
directorio <- left_join(directorio, ponderador, by = "RolGenerico")

## Base de Seguimiento 
#  RRAA: En blanco
#  EstadoVerificacion: �mbito de Estudio
seguimiento <- directorio %>% 
  filter(EstadoVerificacion == '�mbito de estudio'
         , is.na(RRAA))

## Conteo de Rol y Suma de Ponderador por Estado Gestion
seguimiento %>% group_by(`Estado Gestion`) %>% 
  summarise(CuentadeRol     = n_distinct(RolGenerico)
            ,SumaPonderador = sum(Ponderador))

rm(directorio)

##################### Empresas Destacadas #################################
# Base de datos Total, Limpiando para luego juntar con destacadas
total      <- read_excel("02_Total_Dduran.xls", range = "A5:X1500") %>%
  filter(!is.na(A�o)) %>%
  select(Rol, 'Rol Gen�rico', 'Estado de Recepci�n'
         ,'Fecha de Recepci�n') %>%
  rename('RolGenerico' = 'Rol Gen�rico')
  
# Empresas Destacadas  
destacada  <- read_excel("06_Objetos.xlsx", sheet = "ED")

# Juntando base de empresas destacadas con total
# Esto nos da el estado de recepci�n en la destacada
destacada    <- destacada %>% 
  inner_join(total, by = c("Rol", "RolGenerico"))

## Resumen de Empresas destacadas
TD_destacada <- destacada %>% 
  group_by(`Estado de Recepci�n`) %>%
  summarise('Cantidad de Empresas' = n_distinct(RolGenerico))

## Resumen de Empresas destacadas por Secci�n
TD_destacada_seccion <- destacada %>% 
  group_by(`Estado de Recepci�n`, Seccion) %>%
  summarise('Cantidad de Empresas' = n_distinct(RolGenerico))

## Resumen de Empresas destacadas por Empresa
TD_destacada_empresa <- destacada %>% 
  group_by(`Estado de Recepci�n`, Seccion, `Razon Social Unidad Estadistica`) %>%
  summarise('Cantidad de Empresas' = n_distinct(RolGenerico))

## Empresas Importantes IVS (EI)
EI_IVS <- TD_destacada %>% 
  mutate(PorcentajeEmpresas = 100*`Cantidad de Empresas`/187)

# Limpiando Global Environment
rm(destacada,TD_destacada, TD_destacada_empresa, TD_destacada_seccion, total)

##################### Tasa de Recepci�n Actividad ##############################
## Seguimiento por Actividad
seg_actividad <- seguimiento %>% 
  group_by(Seccion, `Apertura IVS`, `Estado Gestion`) %>%
  summarise('Estado de Recepci�n' = n_distinct(RolGenerico))

# Reshape
seg_actividad <- dcast(seg_actividad, `Apertura IVS` ~ `Estado Gestion`)

seg_actividad$Completa[is.na(seg_actividad$Completa)]             <- 0
seg_actividad$`En An�lisis`[is.na(seg_actividad$`En An�lisis`)]   <- 0
seg_actividad$Enviada[is.na(seg_actividad$Enviada)]               <- 0
seg_actividad$Incompleta[is.na(seg_actividad$Incompleta)]         <- 0
seg_actividad$`No Ingresada`[is.na(seg_actividad$`No Ingresada`)] <- 0
seg_actividad$Terminada[is.na(seg_actividad$Terminada)]           <- 0

# Suma Total General
seg_actividad <- seg_actividad %>% 
  mutate(TotalGeneral          = `En An�lisis` + Enviada   + `No Ingresada` 
         + Terminada + Incompleta #+ Completa
         ,'Enviada o Analisis' =  Enviada + `En An�lisis`
         ,NoIngresada          = `No Ingresada` + Incompleta ) %>% #+ Completa
  arrange(`Apertura IVS`)

#### REVISAR
## Tasa Recepci�n Actividad
TR_ACT <- seg_actividad %>% 
  select(`Apertura IVS`, Terminada, 'Enviada o Analisis', NoIngresada) %>%
  rename('Actividad' = 'Apertura IVS'
         ,'Terminadas' = 'Terminada'
         )

TR_ACT <- TR_ACT %>% 
  mutate(TerminadasPor  = 100*Terminadas/(Terminadas + `Enviada o Analisis` + NoIngresada)
         ,EnviadasPor   = 100*`Enviada o Analisis`/(Terminadas + `Enviada o Analisis` + NoIngresada)
         ,NoInresadaPor = 100*NoIngresada/(Terminadas + `Enviada o Analisis` + NoIngresada))

# Para ordenar llamamos a objeto 2
objeto_2 <-read_excel("06_Objetos.xlsx", sheet = "SA")

# Juntando bases 
TR_ACT <- inner_join(objeto_2, TR_ACT, by = "Actividad")

## ORDENAR POR ACTIVIDAD ##
rm(seg_actividad, objeto_2)

##################### IVS MAYO #################################################

## Bases de Datos
estados_roles <- read_excel("03_EstadosRolesLevantados_123_Dduran.xls", range = "B6:R1500") %>%
  filter(!is.na(A�o)) %>%
  rename('RolGenerico' = 'Rol Generico') 
  
objeto_1      <- seguimiento %>% select(RolGenerico, Ponderador)

# Estado Roles con el Seguimiento
ER <- inner_join(estados_roles, objeto_1, by = "RolGenerico")

# Fecha de Recepci�n
FR <- ER %>% 
  group_by(`Fecha de Recepci�n`) %>% 
  summarise('CuentaRolR' = n_distinct(RolGenerico)) %>%
  filter(!is.na(`Fecha de Recepci�n`)) %>%
  rename('Dias de Levantamiento' = 'Fecha de Recepci�n')

# Fecha de Termino
FT <- ER %>% 
  group_by(`Fecha de Termino`) %>% 
  summarise('CuentaRolT' = n_distinct(RolGenerico)
            ,'SumPonderador' = sum(Ponderador)) %>%
  filter(!is.na(`Fecha de Termino`)) %>%
  rename('Dias de Levantamiento' = 'Fecha de Termino')

# Juntando Bases por Fechas
IVS_MAY <- left_join(FR, FT, by = 'Dias de Levantamiento')

IVS_MAY$CuentaRolR[is.na(IVS_MAY$CuentaRolR)] <- 0
IVS_MAY$CuentaRolT[is.na(IVS_MAY$CuentaRolT)] <- 0
IVS_MAY$SumPonderador[is.na(IVS_MAY$SumPonderador)] <- 0

# Suma Acumulada de Cuenta Rol de Recepci�n, de Termino y Suma Ponderador
IVS_MAY <- IVS_MAY %>% within(Enviadas   <- Reduce("+", CuentaRolR, accumulate = TRUE))
IVS_MAY <- IVS_MAY %>% within(Terminadas <- Reduce("+", CuentaRolT,accumulate = TRUE))
IVS_MAY <- IVS_MAY %>% within(SumPondera <- Reduce("+", SumPonderador,accumulate = TRUE))

# Data de Env�ada o An�lisis, No Recepcionadas y Ponderador en 100%
IVS_MAY <- IVS_MAY %>% 
  mutate(Enviada_Analisis = Enviadas - Terminadas
         ,No_Recepcionada = 931 - Enviadas
         ,Term_Ponderadas = SumPondera
         ,Tot_Muestra     = Terminadas + Enviada_Analisis + No_Recepcionada) %>%
  select(`Dias de Levantamiento`, Tot_Muestra, Terminadas
         , Enviada_Analisis, No_Recepcionada, Term_Ponderadas) %>%
  mutate(Terminadas_Porc = 100*Terminadas/Tot_Muestra
         ,Term_Pond_Porc = Term_Ponderadas
         ,Enviada_A_Porc = 100*Enviada_Analisis/Tot_Muestra
         ,No_Recepc_Porc = 100*No_Recepcionada/Tot_Muestra)


# Limpiando Base de datos
rm(ER, estados_roles, FR, FT, objeto_1)


##################### Salidas ##################################################

## Salida de Seguimiento
# Exportando Salidas a Excel
# Gesti�n de Carpeta
setwd("C:/Users/1andr/OneDrive/INE/Gesti�n BCCh/R")

# Archivo de salida
wb<-createWorkbook()

# Creando Hojas para Excel
addWorksheet(wb, sheet = "IVS May")
addWorksheet(wb, sheet = "Empresas importantes IVS")
addWorksheet(wb, sheet = "Tasa Recepci�n Actividad")

# Cargando data en Hojas
writeData(wb, "IVS May", IVS_MAY)
writeData(wb, "Empresas importantes IVS", EI_IVS)
writeData(wb, "Tasa Recepci�n Actividad", TR_ACT)

# Archivo Final
saveWorkbook(wb, "Seguimiento_Coyuntura.xlsx",overwrite = TRUE)

# Salida de ponderador
wb<-createWorkbook()
addWorksheet(wb, sheet = "Ponderador_2020")
writeData(wb, "Ponderador_2020", ponderador)
saveWorkbook(wb, "Ponderador_2020.xlsx",overwrite = TRUE)








