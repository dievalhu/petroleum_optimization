setwd("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B") #directorio de trabajo
rm(list = ls()) #limpieza memoria
# LIBRERIAS
  library(readr)
  library(readxl)
  #install.packages("lpSolve")
  install.packages("xlsx")
  library(lpSolve)
  #install.packages("xtable")
  library(xtable)
eta<-etanol<-0.95
ds<-dias_seguridad_stock<-5

# DEMANDA POR DIA Y POR PRODUCTO
dm<-Demandas<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                        sheet = "Demandas", range = "b1:n9"))
Demandas_mes <- Demandas[,-which(colMeans(is.na(TG_UPS_2021_LMP_B)) >= 0.20)]#eliminar columnas vacias
Días_mes<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", sheet = "Demandas", range = "b10:m11"))
Días_mes <- Días_mes[,-which(colMeans(is.na(TG_UPS_2021_LMP_B)) >= 0.20)]#eliminar columnas vacias
Días<-sum(Días_mes);ld<-length(Días_mes)

# DEMANDA POR TERMINAL
  dm_ter<-as.matrix(dm_ter<- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                        sheet = "Part. Terminal", range = "B1:N9"))
  rownames(dm)<-rownames(dm_ter)<-as.matrix(dm1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                              sheet = "Part. Terminal",range = "A1:A9"))
  dm_ter;dm
  ### Cálculo de la demanda mensual por producto y por terminal
  i<-NULL; A<-NULL;B<-NULL
  for ( i in 1:13) {
    A<-dm_ter[,i]*dm
    B<-rbind(B,A)
  }
  i<-NULL;GS<-NULL;GE<-NULL;GEC<-NULL;C2T<-NULL;DP<-NULL;D2<-NULL;JF<-NULL;RC<-NULL
  for ( i in 1:13) {
    GS<-rbind(GS,B[i+(i-1)*7,])
    GE<-rbind(GE,B[i+(i-1)*7+1,])
    GEC<-rbind(GEC,B[i+(i-1)*7+2,])
    C2T<-rbind(C2T,B[i+(i-1)*7+3,])
    DP<-rbind(DP,B[i+(i-1)*7+4,])
    D2<-rbind(D2,B[i+(i-1)*7+5,])
    JF<-rbind(JF,B[i+(i-1)*7+6,])
    RC<-rbind(RC,B[i+(i-1)*7+7,])
  }
  nam<-rownames(GS)<-rownames(GE)<-rownames(GEC)<-rownames(C2T)<-rownames(DP)<-rownames(D2)<-rownames(JF)<-rownames(RC)<-as.matrix(A1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                                                                                                           sheet = "Part. Terminal",range = "B1:N1", col_names = FALSE))
  ### Cálculo de la demanda día por Refinerías y Terminales
  i<-NULL; GSd<-GEd<-GECd<-C2Td<-DPd<-D2d<-JFd<-RCd<-NULL; A<-B<-C<-D<-E<-F<-G<-H<-NULL
  for ( i in 1:12) {
    A<-GS[,i]/Días_mes[i];GSd<-cbind(GSd,A)
    B<-GE[,i]/Días_mes[i];GEd<-cbind(GEd,B)
    C<-GEC[,i]/Días_mes[i];GECd<-cbind(GECd,C)
    D<-C2T[,i]/Días_mes[i];C2Td<-cbind(C2Td,D)
    E<-DP[,i]/Días_mes[i];DPd<-cbind(DPd,E)
    F<-D2[,i]/Días_mes[i];D2d<-cbind(D2d,F)
    G<-JF[,i]/Días_mes[i];JFd<-cbind(JFd,G)
    H<-RC[,i]/Días_mes[i];RCd<-cbind(RCd,H)
  }
  GS<-GS[,-which(colMeans(is.na(GS)) >= 0.20)]
  GE<-GE[,-which(colMeans(is.na(GE)) >= 0.20)]
  GEC<-GEC[,-which(colMeans(is.na(GEC)) >= 0.20)]
  C2T<-C2T[,-which(colMeans(is.na(C2T)) >= 0.20)]
  DP<-DP[,-which(colMeans(is.na(DP)) >= 0.20)]
  D2<-D2[,-which(colMeans(is.na(D2)) >= 0.20)]
  JF<-JF[,-which(colMeans(is.na(JF)) >= 0.20)]
  RC<-RC[,-which(colMeans(is.na(RC)) >= 0.20)]
  mes_nam<-as.matrix(A1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                           sheet = "Demandas",range = "b1:m1", col_names = FALSE))
  mes_nam<-mes_nam[,-which(colMeans(is.na(mes_nam)) >= 0.20)]
  colnames(GS)<-colnames(GE)<-colnames(GEC)<-colnames(C2T)<-colnames(DP)<-colnames(D2)<-colnames(JF)<-colnames(RC)<-mes_nam
  GSd<-round(GSd,2);GEd<-round(GEd,2);GECd<-round(GECd,2);C2Td<-round(C2Td,2);DPd<-round(DPd,2);D2d<-round(D2d,2);JFd<-round(JFd,2);RCd<-round(RCd,2)
  GSd<-GSd[,-which(colMeans(is.na(GSd)) >= 0.20)];GSd[1,]<-GSd[1,]+GSd[6,];GSd[4,]<-GSd[4,]+GSd[5,]+GSd[6,];GSd[2,]<-GSd[2,]+GSd[9,]+GSd[10,];GSd[5,]<-GSd[5,]+GSd[11,]+GSd[12,]+GSd[13,];GSd<-rbind(GSd[1,],GSd[2,],GSd[3,],GSd[4,],GSd[5,])
  GEd<-GEd[,-which(colMeans(is.na(GEd)) >= 0.20)];GEd[1,]<-GEd[1,]+GEd[6,];GEd[4,]<-GEd[4,]+GEd[5,]+GEd[6,];GEd[2,]<-GEd[2,]+GEd[9,]+GEd[10,];GEd[5,]<-GEd[5,]+GEd[11,]+GEd[12,]+GEd[13,];GEd<-rbind(GEd[1,],GEd[2,],GEd[3,],GEd[4,],GEd[5,])
  GECd<-GECd[,-which(colMeans(is.na(GECd)) >= 0.20)];GECd[1,]<-GECd[1,]+GECd[6,];GECd[4,]<-GECd[4,]+GECd[5,]+GECd[6,];GECd[2,]<-GECd[2,]+GECd[9,]+GECd[10,];GECd[5,]<-GECd[5,]+GECd[11,]+GECd[12,]+GECd[13,];GECd<-rbind(GECd[1,],GECd[2,],GECd[3,],GECd[4,],GECd[5,])
  C2Td<-C2Td[,-which(colMeans(is.na(C2Td)) >= 0.20)];C2Td<-C2Td[1:2,]
  DPd<-DPd[,-which(colMeans(is.na(DPd)) >= 0.20)];DPd[1,]<-DPd[1,]+DPd[4,]+DPd[6,]+DPd[7,]+DPd[8,];DPd[2,]<-DPd[2,]+DPd[9,]+DPd[10,];DPd[5,]<-DPd[11,]+DPd[12,]+DPd[13,];DPd<-rbind(DPd[1,],DPd[2,],DPd[3,],DPd[5,])
  D2d<-D2d[,-which(colMeans(is.na(D2d)) >= 0.20)];D2d[1,]<-D2d[1,]+D2d[4,]+D2d[6,]+D2d[7,]+D2d[8,];D2d[2,]<-D2d[2,]+D2d[9,]+D2d[10,];D2d[5,]<-D2d[11,]+D2d[12,]+D2d[13,];D2d<-rbind(D2d[1,],D2d[2,],D2d[3,],D2d[5,])
  JFd<-JFd[,-which(colMeans(is.na(JFd)) >= 0.20)];JFd<-JFd[4:5,]
  RCd<-RCd[,-which(colMeans(is.na(RCd)) >= 0.20)];colnames(RCd)<-mes_nam;RCd<-RCd[1,]
  rownames(GSd)<-rownames(GEd)<-rownames(GECd)<-nam[1:5];rownames(DPd)<-c(nam[1:3],nam[5]);rownames(D2d)<-c(nam[1:3],nam[5])
  colnames(GSd)<-colnames(GEd)<-colnames(GECd)<-colnames(DPd)<-colnames(D2d)<-colnames(JFd)<-mes_nam

# PRODUCCIÓN POR REFINERÍA
  ## capacidades operativas en las 3 R.
  cap_nominal_operacion<-TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                         sheet = "Capacidad Operativa",range = "B2:N16", col_names = FALSE)
  dia<- TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                       sheet = "Días Operación",range = "B2:M16", col_names = FALSE)
  rend<- TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                        sheet = "Rendimiento",range = "B2:T29", col_names = FALSE)
  names_row_rend<- TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                  sheet = "Rendimiento",range = "A2:A29", col_names = FALSE)
  names_col_rend<- TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                  sheet = "Rendimiento",range = "B1:T1", col_names = FALSE)
  cap <- as.matrix(cap_nominal_operacion); dia <- as.matrix(dia); rend <- as.matrix(rend)
  names_row_rend<-as.matrix(names_row_rend);names_col_rend<-as.matrix(names_col_rend)
  rownames(dia)<-rownames(cap)<-c("CDU 1","VDU 1","CDU 2","VDU 2","VBU 1","VBU 2","FCC","CCR","HDT","HDS","PARSONS","UNIVERSAL","CAUTIVO","AMAZONAS I","AMAZONAS II")
  rownames(rend)<-names_row_rend<-as.matrix(names_row_rend);colnames(rend)<-names_col_rend
  cap_op<- cap
  i<-NULL
  for ( i in 2:13) {
    cap[,i]<-cap[,1]*cap[,i]*dia[,i-1]
  }
  cap_operativa_mes<-cap[,-1]#; cap_operativa_mes
  Planta_Gas_Estabilizadora<- as.matrix(TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                                       sheet = "Capacidad Operativa",range = "B17:N20", col_names = FALSE))
  cap_gas<-Planta_Gas_Estabilizadora<-Planta_Gas_Estabilizadora[-2,]
  dia_planta_gas<- as.matrix(TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                            sheet = "Días Operación",range = "B17:M18", col_names = FALSE))
  Planta_Gas_Estabilizadora
  i<-NULL
  for ( i in 2:13) {
    cap_gas[1,i]<-cap_gas[1,1]*cap_gas[1,i]*dia_planta_gas[1,i-1]
    cap_gas[2,i]<-cap_gas[2,i]*dia_planta_gas[2,i-1]*6.28981*28.31685/0.355
    cap_gas[3,i]<-cap_gas[3,i]*dia_planta_gas[2,i-1]*5.451*6.28981
  }
  cap_gas_mes<-cap_gas[,-1]
  
  ## Producción por R
  glp1_dat<-rend[1,1]*cap_operativa_mes[1,]+rend[1,2]*cap_operativa_mes[3,]
  acmb_dat<-rend[2,1]*cap_operativa_mes[1,]+rend[2,2]*cap_operativa_mes[3,]
  nlv_dat<-rend[3,1]*cap_operativa_mes[1,]+rend[3,2]*cap_operativa_mes[3,]
  nps_dat<-rend[4,1]*cap_operativa_mes[1,]+rend[4,2]*cap_operativa_mes[3,]-RC[1,]*0.2694-cap_operativa_mes[9,]+(cap_operativa_mes[9,]*rend[10,18]-cap_operativa_mes[8,])
  jfe_dat<-rend[5,1]*cap_operativa_mes[1,]+rend[5,2]*cap_operativa_mes[3,]
  ds2_dat<-rend[6,1]*cap_operativa_mes[1,]+rend[6,2]*cap_operativa_mes[3,]
  rat_dat<-rend[7,1]*cap_operativa_mes[1,]+rend[7,2]*cap_operativa_mes[3,]
  vgo_dat<-rend[8,3]*cap_operativa_mes[2,]+rend[8,4]*cap_operativa_mes[4,]
  rva_dat<-rend[9,3]*cap_operativa_mes[2,]+rend[9,4]*cap_operativa_mes[4,]
  vcmb_dat<-rend[2,5]*cap_operativa_mes[5,]+rend[2,6]*cap_operativa_mes[6,]
  nsv_dat<-rend[10,5]*cap_operativa_mes[5,]+rend[10,6]*cap_operativa_mes[6,]
  rsv_dat<-rend[11,5]*cap_operativa_mes[5,]+rend[11,6]*cap_operativa_mes[6,]
  ccmb_dat<-rend[2,7]*cap_operativa_mes[7,]
  ntr_dat<-rend[12,7]*cap_operativa_mes[7,]
  acc_dat<-rend[13,7]*cap_operativa_mes[7,]
  rcmb_dat<-rend[2,8]*cap_operativa_mes[8,]
  hid_dat<-rend[14,8]*cap_operativa_mes[8,]
  nrf_dat<-rend[15,8]*cap_operativa_mes[8,]
  dp_dat<-rend[28,19]*cap_operativa_mes[10,]
  nbl_dat<-rend[16,9]*cap_operativa_mes[11,]+rend[16,10]*cap_operativa_mes[12,]+rend[16,11]*cap_operativa_mes[13,]-C2T[2,]
  jfl_dat<-rend[17,9]*cap_operativa_mes[11,]+rend[17,10]*cap_operativa_mes[12,]+rend[17,11]*cap_operativa_mes[13,]
  d1l_dat<-rend[18,9]*cap_operativa_mes[11,]+rend[18,10]*cap_operativa_mes[12,]+rend[18,11]*cap_operativa_mes[13,]
  d2l_dat<-rend[19,9]*cap_operativa_mes[11,]+rend[19,10]*cap_operativa_mes[12,]+rend[19,11]*cap_operativa_mes[13,]
  fol_dat<-rend[20,9]*cap_operativa_mes[11,]+rend[20,10]*cap_operativa_mes[12,]+rend[20,11]*cap_operativa_mes[13,]
  svl_dat<-rend[21,11]*cap_operativa_mes[13,]
  mnl_dat<-rend[22,11]*cap_operativa_mes[13,]
  aol_dat<-rend[24,9]*cap_operativa_mes[11,]
  cmbs_dat<-rend[2,12]*cap_operativa_mes[14,]+rend[2,13]*cap_operativa_mes[15,]
  nbs_dat<-rend[16,12]*cap_operativa_mes[14,]+rend[16,13]*cap_operativa_mes[15,]+rend[26,15]*(cap_gas_mes[2,]+cap_gas_mes[3,])
  jfs_dat<-rend[17,12]*cap_operativa_mes[14,]+rend[17,13]*cap_operativa_mes[15,]
  d1s_dat<-rend[18,12]*cap_operativa_mes[14,]+rend[18,13]*cap_operativa_mes[15,]
  d2s_dat<-rend[19,12]*cap_operativa_mes[14,]+rend[19,13]*cap_operativa_mes[15,]
  fos_dat<-rend[20,12]*cap_operativa_mes[14,]+rend[20,13]*cap_operativa_mes[15,]
  sps_dat<-rend[23,12]*cap_operativa_mes[14,]+rend[23,13]*cap_operativa_mes[15,]
  rf3_produccion<-NULL;rf2_produccion<-NULL;rf1_produccion<-NULL
  rf3_produccion<-rbind(rf3_produccion,cmbs_dat,nbs_dat,jfs_dat,d1s_dat,d2s_dat,fos_dat,sps_dat)
  rf2_produccion<-rbind(rf2_produccion,nbl_dat,jfl_dat,d1l_dat,d2l_dat,fol_dat,svl_dat,mnl_dat,aol_dat)
  rf1_produccion<-rbind(rf1_produccion,glp1_dat,acmb_dat,nlv_dat,nps_dat,jfe_dat,ds2_dat,rat_dat,vgo_dat,rva_dat,vcmb_dat,nsv_dat,rsv_dat,ccmb_dat,ntr_dat,acc_dat,rcmb_dat,hid_dat,nrf_dat,dp_dat)
  cap_operativa_mes<-round(cap_operativa_mes,2);rf3_produccion<-round(rf3_produccion,2);rf2_produccion<-round(rf2_produccion,2);rf1_produccion<-round(rf1_produccion,2)
  cap_operativa_mes<-cap_operativa_mes[,-which(colMeans(is.na(cap_operativa_mes)) >= 0.20)]
  rf3_produccion<-rf3_produccion[,-which(colMeans(is.na(rf3_produccion)) >= 0.20)]
  rf2_produccion<-rf2_produccion[,-which(colMeans(is.na(rf2_produccion)) >= 0.20)]
  rf1_produccion<-rf1_produccion[,-which(colMeans(is.na(rf1_produccion)) >= 0.20)]
  dia<-dia[,-which(colMeans(is.na(dia)) >= 0.20)]
  colnames(cap_operativa_mes)<-colnames(rf3_produccion)<-colnames(rf2_produccion)<-colnames(rf1_produccion)<-colnames(dia)<-mes_nam
  cap_op;cap_operativa_mes;rend;rf3_pd<-rf3_produccion;rf2_pd<-rf2_produccion;rf1_pd<-rf1_produccion;dia
  i<-NULL
  for ( i in 1:ld) {
    rf3_pd[,i]<-rf3_produccion[,i]/Días_mes[i]
    rf2_pd[,i]<-rf2_produccion[,i]/Días_mes[i]
    rf1_pd[,i]<-rf1_produccion[,i]/Días_mes[i]
  }
  

# MEZCLAS EN TERMINALES GASOLINAS
  cap_tanques<- as.matrix(TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                                       sheet = "Cap. Op. Tanques",range = "b1:p14"))
  rownames(cap_tanques)<-as.matrix(A1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                           sheet = "Cap. Op. Tanques",range = "a2:a14", col_names = FALSE))
  Vol_inicial<- as.matrix(TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                          sheet = "Vol. Inicial",range = "b1:p14"))
  rownames(Vol_inicial)<-as.matrix(A1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                    sheet = "Vol. Inicial",range = "a2:a14", col_names = FALSE))
  ##  R3
  i<-NULL; mz_ge_rsh<-NULL;B<-NULL
  fo<- MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                      sheet = "G-RSH",range = "B2:c2", col_names = FALSE)
  rest<- MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                        sheet = "G-RSH",range = "B3:c8", col_names = FALSE)
  dir<- MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                       sheet = "G-RSH",range = "d3:d8", col_names = FALSE)
  for ( i in 1:ld) {
    sol<- MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                         sheet = "G-RSH",range = "e3:e8", col_names = FALSE)
    fo <- as.matrix(fo);rest<-as.matrix(rest);dir<-as.matrix(dir);sol<-as.matrix(sol)
    sol<-GE[3,i]*sol
    mylp<-lp ("min", fo,rest,
              dir, sol)
    mz_ge_rsh<-as.matrix(mylp$solution)
    rf3_produccion[2,i]<-rf3_produccion[2,i]-mz_ge_rsh[1,1]
    B<-cbind(B,mz_ge_rsh)
  }
  mz_ge_rsh<-B
  mz_ge_rsh
  ###calculo calidades productos R3
  i<-NULL;k<-NULL
  for( i in 1:ld) {
    k<-cbind(k,rest[3:6,]%*%mz_ge_rsh[,i]/sum(mz_ge_rsh[,i]))
  }
  cal_mz_ge_rsh<-round(k,2);k[is.nan(k)]<-0;cal_mz_ge_rsh[2,]<-cal_mz_ge_rsh[2,]*119
  cal_mz_ge_rsh[3,]<-exp(log(cal_mz_ge_rsh[3,])/1.25)
  rownames(mz_ge_rsh)<-c("Nafta base R3","G. Súper preparación de G. Extra")
  rownames(cal_mz_ge_rsh)<-c("RON","Azufre, [ppm]","PVR, [kPa]","Aromáticos, [% v/v]")
  
  ## T1
  G_EXTRA_BEAT<-GE[4,]+GE[7,]+GE[8,]
  nbs1<-rf3_produccion[2,]
  fo<- MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                      sheet = "G-T1",range = "B2:d2", col_names = FALSE)
  rest<- MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                        sheet = "G-T1",range = "B3:d11", col_names = FALSE)
  dir<- MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                       sheet = "G-T1",range = "e3:e11", col_names = FALSE)
  i<-NULL; mz_ge_beat<-NULL;B<-NULL;sol<-NULL
  for ( i in 1:ld) {
    sol<- MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                         sheet = "G-T1",range = "f3:f11", col_names = FALSE)
    fo <- as.matrix(fo);rest<-as.matrix(rest);dir<-as.matrix(dir);sol<-as.matrix(sol)
    sol[1:8]<-G_EXTRA_BEAT[i]*sol[1:8]
    sol[9]<-sol[9]*nbs1[i]
    #sol
    mylp<-lp ("min", fo,rest,
              dir, sol)
    mz_ge_beat<-as.matrix(mylp$solution)
    nbs1[i]<-nbs1[i]-mz_ge_beat[1,1]
    B<-cbind(B,mz_ge_beat)
  }
  mz_ge_beat<-B
  #nbs1<-round(nbs1,2)
  nbs1<-nbs1
  nbs1
  mz_ge_beat
  ###calculo calidades productos T1
  i<-NULL;k<-NULL
  for( i in 1:ld) {
    k<-cbind(k,rest[3:8,]%*%mz_ge_beat[,i]/sum(mz_ge_beat[,i]))
  }
  cal_mz_ge_beat<-round(k,2);k[is.nan(k)]<-0;cal_mz_ge_beat[2,]<-cal_mz_ge_beat[2,]*119
  cal_mz_ge_beat[3,]<-exp(log(cal_mz_ge_beat[3,])/1.25)
  rownames(mz_ge_beat)<-c("Nafta base RSH","G. Súper preparación de G. Extra","G. Extra de R1")
  rownames(cal_mz_ge_beat)<-c("RON","Azufre, [ppm]","PVR, [kPa]","Aromáticos, [% v/v]","Olefinas, [% v/v]","Benceno, [% v/v]")
  
  ##R1
  mz_ge_esm<-GE[6,]+mz_ge_beat[3,]
  mz_gs_esm<-mz_ge_beat[2,]+mz_ge_rsh[2,]+GS[1,]+GS[4,]+GS[6,]+GS[7,]+GS[8,]
  mz_ge_esm;mz_gs_esm
  gyn_rf1<-rbind(mz_gs_esm, mz_ge_esm,GEC[1,]*eta,C2T[1,],rf1_produccion[3,],rf1_produccion[4,],rf1_produccion[18,],rf1_produccion[14,])
  fo<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                               sheet = "G-R1", range = "b2:v3", na = "NA"))
  rest<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                 sheet = "G-R1", range = "b4:v33", na = "NA", col_names = FALSE))
  dir<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                sheet = "G-R1", range = "w4:w33", na = "NA", col_names = FALSE))
  i<-NULL; gasolinas_rf1<-NULL;B<-NULL;sol<-NULL
  for ( i in 1:ld) {
    sol<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                  sheet = "G-R1", range = "x4:x33", na = "NA", col_names = FALSE))
    sol[1:8]<-gyn_rf1[1,i]*sol[1:8]
    sol[9:16]<-gyn_rf1[2,i]*sol[9:16]
    sol[17:24]<-gyn_rf1[3,i]*sol[17:24]
    sol[25:26]<-gyn_rf1[4,i]*sol[25:26]
    sol[27]<-gyn_rf1[5,i]*sol[27]
    sol[28]<-gyn_rf1[6,i]*sol[28]
    sol[29]<-gyn_rf1[7,i]*sol[29]
    sol[30]<-gyn_rf1[8,i]*sol[30]
    mylp<-lp ("min", fo,rest,
              dir, sol)
    gasolinas_rf1<-as.matrix(mylp$solution)
    B<-cbind(B,gasolinas_rf1)
  }
  gasolinas_rf1<-round(B,2)
  i<-NULL;nl<-NULL;np<-NULL;nr<-NULL;nt<-NULL;nao<-NULL
  for ( i in 1:4) {
    nl<-rbind(nl,B[i+(i-1)*4,])
    np<-rbind(np,B[i+(i-1)*4+1,])
    nr<-rbind(nr,B[i+(i-1)*4+2,])
    nt<-rbind(nt,B[i+(i-1)*4+3,])
    nao<-rbind(nao,B[i+(i-1)*4+4,])
  }
  i<-NULL;nl1<-NULL;np1<-NULL;nr1<-NULL;nt1<-NULL;nao1<-NULL
  for( i in 1:ld) {
    nl1<-cbind(nl1,sum(nl[,i]))
    np1<-cbind(np1,sum(np[,i]))
    nr1<-cbind(nr1,sum(nr[,i]))
    nt1<-cbind(nt1,sum(nt[,i]))
    nao1<-cbind(nao1,sum(nao[,i]))
  }
  nl<-gyn_rf1[5,]-nl1;np<-gyn_rf1[6,]-np1;nr<-gyn_rf1[7,]-nr1;nt<-gyn_rf1[8,]-nt1
  naoree<-nao1
  nbo<-round(rbind(nl,np,nr,nt),2)
  ###cálculo calidades productos R1
  i<-NULL;k<-NULL;x<-NULL;x1<-NULL;x2<-NULL
  for( i in 1:ld) {
    k<-cbind(k,rest[3:8,1:4]%*%nbo[,i]/sum(nbo[,i]))
    x<-cbind(x,rest[3:8,1:5]%*%gasolinas_rf1[1:5,i]/sum(gasolinas_rf1[1:5,i]))
    x1<-cbind(x1,rest[3:8,1:5]%*%gasolinas_rf1[6:10,i]/sum(gasolinas_rf1[6:10,i]))
    x2<-cbind(x2,rest[3:8,1:5]%*%gasolinas_rf1[11:15,i]/sum(gasolinas_rf1[11:15,i]))
  }
  cal_nbo_re<-round(k,2);k[is.nan(k)]<-0
  cal_nbo_re1<-round(k,2);k[is.nan(k)]<-0;cal_nbo_re1[2,]<-cal_nbo_re1[2,]*119;cal_nbo_re1[3,]<-exp(log(cal_nbo_re1[3,])/1.25)
  cal_mz_gs_rf1<-round(x,2);x[is.nan(x)]<-0;cal_mz_gs_rf1[2,]<-cal_mz_gs_rf1[2,]*117.5;cal_mz_gs_rf1[3,]<-exp(log(cal_mz_gs_rf1[3,])/1.25)
  cal_mz_ge_rf1<-round(x1,2);x1[is.nan(x1)]<-0;cal_mz_ge_rf1[2,]<-cal_mz_ge_rf1[2,]*119;cal_mz_ge_rf1[3,]<-exp(log(cal_mz_ge_rf1[3,])/1.25)
  cal_mz_gp_rf1<-round(x2,2);x[is.nan(x2)]<-0;cal_mz_gp_rf1[2,]<-cal_mz_gp_rf1[2,]*119;cal_mz_gs_rf1[3,]<-exp(log(cal_mz_gp_rf1[3,])/1.25)
  gasolinas_rf1
  gs_re<-gasolinas_rf1[1:5,]
  ge_re<-gasolinas_rf1[6:10,]
  gp_re<-gasolinas_rf1[11:15,]
  gpa_re<-gasolinas_rf1[16:17,]
  rownames(gs_re)<-rownames(ge_re)<-rownames(gp_re)<-c("NL","NP","NR","NT","NAO")
  rownames(nbo)<-c("NL","NP","NR","NT")
  rownames(gpa_re)<-c("NL","NP")
  rownames(cal_nbo_re1)<-rownames(cal_mz_gs_rf1)<-rownames(cal_mz_ge_rf1)<-rownames(cal_mz_gp_rf1)<-c("RON","Azufre, [ppm]","PVR, [kPa]","Aromáticos, [% v/v]","Olefinas, [% v/v]","Benceno, [% v/v]")
  
  ##R2
  mz_gs_rl<-GS[2,]+GS[10,]
  mz_ge_rl<-GE[9,]
  mz_gp_rl<-GEC[2,]*eta+GEC[10,]*eta
  i<-NULL;k<-NULL
  for( i in 1:ld) {
    k<-cbind(k,sum(nbo[,i]))
  }
  nbo_re<-k
  nbo_rl<-rf2_produccion[1,]- C2T[2,]
  gyn_rf2<-rbind(mz_gs_rl,mz_ge_rl,mz_gp_rl,nbo_re,nbo_rl)
  fo<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                               sheet = "G-R2", range = "b2:m3", na = "NA"))
  dir<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                sheet = "G-R2", range = "n4:n29", na = "NA", col_names = FALSE))
  i<-NULL; gasolinas_rf2<-NULL;B<-NULL;sol<-NULL
  for ( i in 1:ld) {
    rest<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                   sheet = "G-R2", range = "b4:m29", na = "NA", col_names = FALSE))
    sol<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                  sheet = "G-R2", range = "o4:o29", na = "NA", col_names = FALSE))
    sol[1:8]<-gyn_rf2[1,i]*sol[1:8]
    sol[9:16]<-gyn_rf2[2,i]*sol[9:16]
    sol[17:24]<-gyn_rf2[3,i]*sol[17:24]
    sol[25]<-gyn_rf2[4,i]*sol[25]
    sol[26]<-gyn_rf2[5,i]*sol[26]
    rest[3:8,1]<-rest[11:16,5]<-rest[19:24,9]<-cal_nbo_re[,i]
    mylp<-lp ("min", fo,rest,
              dir, sol)
    gasolinas_rf2<-as.matrix(mylp$solution)
    B<-cbind(B,gasolinas_rf2)
  }
  gasolinas_rf2<-round(B,2)
  i<-NULL;nbo_re1<-NULL;nbo_rl1<-NULL;nao1<-NULL;n801<-NULL
  for ( i in 1:3) {
    nbo_re1<-rbind(nbo_re1,B[i+(i-1)*3,])
    nbo_rl1<-rbind(nbo_rl1,B[i+(i-1)*3+1,])
    nao1<-rbind(nao1,B[i+(i-1)*3+2,])
    n801<-rbind(n801,B[i+(i-1)*3+3,])
  }
  i<-NULL;nbo_re2<-NULL;nbo_rl2<-NULL;nao2<-NULL;n802<-NULL
  for( i in 1:ld) {
    nbo_re2<-cbind(nbo_re2,sum(nbo_re1[,i]))
    nbo_rl2<-cbind(nbo_rl2,sum(nbo_rl1[,i]))
    nao2<-cbind(nao2,sum(nao1[,i]))
    n802<-cbind(n802,sum(n801[,i]))
  }
  nbo_re1<-gyn_rf2[4,]-nbo_re2;nbo_rl1<-gyn_rf2[5,]-nbo_rl2
  naorf2<-nao2;n80rf2<-n802
  nborf2<-round(rbind(nbo_re1,nbo_rl1),2)
  ###cálculo calidades productos R2
  rest<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                 sheet = "G-R2", range = "b4:m29", na = "NA", col_names = FALSE))
  i<-NULL;x<-NULL;x1<-NULL;x2<-NULL
  for( i in 1:ld) {
    x<-cbind(x,rest[3:8,1:4]%*%gasolinas_rf2[1:4,i]/sum(gasolinas_rf2[1:4,i]))
    x1<-cbind(x1,rest[3:8,1:4]%*%gasolinas_rf2[5:8,i]/sum(gasolinas_rf2[5:8,i]))
    x2<-cbind(x2,rest[3:8,1:4]%*%gasolinas_rf2[9:12,i]/sum(gasolinas_rf2[9:12,i]))
  }
  cal_mz_gs_rf2<-round(x,2);x[is.nan(x)]<-0;cal_mz_gs_rf2[2,]<-cal_mz_gs_rf2[2,]*119;cal_mz_gs_rf2[3,]<-exp(log(cal_mz_gs_rf2[3,])/1.25)
  cal_mz_ge_rf2<-round(x1,2);x1[is.nan(x1)]<-0;cal_mz_ge_rf2[2,]<-cal_mz_ge_rf2[2,]*119;cal_mz_ge_rf2[3,]<-exp(log(cal_mz_ge_rf2[3,])/1.25)
  cal_mz_gp_rf2<-round(x2,2);x[is.nan(x2)]<-0;cal_mz_gp_rf2[2,]<-cal_mz_gp_rf2[2,]*119;cal_mz_gs_rf2[3,]<-exp(log(cal_mz_gp_rf2[3,])/1.25)
  gasolinas_rf2
  gs_rl<-gasolinas_rf2[1:4,]
  ge_rl<-gasolinas_rf2[5:8,]
  gp_rl<-gasolinas_rf2[9:12,]
  rownames(gs_rl)<-rownames(ge_rl)<-rownames(gp_rl)<-c("NBRE","NBRL","NAO 93 RON","NAO 80 RON")
  rownames(cal_mz_gs_rf2)<-rownames(cal_mz_ge_rf2)<-rownames(cal_mz_gp_rf2)<-c("RON","Azufre, [ppm]","PVR, [kPa]","Aromáticos, [% v/v]","Olefinas, [% v/v]","Benceno, [% v/v]")
  
  ##T2
  mz_gs_T2<-GS[5,]+GS[10,]+GS[11,]+GS[12,]
  mz_gp_T2<-GEC[5,]*eta+GEC[10,]*eta+GEC[11,]*eta+GEC[12,]*eta
  gyn_T2<-rbind(mz_gs_T2,mz_gp_T2,nborf2)
  fo<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                               sheet = "G-T2", range = "b2:i3", na = "NA"))
  dir<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                sheet = "G-T2", range = "j4:j21", na = "NA", col_names = FALSE))
  i<-NULL; gasolinas_T2<-NULL;B<-NULL;sol<-NULL
  for ( i in 1:ld) {
    rest<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                   sheet = "G-T2", range = "b4:i21", na = "NA", col_names = FALSE))
    sol<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                  sheet = "G-T2", range = "k4:k21", na = "NA", col_names = FALSE))
    sol[1:8]<-gyn_T2[1,i]*sol[1:8]
    sol[9:16]<-gyn_T2[2,i]*sol[9:16]
    sol[17]<-gyn_T2[3,i]*sol[17]
    sol[18]<-gyn_T2[4,i]*sol[18]
    rest[3:8,1]<-rest[11:16,5]<-cal_nbo_re[,i]
    mylp<-lp ("min", fo,rest,
              dir, sol)
    gasolinas_T2<-as.matrix(mylp$solution)
    B<-cbind(B,gasolinas_T2)
  }
  gasolinas_T2<-round(B,2)
  ###cálculo calidades productos T2
  rest<-as.matrix(MPL_DREF_Modelo1 <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", 
                                                 sheet = "G-T2", range = "b4:i21", na = "NA", col_names = FALSE))
  i<-NULL;x<-NULL;x1<-NULL;x2<-NULL
  for( i in 1:ld) {
    x<-cbind(x,rest[3:8,1:4]%*%gasolinas_T2[1:4,i]/sum(gasolinas_T2[1:4,i]))
    x1<-cbind(x1,rest[3:8,1:4]%*%gasolinas_T2[5:8,i]/sum(gasolinas_T2[5:8,i]))
  }
  cal_mz_gs_t2<-round(x,2);x[is.nan(x)]<-0;cal_mz_gs_t2[2,]<-cal_mz_gs_t2[2,]*119;cal_mz_gs_t2[3,]<-exp(log(cal_mz_gs_t2[3,])/1.25)
  cal_mz_gp_t2<-round(x1,2);x1[is.nan(x1)]<-0;cal_mz_gp_t2[2,]<-cal_mz_gp_t2[2,]*119;cal_mz_gp_t2[3,]<-exp(log(cal_mz_gp_t2[3,])/1.25)
  gasolinas_T2
  gs_T2<-gasolinas_T2[1:4,]
  gp_T2<-gasolinas_T2[5:8,]
  rownames(gs_T2)<-rownames(gp_T2)<-c("NBRE","NBRL","NAO 93 RON","NAO 80 RON")
  rownames(cal_mz_gs_t2)<-rownames(cal_mz_gp_t2)<-c("RON","Azufre, [ppm]","PVR, [kPa]","Aromáticos, [% v/v]","Olefinas, [% v/v]","Benceno, [% v/v]")
  
  i<-NULL;nbo_re1<-NULL;nbo_rl1<-NULL;nao1<-NULL;n801<-NULL
  for ( i in 1:2) {
    nbo_re11<-rbind(nbo_re1,B[i+(i-1)*3,])
    nbo_rl11<-rbind(nbo_rl1,B[i+(i-1)*3+1,])
    nao1<-rbind(nao1,B[i+(i-1)*3+2,])
    n801<-rbind(n801,B[i+(i-1)*3+3,])
  }
  i<-NULL;nbo_re3<-NULL;nbo_rl3<-NULL;nao2<-NULL;n802<-NULL
  for( i in 1:ld) {
    nbo_re3<-cbind(nbo_re3,sum(nbo_re11[,i]))
    nbo_rl3<-cbind(nbo_rl3,sum(nbo_rl11[,i]))
    nao2<-cbind(nao2,sum(nao1[,i]))
    n802<-cbind(n802,sum(n801[,i]))
  }
  nbo_re3<-gyn_T2[3,]-nbo_re3;nbo_rl3<-gyn_T2[4,]-nbo_rl3
  naorT2<-nao2;n80rT2<-n802
  nborT2<-round(rbind(nbo_re3,nbo_rl3,nbs1),2)
  nborT2<-rbind(nbo_re3,nbo_rl3,nbs1)
  nao_t2<-gs_T2[3,]+gp_T2[3,]
  
  nao_r1<-naoree
  nao_r2<-naorf2
  nao_t2<-naorT2
  
  n80_r2<-n80rf2
  n80_t2<-n80rT2
  
  nbo_r1<-nbo
  
  nao<-rbind(nao_r1,nao_r2,nao_t2)
  n80<-rbind(n80_r2,n80_t2)
  rownames(nao)<-c("R1","R2","T2");rownames(n80)<-c("R2","T2");rownames(nbo)<-c("NL","NP","NR","NT");rownames(nborT2)<-c("NB-R1","NB-R2","NB-R3")
  a_meses<- read_excel("TG-UPS-2021-LMP&B.xlsx", sheet = "Demandas", range = "b1:n1", col_names = FALSE, na = "0")
  colnames(nao)<-colnames(n80)<-a_meses<-a_meses[,-which(colMeans(is.na(a_meses)) >= 0.20)]#eliminar columnas vacías

  ##Otros productos
  restm<- as.matrix(TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx",sheet = "M_I", range = "b3:ak50", col_names = FALSE))
  dirm<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx",sheet = "M_I", range = "al3:al50", col_names = FALSE))
  fom<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx", sheet = "M_I", range = "d2:ak2", col_names = FALSE))
  solm<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("C:/Users/Usuario/Desktop/TG-UPS-2021-LMP&B/TG-UPS-2021-LMP&B.xlsx",sheet = "M_I", range = "am3:am50", col_names = FALSE))
  DM1<-dirm[1:(ld)]
  DM2<-dirm[13:(12+ld)]
  DM3<-dirm[25:(24+ld)]
  DM4<-dirm[37:(36+ld)]
  dirm<-c(DM1,DM2,DM3,DM4)
  SM1<-solm[1:(ld)]
  SM2<-solm[13:(12+ld)]
  SM3<-solm[25:(24+ld)]
  FM1<-fom[1:(ld)]
  FM2<-fom[13:(12+ld)]
  FM3<-fom[25:(24+ld)]
  fom<-c(FM1,FM2,FM3)
  AM1<-restm[1:(ld),1:(ld)]
  AM2<-restm[13:(12+ld),1:(ld)]
  AM3<-restm[25:(24+ld),1:(ld)]
  AM4<-restm[1:(ld),13:(ld+12)]
  AM5<-restm[13:(12+ld),13:(ld+12)]
  AM7<-restm[1:(ld),25:(ld+24)]
  AM8<-restm[13:(12+ld),25:(ld+24)]
  AM9<-restm[25:(24+ld),25:(ld+24)]
  AM10<-restm[37:(36+ld),1:(ld)]
  AM11<-restm[37:(36+ld),13:(ld+12)]
  AM12<-restm[37:(36+ld),25:(ld+24)]
  AM13<-cbind(AM1,AM4,AM7)
  AM14<-cbind(AM2,AM5,AM8)
  AM16<-cbind(AM10,AM11,AM12)
  #DP
  AM6<-restm[25:(24+ld),13:(ld+12)]/285000
  AM15<-cbind(AM3,AM6,AM9)
  restm1<-rbind(AM13,AM14,AM15,AM16)
  SM1<-Demandas_mes[5,]-rf1_produccion[19,]
  DPd1<-NULL
  for ( i in 1:ld) {
    DPd1[i]<-sum(DPd[,i])
  }
  SM2<-DPd1
  SM4<-round(SM1/285000,0)
  solm1<-c(SM1,SM2,SM3,SM4)
  vidp<-sum(Vol_inicial[5,])
  solm1[1]<-solm1[1]-vidp
  mylp<-lp ("min", fom,restm1,
            dirm,solm1, int.vec=c((2*ld+1):(3*ld)))
  DP_I<-as.matrix(mylp$solution)
  
  #D2
  AM6<-restm[25:(24+ld),13:(ld+12)]/285000
  AM15<-cbind(AM3,AM6,AM9)
  restm2<-rbind(AM13,AM14,AM15,AM16)
  SM1<-Demandas_mes[6,]-rf1_produccion[6,]+cap_operativa_mes[10,]-rf2_produccion[4,]-rf3_produccion[5,]
  D2d1<-NULL
  for ( i in 1:ld) {
    D2d1[i]<-sum(D2d[,i])*15
  }
  SM2<-D2d1
  SM4<-round(SM1/285000,0)
  solm2<-c(SM1,SM2,SM3,SM4)
  solm2[1]<-solm2[1]-sum(Vol_inicial[6,])
  mylp<-lp ("min", fom,restm2,
            dirm,solm2, int.vec=c((2*ld+1):(3*ld)))
  D2_I<-as.matrix(mylp$solution)
  
  #JF
  AM6<-restm[25:(24+ld),13:(ld+12)]/40000
  AM15<-cbind(AM3,AM6,AM9)
  restm3<-rbind(AM13,AM14,AM15,AM16)
  SM1<-Demandas_mes[7,]-rf1_produccion[5,]-rf2_produccion[2,]-rf3_produccion[3,]+7000
  JFd1<-NULL
  for ( i in 1:ld) {
    JFd1[i]<-sum(JFd[,i])
  }
  SM2<-JFd1
  SM4<-round(SM1/40000,0)
  solm3<-c(SM1,SM2,SM3,SM4)
  solm3[1]<-solm3[1]-sum(Vol_inicial[7,])
  mylp<-lp ("min", fom,restm3,
            dirm,solm3, int.vec=c((2*ld+1):(3*ld)))
  JF_I<-as.matrix(mylp$solution)
  
###Soluciones.
  # Naftas residuales
  nborT2
  #print(xtable(nborT2), include.rownames = TRUE) 
  #Naftas Importadas
    #Nafta Alto Octano 93 RON
    nao
    #Nafta 80 RON
    n80
  #Preparación de gasolinas
    #Demandas/mes por terminal
    #GS
    GS
    #GE
    GE
    #GEC
    GEC
    #C2T
    C2T
    #R1
      #GS
      gs_re;cal_mz_gs_rf1
      #print(xtable(gs_re), include.rownames = TRUE)
      #print(xtable(cal_mz_gs_rf1), include.rownames = TRUE)
      #GE
      ge_re;cal_mz_ge_rf1
      #GEC
      gp_re;cal_mz_gp_rf1
      #GPA
      gpa_re
      #NB-R1
      nbo;cal_nbo_re1
    #R2 
      #GS
      gs_rl;cal_mz_gs_rf2
      #GE
      ge_rl;cal_mz_ge_rf2
      #GEC
      gp_rl;cal_mz_gp_rf2
    #R3
      #GE
      mz_ge_rsh;cal_mz_ge_rsh
    #T1
      #GE
      mz_ge_beat;cal_mz_ge_beat
    #T2
      #GS     
      gs_T2;cal_mz_gs_t2
      #GEC     
      gp_T2;cal_mz_gp_t2
    #Importaciones x mes
      DP2J<-rbind(round(colSums(nao)/295000,0),round(colSums(n80)/295000,0),DP_I[(2*ld+1):(3*ld)],D2_I[(2*ld+1):(3*ld)],JF_I[(2*ld+1):(3*ld)])
      rownames(DP2J)<-c("NAO","N80","DP","D2","JF")
      DP2J
      #print(xtable(DP2J), include.rownames = TRUE)  
      #print(xtable(nao), include.rownames = TRUE)
      #print(xtable(n80), include.rownames = TRUE)

      
      #Descargas por dia por terminal
      if (Días<=100){
        Cap_tanques<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("TG-UPS-2021-LMP&B.xlsx", 
                                                               sheet = "Cap. Op. Tanques", range = "b1:h14"))
        rownames(Cap_tanques)<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("TG-UPS-2021-LMP&B.xlsx", 
                                                                         sheet = "Cap. Op. Tanques", range = "a2:a14", col_names = FALSE))
        Matriz_rest<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("TG-UPS-2021-LMP&B.xlsx", 
                                                               sheet = "M_STf", range = "D1:kq701"))
        Matriz_sol<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("TG-UPS-2021-LMP&B.xlsx", 
                                                              sheet = "M_STf", range = "b2:b701", col_names = FALSE))
        Matriz_dir<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("TG-UPS-2021-LMP&B.xlsx", 
                                                              sheet = "M_STf", range = "c2:c701", col_names = FALSE))
        Matriz_fo<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("TG-UPS-2021-LMP&B.xlsx", 
                                                             sheet = "M_STf", range = "D702:kq702", col_names = FALSE))
        Matriz_rest_2<-as.matrix(TG_UPS_2021_LMP_B <- read_excel("TG-UPS-2021-LMP&B.xlsx", 
                                                                 sheet = "M_STf", range = "D703:kq799", col_names = FALSE))
        rest_S1_2<-Matriz_rest_2[1:Días,1:Días]
        rest_I1_2<-Matriz_rest_2[1:Días,101:(100+Días)]
        rest_q1_2<-Matriz_rest_2[1:Días,201:(200+Días)]
        rest8<-cbind(rest_S1_2,rest_I1_2,rest_q1_2)
        
        rest_S1<-Matriz_rest[1:Días,1:Días]
        rest_S2<-Matriz_rest[101:(100+Días),1:Días]
        rest_S3<-Matriz_rest[201:(200+Días),1:Días]
        rest_S4<-Matriz_rest[301:(300+Días),1:Días]
        rest_S5<-Matriz_rest[401:(400+Días),1:Días]
        rest_S6<-Matriz_rest[501:(500+Días),1:Días]
        rest_S7<-Matriz_rest[601:(600+Días),1:Días]
        rest_I1<-Matriz_rest[1:Días,101:(100+Días)]
        rest_I2<-Matriz_rest[101:(100+Días),101:(100+Días)]
        rest_I3<-Matriz_rest[201:(200+Días),101:(100+Días)]
        rest_I4<-Matriz_rest[301:(300+Días),101:(100+Días)]
        rest_I5<-Matriz_rest[401:(400+Días),101:(100+Días)]
        rest_I6<-Matriz_rest[501:(500+Días),101:(100+Días)]
        rest_I7<-Matriz_rest[601:(600+Días),101:(100+Días)]
        rest_q1<-Matriz_rest[1:Días,201:(200+Días)]
        rest_q2<-Matriz_rest[101:(100+Días),201:(200+Días)]
        rest_q3<-Matriz_rest[201:(200+Días),201:(200+Días)]
        rest_q4<-Matriz_rest[301:(300+Días),201:(200+Días)]
        rest_q5<-Matriz_rest[401:(400+Días),201:(200+Días)]
        rest_q6<-Matriz_rest[501:(500+Días),201:(200+Días)]
        rest_q7<-Matriz_rest[601:(600+Días),201:(200+Días)]
        rest1<-cbind(rest_S1,rest_I1,rest_q1)
        rest2<-cbind(rest_S2,rest_I2,rest_q2)
        rest3<-cbind(rest_S3,rest_I3,rest_q3)
        rest6<-cbind(rest_S6,rest_I6,rest_q6)
        rest7<-cbind(rest_S7,rest_I7,rest_q7)
        
        dir_S1<-Matriz_dir[1:Días]
        dir_S2<-Matriz_dir[101:(100+Días)]
        dir_S3<-Matriz_dir[201:(200+Días)]
        dir_S4<-Matriz_dir[301:(300+Días)]
        dir_S5<-Matriz_dir[401:(400+Días)]
        dir_S6<-Matriz_dir[501:(500+Días)]
        dir_S7<-Matriz_dir[601:(600+Días)]
        
        fo_S1<-Matriz_fo[1:Días]
        fo_S2<-Matriz_fo[101:(100+Días)]
        fo_S3<-Matriz_fo[201:(200+Días)]
        fo<-c(fo_S1,fo_S2,fo_S3)
        
        #DP-R1
        i<-NULL; PDPd<-NULL; A<-NULL
        for ( i in 1:ld) {
          A<-rf1_produccion[19,i]/Días_mes[i];PDPd<-cbind(PDPd,A)
        }
        
        i<-NULL;sol_DP_d<-NULL;B<-C<-NULL
        for ( i in 1:ld) {
          sol_DP_d<-rep(DPd[1,i],Días_mes[i]);B<-c(B,sol_DP_d)
          sol_DP_p<-rep(PDPd[1,i],Días_mes[i]);C<-c(C,sol_DP_p)
        }
        sol_S1<-Matriz_sol[1:Días]*B
        sol_S1[1]<-sol_S1[1]-Vol_inicial[5,1]-Vol_inicial[5,6]
        sol_S2<-Matriz_sol[101:(100+Días)]*B-C
        sol_S3<-Matriz_sol[201:(200+Días)]*0.8*(Cap_tanques[5,1]+Cap_tanques[5,6])
        sol_S4<-Matriz_sol[301:(300+Días)]
        sol_S5<-Matriz_sol[401:(400+Días)]
        sol_S6<-Matriz_sol[501:(500+Días)]*285000
        #sol_S7<-Matriz_sol[601:(600+Días)]
        
        rest_I4<-rest_I4/(3500*24)
        rest4<-cbind(rest_S4,rest_I4,rest_q4)
        
        rest_I5<-rest_I5/60000
        rest5<-cbind(rest_S5,rest_I5,rest_q5)
        
        rest<-rbind(rest1,rest2,rest3,rest4,rest8)
        dir<-c(dir_S1,dir_S2,dir_S3,dir_S3,dir_S3)
        sol<-c(sol_S1,sol_S2,sol_S3,sol_S4,sol_S6)
        
        mylp<-lp ("min", fo,rest,
                  dir,sol, binary.vec=c((2*Días+1):(3*Días)))
        DP_Id<-as.matrix(mylp$solution)
        DP_Id
        
        #D2
        i<-NULL; PD2d<-NULL; A<-NULL
        for ( i in 1:ld) {
          A<-rf1_produccion[19,i]/Días_mes[i];PD2d<-cbind(PD2d,A)
        }
        
        i<-NULL;sol_D2_d<-NULL;B<-C<-NULL
        for ( i in 1:ld) {
          sol_D2_d<-rep(D2d[1,i],Días_mes[i]);B<-c(B,sol_D2_d)
          sol_D2_p<-rep(PD2d[1,i],Días_mes[i]);C<-c(C,sol_D2_p)
        }
        sol_S1<-Matriz_sol[1:Días]*B
        sol_S1[1]<-sol_S1[1]-Vol_inicial[6,1]-Vol_inicial[6,6]
        sol_S2<-Matriz_sol[101:(100+Días)]*B
        sol_S3<-Matriz_sol[201:(200+Días)]*0.8*(Cap_tanques[6,1]+Cap_tanques[6,6])
        sol_S4<-Matriz_sol[301:(300+Días)]
        sol_S5<-Matriz_sol[401:(400+Días)]
        sol_S6<-Matriz_sol[501:(500+Días)]*285000
        #sol_S7<-Matriz_sol[601:(600+Días)]
        
        rest_I4<-rest_I4/(3500*24)
        rest4<-cbind(rest_S4,rest_I4,rest_q4)
        
        rest_I5<-rest_I5/60000
        rest5<-cbind(rest_S5,rest_I5,rest_q5)
        
        rest<-rbind(rest1,rest2,rest3,rest4)
        dir<-c(dir_S1,dir_S2,dir_S3,dir_S3)
        sol<-c(sol_S1,sol_S2,sol_S3,sol_S4)
        
        mylp<-lp ("min", fo,rest,
                  dir,sol, binary.vec=c((2*Días+1):(3*Días)))
        D2_Id<-as.matrix(mylp$solution)
        D2_Id
      }