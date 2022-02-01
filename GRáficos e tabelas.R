library(tidyverse)
library(knitr)
#library(kableExtra)
library(flextable)
library(ggrepel)
library(RODBC)
#library(ggQC)
options(OutDec = ",")
set_flextable_defaults(decimal.mark = ",",big.mark = ".",font.size=9)

#Pegar os dados direto do access mensais
chan1 <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=F:/Documents/PLS.accdb")
PLS_Totais <- sqlFetch(chan1,"cons_todos_m")
#Anuais relativos água e esgoto
PLS_Anuais_A_E <- sqlFetch(chan1,"cons_agua_ener_a")
PLS_veic <- sqlFetch(chan1,"cons_veic_a")

# #Usando a máquina de Araraquara mensais (Desktop)
# chan1 <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=//Notesamsung/Users/OTN/Documents/PLS.accdb")
# PLS_Totais <- sqlFetch(chan1,"cons_todos_m")
# PLS_Anuais_A_E <- sqlFetch(chan1,"cons_agua_ener_a")

close(chan1)

#Criar os dados que interessam
d <- PLS_Totais %>% filter(ano %in% c("2016","2017","2018","2019","2020","2021"), mes %in% 1:12) %>%
    mutate(mes = case_when(
      mes == 1 ~"Jan",
      mes == 2 ~"Fev",
      mes == 3 ~"Mar",
      mes == 4 ~"Abr",
      mes == 5 ~"Mai",
      mes == 6 ~"Jun",
      mes == 7 ~"Jul",
      mes == 8 ~"Ago",
      mes == 9 ~"Set",
      mes == 10 ~"Out",
      mes == 11 ~"Nov",
      mes == 12 ~"Dez"
  )
  )
#colocar meses em ordem
d$mes <- factor(d$mes,levels = c("Jan", "Fev", "Mar", "Abr", "Mai","Jun"
                                 ,"Jul","Ago","Set","Out","Nov","Dez"))
d$ano<- factor(d$ano,levels = c("2016","2017","2018","2019","2020","2021"))

#Cria filtro para dados anuais
d_anual <- PLS_Anuais_A_E %>% filter(ano %in% c("2016","2017","2018","2019","2020","2021"))
d_anual$ano<- factor(d_anual$ano,levels = c("2016","2017","2018","2019","2020","2021"))

#Cria um ID único mes/ano e põe ordem

d <- d %>% mutate(`Mes/Ano` = paste(mes, ano, sep = "/"))
d$`Mes/Ano` <- factor(d$`Mes/Ano`,levels=paste(d$mes, d$ano, sep = "/"), 
                      ordered = TRUE)
PLS_veic$ano=as.factor(PLS_veic$ano)

asspe_colors <- c("#284194","#2191fb","#023618","#61988E","#ee7646","#820900")

rm(PLS_Anuais_A_E,PLS_Totais)

#Função para calcular a variação entre 2 anos
# d: banco de dados
# indic: nome indicador
# Soma: se Falso, calcula a média

FunVar <- function(d,indic, Soma = TRUE)
{
  indic = enquo(indic)
  if(Soma == TRUE)
  {
    IND <- d %>% filter(ano %in% c("2019","2021")) %>%  
    group_by(ano) %>% summarise('IND' = sum(!!indic)) %>% pull(IND)
  }
  else
  {
    IND <- d %>% filter( ano %in% c("2019","2021")) %>%  
      group_by(ano) %>% summarise('IND' = mean(!!indic)) %>% pull(IND)
  }
  
  IND_var <- ((IND[2]/IND[1])-1)*100
  
  list(Indicador = IND, Ind_Var=IND_var)
}





# Consumo de papel total (CPt) ---------------------------------------------------------------------



#Tabela usando flextable e opção valign (alinhamento vertical) e linha de Total
tabCPp <- d %>% select(ano, mes, cpt) %>% spread(ano,cpt) %>%
  #  mutate_at(vars(mes), funs(as.character(.))) %>%
  bind_rows(summarise(.,mes = "Total",
                      `2016` = sum(`2016`,na.rm = TRUE),
                      `2017` = sum(`2017`,na.rm = TRUE),
                      `2018` = sum(`2018`,na.rm = TRUE),
                      `2019` = sum(`2019`,na.rm = TRUE), 
                      `2020` =sum(`2020`,na.rm = TRUE),
                      `2021` =sum(`2021`,na.rm = TRUE)))
colnames(tabCPp)[1] <- "Mês"
tabCPp %>% mutate_if(is.numeric, format, big.mark=".")  %>% flextable() %>% 
  valign( valign = "center", part = "all") %>% align(align = "center",part = "all") %>%
  set_caption( "Indicador 2.7 - CPt")  %>% 
  bold(bold = TRUE, part = "header") %>% bold(bold = TRUE, i = nrow(tabCPp)) %>% 
  autofit()

#gráfico de 2021
filter(d,ano %in% c("2019","2021")) %>% ggplot(aes(x=mes,y=cpt, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, 
                linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(cpt, big.mark = "."), fill= ano, fontface="bold"),color="white", , size = 3)+
  scale_color_manual(values = asspe_colors[c(3,5)])+
  scale_fill_manual(values = asspe_colors[c(3,5)])+
  ylab("CPt\n(Resmas)") + xlab("Mês")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))


#barplot

d %>% group_by(ano) %>% summarise(cpt = sum(cpt)) %>%  
  ggplot(mapping=aes(x=ano, y=cpt, fill =ano)) + geom_bar(stat='identity')+ theme_bw()+
  geom_label_repel(aes(label=format(cpt, big.mark = "."), fill= ano, fontface="bold"),
                   color="white",vjust=0)+
  scale_color_manual(values = asspe_colors)+
  scale_fill_manual(values = asspe_colors)+ylab("CPt\n(Resmas)") + xlab("Ano")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))

## variação 2019-2021

CPt <- FunVar(d=d,indic = cpt)


# Gasto de papel próprio (GPp) ---------------------------------------------------------------------

#Tabela usando flextable e opção valign (alinhamento vertical) elinha de Total

tabGPp <- d %>% select(ano, mes, gpp) %>% spread(ano,gpp) %>%
  bind_rows(summarise(.,mes = "Total",
                      `2016` = sum(`2016`,na.rm = TRUE),
                      `2017` = sum(`2017`,na.rm = TRUE),
                      `2018` = sum(`2018`,na.rm = TRUE),
                      `2019` = sum(`2019`,na.rm = TRUE), 
                      `2020` = sum(`2020`,na.rm = TRUE),
                      `2021` = sum(`2021`,na.rm = TRUE)))%>% rename(Mês = `mes`) 
tabGPp %>% mutate_if(is.numeric, format, big.mark=".",nsmall = 2) %>% flextable() %>% 
  valign( valign = "center", part = "all") %>%  
  set_caption( "Indicador 2.10 - $GP_p$") %>% align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% bold(bold = TRUE, i = nrow(tabGPp)) %>%
  autofit()

#gráfico de 19/21
d %>% filter(ano %in% c("2019","2021")) %>% ggplot(aes(x=mes,y=gpp, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, 
                linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(gpp, big.mark = ".", nsmall = 2, digits = 2), fill= ano, fontface="bold"),
                   color="white",size = 3)+
  scale_color_manual(values = asspe_colors[c(3,5)])+
  scale_fill_manual(values = asspe_colors[c(3,5)])+
  ylab("GPp\n(R$)") + xlab("Mês")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))


#barplot

d %>% group_by(ano) %>% summarise(gpp = sum(gpp)) %>%  
  ggplot(mapping=aes(x=ano, y=gpp, fill =ano)) + geom_bar(stat='identity')+ theme_bw()+
  geom_label_repel(aes(label=format(gpp, big.mark = ".",nsmall = 2, digits = 2), fill= ano, fontface="bold"),
                   color="white",vjust=0,size=3)+
  scale_color_manual(values = asspe_colors)+
  scale_fill_manual(values = asspe_colors)+ylab("GPp\n(R$)") + xlab("Ano")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))

## variação 2019-2021

GPp <- FunVar(d=d,indic = gpp)

# Gastos Relativos Telefonia fixa - GRtf ----------------------------------------

tabGRTf <- d %>% select(ano, mes,grtf) %>% spread(ano,grtf) %>% 
  rename(Mês = `mes`) %>% flextable() %>% colformat_double(j=c(2:7),big.mark=".",decimal.mark = ",",digits=2)  %>% 
  set_caption( "Indicador 6.3 - $GRT_f$") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% 
  valign( valign = "center", part = "all") %>% 
  autofit()
tabGRTf

#graf 19/21

d %>% filter(ano %in% c("2019","2021")) %>% ggplot(aes(x=mes,y=grtf, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(grtf, big.mark = ".",digits = 2, nsmall = 2), 
                       fill= ano, fontface="bold"),color="white")+
  scale_color_manual(values = asspe_colors[c(1,5)])+
  scale_fill_manual(values = asspe_colors[c(1,5)])+
  ylab(expression(atop(~GRT[f],"(R$ \\ linha fixa)"))) + xlab("Mês")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))


#barplot

d %>% group_by(ano) %>% summarise(grtf = mean(grtf)) %>%  
  ggplot(mapping=aes(x=ano, y=grtf, fill =ano)) + geom_bar(stat='identity')+ theme_bw()+
  geom_label_repel(aes(label=format(grtf, big.mark = ".",digits = 2, nsmall = 2), fill= ano, fontface="bold"),
                   color="white",vjust=0)+
  scale_color_manual(values = asspe_colors)+
  scale_fill_manual(values = asspe_colors)+ylab(
    expression(atop("Gasto médio anual"~GRT[f],"(R$ \\ linha fixa)"))) + xlab("Ano")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))

# variação 2019-2021 -  média

GRTf <- FunVar(d=d,indic = grtf, Soma = FALSE)

# Gastos Relativos Telefonia Móvel - GRTm ---------------------------------

#Tabela
tabGRTm <- d %>% select(ano, mes,grtm) %>% spread(ano,grtm) %>% 
  rename(Mês = `mes`) %>% flextable() %>% 
  colformat_double(j=c(2:7),big.mark=".",decimal.mark = ",",digits=2)  %>% 
  set_caption( "Indicador 6.3 - $GRT_m$") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% 
  valign( valign = "center", part = "all") %>% 
  autofit()
tabGRTm

#graf 19/21
filter(d,ano %in% c("2019","2021")) %>% ggplot(aes(x=mes,y=grtm, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, 
                linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(grtm, big.mark = ".",digits = 2, nsmall = 2), 
                       fill= ano, fontface="bold"),color="white")+
  scale_color_manual(values = asspe_colors[c(1,5)])+
  scale_fill_manual(values = asspe_colors[c(1,5)])+
  ylab(expression(atop(~GRT[m],"(R$ \\ linha móvel)"))) + xlab("Mês")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",digits = 2, nsmall = 2, 
                                                 scientific = FALSE))
#barplot

d %>% group_by(ano) %>% summarise(grtm = mean(grtm)) %>%  
  ggplot(mapping=aes(x=ano, y=grtm, fill =ano)) + geom_bar(stat='identity')+ theme_bw()+
  geom_label_repel(aes(label=format(grtm, big.mark = ".",digits = 2, nsmall = 2), fill= ano, fontface="bold"),
                   color="white",vjust=0)+
  scale_color_manual(values = asspe_colors)+
  scale_fill_manual(values = asspe_colors)+ylab(
    expression(atop("Gasto médio anual"~GRT[m],"(R$ \\ linha móvel)"))) + xlab("Ano")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))

# variação 2019 -2021- média

GRTm <- FunVar(d=d,indic = grtm, Soma = FALSE)

# Consumo de Energia Elétrica - CE ----------------------------------------

#Tabela CE
tabCE <- d %>% select(ano, mes, cee) %>% spread(ano,cee) %>%
  #  mutate_at(vars(mes), funs(as.character(.))) %>%
  bind_rows(summarise(.,mes = "Total",
                      `2016` = sum(`2016`,na.rm = TRUE),
                      `2017` = sum(`2017`,na.rm = TRUE),
                      `2018` = sum(`2018`,na.rm = TRUE),
                      `2019` = sum(`2019`,na.rm = TRUE), 
                      `2020` =sum(`2020`,na.rm = TRUE),
                      `2021` =sum(`2021`,na.rm = TRUE))) %>% 
rename(Mês = `mes`)
tabCE <- tabCE %>% flextable() %>% 
  colformat_double(j=c(2:7),big.mark=".",decimal.mark = ",",digits = 1)  %>% 
  set_caption( "Indicador 7.1 - CE (KWh)") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% 
  valign( valign = "center", part = "all") %>% 
  bold(bold = TRUE, i = nrow(tabCE)) %>% 
  autofit()
tabCE

#graf 19/21
filter(d,ano %in% c("2019","2021")) %>% ggplot(aes(x=mes,y=cee, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(cee, big.mark = "."),fill= ano, fontface="bold"), color="white")+
  scale_color_manual(values = asspe_colors[c(1,4)])+
  scale_fill_manual(values = asspe_colors[c(1,4)])+
  #Aqui retira o formato de número científico
  scale_y_continuous(name="CE\n(Kwh)", labels = function(x) format(x, big.mark = ".", scientific = FALSE))+ 
  xlab("Mês")

#barplot Soma CE

d %>% group_by(ano) %>% summarise(cee = sum(cee)) %>%  
  ggplot(mapping=aes(x=ano, y=cee, fill =ano)) + geom_bar(stat='identity')+ theme_bw()+
  geom_label_repel(aes(label=format(cee, big.mark = ".",digits = 2, nsmall = 1),
                       fill= ano, fontface="bold"),
                   color = "white", vjust=0)+
  scale_color_manual(values = asspe_colors)+
  scale_fill_manual(values = asspe_colors)+ylab(
    expression(atop("Consumo bruto anual"~CE,"(KWh)"))) + xlab("Ano")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))

# variação 2019-2021
CEE <- FunVar(d=d,indic = cee)

## Consumo Relativo de EE (CRE)

#Tabela cRE

tabm2Tot <- d_anual %>% select(ano,m2Total,cre) %>% 
  rename(Ano = `ano`) %>% rename(`m2 Total` = m2Total) %>% 
  rename(CRE = `cre`) %>% 
  flextable() %>%  
  colformat_double(j=c(2),big.mark= ".", digits = 0)  %>% 
  colformat_double(j=c(3),big.mark= ".", digits = 2)  %>% 
  set_caption( "Área total das Unidades do TRE-SP e indicador 7.2 -  CRE") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% valign( valign = "center", part = "all") %>% 
  autofit()
tabm2Tot

#Evolução da área construída TRE
ggplot(d_anual, aes(x=ano,y= m2Total, fill=ano))+ theme_bw()+
  geom_bar(stat='identity')+ 
 # geom_point(shape=21, size=5,color=asspe_colors[1], fill=asspe_colors[1])+
  geom_label_repel(aes(label=format(m2Total, big.mark = "."),
                       fill= ano, fontface="bold"), 
                   color="white")+
  scale_color_manual(values = asspe_colors)+
  scale_fill_manual(values = asspe_colors)+
  #Aqui retira o formato de número científico
  scale_y_continuous(name=expression(~m^{2}~Total), 
                     labels = function(x) format(x, big.mark = ".", scientific = FALSE))+ 
  xlab("Ano")

# GRáfico CRE

d_anual %>%  ggplot(aes(x=ano,y=cre, group = 1))+ theme_bw()+
  geom_line(size = 1.1,color=asspe_colors[3])+ 
  geom_point(fill = asspe_colors[3],shape=21, size=5)+
  geom_label_repel(aes(label=format(cre, big.mark = ".",nsmall=3, digits=3)),
                   fill= asspe_colors[3], fontface="bold", 
                   color="white")+

  #Aqui retira o formato de número científico
  scale_y_continuous(name=
                       expression(atop("CRE",~KWh/m^{2}~Total)), 
                     labels = function(x) format(x, big.mark = ".", scientific = FALSE),
                     limits = c(0,35))+ 
  xlab("Ano")

#Comparação

#Media
CRE <- d_anual %>% select(cre) %>% pull()

#variação
CRE_var <- ((CRE[6]/CRE[4])-1)*100

# Gasto de Energia Elétrica - GE ----------------------------------------


#Tabela GE
tabGE <- d %>% select(ano, mes, ge) %>% spread(ano,ge) %>%
  #  mutate_at(vars(mes), funs(as.character(.))) %>%
  bind_rows(summarise(.,mes = "Total",
                      `2016` = sum(`2016`,na.rm = TRUE),
                      `2017` = sum(`2017`,na.rm = TRUE),
                      `2018` = sum(`2018`,na.rm = TRUE),
                      `2019` = sum(`2019`,na.rm = TRUE), 
                      `2020` =sum(`2020`,na.rm = TRUE),
                      `2021` =sum(`2021`,na.rm = TRUE))) %>% 
  rename(Mês = `mes`)
tabGE <- tabGE %>% flextable() %>% 
  colformat_double(j=c(2:7),big.mark=".",digits=2)  %>% 
  set_caption( "Indicador 7.3 - GE (R$)") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% valign( valign = "center", part = "all") %>% 
  bold(bold = TRUE, i = nrow(tabGE)) %>%  autofit()
tabGE

#graf 19/21
filter(d,ano %in% c("2019","2021")) %>% ggplot(aes(x=mes,y=ge, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(ge, big.mark = ".",digits=2,nsmall=2),fill= ano, fontface="bold"), 
                   color="white",size=3)+
  scale_color_manual(values = asspe_colors[c(1,4)])+
  scale_fill_manual(values = asspe_colors[c(1,4)])+
  #Aqui retira o formato de número científico
  scale_y_continuous(name="GE\n(R$)", labels = function(x) format(x, big.mark = ".", scientific = FALSE))+ 
  xlab("Mês")

#barplot Soma GE

d %>% group_by(ano) %>% summarise(ge = sum(ge)) %>%  
  ggplot(mapping=aes(x=ano, y=ge, fill =ano)) + geom_bar(stat='identity')+ theme_bw()+
  geom_label(aes(label=format(ge, big.mark = ".",digits = 2, nsmall = 2), fill= ano, fontface="bold"),
             color="white",vjust=2)+
  scale_color_manual(values = asspe_colors)+
  scale_fill_manual(values = asspe_colors)+ylab(
    expression(atop("Gasto bruto anual"~GE,"(R$)"))) + xlab("Ano")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))
# variação 2019-2021 
GE <- FunVar(d=d,indic = ge)

## Gasto Relativo de EE (GRE)

#tabela
tabgre <- d_anual %>% select(ano,m2Total,gre) %>% 
  rename(Ano = `ano`) %>% rename(`m2 Total` = m2Total) %>% 
  rename(`GRE` = gre) %>% 
  flextable() %>% 
  colformat_double(j=c(2),big.mark=".",digits=0)  %>% 
  colformat_double(j=c(3),big.mark=".",digits=2)  %>% 
  set_caption( "Área total das Unidades do TRE-SP e indicador 7.4 -  GRE") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% valign( valign = "center", part = "all") %>% 
    autofit()
tabgre


# GRáfico GRE

d_anual %>%  ggplot(aes(x=ano,y=gre, group = 1))+ theme_bw()+
  geom_line(size = 1.1,color=asspe_colors[3])+ 
  geom_point(fill = asspe_colors[3],shape=21, size=5)+
  geom_label_repel(aes(label=format(gre, big.mark = ".",nsmall=3, digits=3)),
                   fill= asspe_colors[3], fontface="bold", 
                   color="white")+
  
  #Aqui retira o formato de número científico
  scale_y_continuous(name=
                       expression(atop("GRE","R$"/m^{2}~Total)), 
                     labels = function(x) format(x, big.mark = ".", scientific = FALSE),
                    limits=c(0,21) )+ 
  xlab("Ano")


# Consumo de ÁGUA - CA ----------------------------------------

#Tabela CA
tabCA <- d %>% select(ano, mes, ca) %>% spread(ano,ca) %>%
  #  mutate_at(vars(mes), funs(as.character(.))) %>%
  bind_rows(summarise(.,mes = "Total",
                      `2016` = sum(`2016`,na.rm = TRUE),
                      `2017` = sum(`2017`,na.rm = TRUE),
                      `2018` = sum(`2018`,na.rm = TRUE),
                      `2019` = sum(`2019`,na.rm = TRUE), 
                      `2020` =sum(`2020`,na.rm = TRUE),
                      `2021` =sum(`2021`,na.rm = TRUE))) %>% 
  rename(Mês = `mes`)
tabCA <- tabCA %>% flextable() %>% 
  colformat_num(j=c(2:7),big.mark=".")  %>% 
  set_caption( "Indicador 8.1 - CA (m3)") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% valign( valign = "center", part = "all") %>% 
  bold(bold = TRUE, i = nrow(tabCA)) %>%  autofit()
tabCA

#graf 19/21
filter(d,ano %in% c("2019","2021")) %>% ggplot(aes(x=mes,y=ca, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(ca, big.mark = "."),fill= ano, fontface="bold"), color="white")+
  scale_color_manual(values = asspe_colors[c(5,2)])+
  scale_fill_manual(values = asspe_colors[c(5,2)])+
  #Aqui retira o formato de número científico
  scale_y_continuous(name=expression("CA"~(m^{3})), labels = function(x) format(x, big.mark = ".", scientific = FALSE))+ 
  xlab("Mês")

#barplot Soma Ca

d %>% group_by(ano) %>% summarise(ca = sum(ca)) %>%  
  ggplot(mapping=aes(x=ano, y=ca, fill =ano)) + geom_bar(stat='identity')+ theme_bw()+
  geom_label_repel(aes(label=format(ca, big.mark = ".",digits = 2, nsmall = 1), fill= ano, fontface="bold"),
                   color="white",vjust=0)+
  scale_color_manual(values = asspe_colors)+
  scale_fill_manual(values = asspe_colors)+ylab(
    expression(atop("Consumo bruto anual"~CA,(m^{3})))) + xlab("Ano")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))
# variação 2019-2021 
CA <- FunVar(d=d,indic = ca)

## Consumo Relativo de água (CRA)

#Tabela cRA

tabcra <- d_anual %>% select(ano,m2Total,cra) %>% 
  rename(Ano = `ano`) %>% rename(`m2 Total` = m2Total) %>% 
  rename(CRA = `cra`) %>% 
  flextable() %>%  
  colformat_double(j=c(2),big.mark=".",digits=0)  %>% 
  colformat_double(j=c(3),big.mark=".",digits=2)  %>% 
  set_caption( "Área total das Unidades do TRE-SP e indicador 8.2 -  CRA") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% valign( valign = "center", part = "all") %>% 
  autofit()
tabcra



# GRáfico CRA

d_anual %>%  ggplot(aes(x=ano,y=cra, group = 1))+ theme_bw()+
  geom_line(size = 1.1,color=asspe_colors[2])+ 
  geom_point(fill = asspe_colors[2],shape=21, size=5)+
  geom_label_repel(aes(label=format(cra, big.mark = ".",nsmall=3, digits=3)),
                   fill= asspe_colors[2], fontface="bold", 
                   color="white", size=3.75)+
  
  #Aqui retira o formato de número científico
  scale_y_continuous(name=
                       expression(atop("CRA",~m^{3}/m^{2}~Total)), 
                     labels = function(x) format(x, big.mark = ".", scientific = FALSE),
                     limits = c(0,.3))+ 
  xlab("Ano")


# Gasto com água - GA ----------------------------------------


#Tabela GA
tabGA <- d %>% select(ano, mes, ga) %>% spread(ano,ga) %>%
  #  mutate_at(vars(mes), funs(as.character(.))) %>%
  bind_rows(summarise(.,mes = "Total",
                      `2016` = sum(`2016`,na.rm = TRUE),
                      `2017` = sum(`2017`,na.rm = TRUE),
                      `2018` = sum(`2018`,na.rm = TRUE),
                      `2019` = sum(`2019`,na.rm = TRUE), 
                      `2020` =sum(`2020`,na.rm = TRUE),
                      `2021` =sum(`2021`,na.rm = TRUE))) %>% 
  rename(Mês = `mes`)
tabGA <- tabGA %>% flextable() %>% 
  colformat_double(j=c(2:7),big.mark=".",digits=2)  %>% 
  set_caption( "Indicador 8.3 - GA (R$)") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% valign( valign = "center", part = "all") %>% 
  bold(bold = TRUE, i = nrow(tabGA)) %>%  autofit()
tabGA

#graf 19/21
filter(d,ano %in% c("2019","2021")) %>% ggplot(aes(x=mes,y=ga, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(ga, big.mark = ".",digits=2,nsmall=2),fill= ano, fontface="bold"), 
                   color="white",size=3)+
  scale_color_manual(values = asspe_colors[c(5,2)])+
  scale_fill_manual(values = asspe_colors[c(5,2)])+
  #Aqui retira o formato de número científico
  scale_y_continuous(name="GA\n(R$)", labels = function(x) format(x, big.mark = ".", scientific = FALSE))+ 
  xlab("Mês")

#barplot Soma GA

d %>% group_by(ano) %>% summarise(ga = sum(ga)) %>%  
  ggplot(mapping=aes(x=ano, y=ga, fill =ano)) + geom_bar(stat='identity')+ theme_bw()+
  geom_label(aes(label=format(ga, big.mark = ".",digits = 2, nsmall = 2), fill= ano, fontface="bold"),
                   color="white",vjust=2,size=3.5)+
  scale_color_manual(values = asspe_colors)+
  scale_fill_manual(values = asspe_colors)+ylab(
    expression(atop("Consumo bruto anual"~GA,"(R$)"))) + xlab("Ano")+
  scale_y_continuous(labels = function(x) format(x, big.mark = ".",
                                                 scientific = FALSE))

# variação 2019-2021 
GA <- FunVar(d=d,indic = ga)

## Gasto Relativo com água (GRA)

#tabela
tabgra <- d_anual %>% select(ano,m2Total,gra) %>% 
  rename(Ano = `ano`) %>% rename(`m2 Total` = m2Total) %>% 
  rename(`GRA` = gra) %>% 
  flextable() %>% 
  colformat_double(j=c(2),big.mark=".",digits=0)  %>% 
  colformat_double(j=c(3),big.mark=".",digits=2)  %>% 
  set_caption( "Área total das Unidades do TRE-SP e indicador 8.4 -  GRA") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% valign( valign = "center", part = "all") %>% 
  autofit()
tabgra


# GRáfico GRA

d_anual %>%  ggplot(aes(x=ano,y=gra, group = 1))+ theme_bw()+
  geom_line(size = 1.1,color=asspe_colors[2])+ 
  geom_point(fill = asspe_colors[2],shape=21, size=5)+
  geom_label_repel(aes(label=format(gra, big.mark = ".",nsmall=2, digits=2)),
                   fill= asspe_colors[2], fontface="bold", 
                   color="white",size=3.8)+
  
  #Aqui retira o formato de número científico
  scale_y_continuous(name=
                       expression(atop("GRA","R$"/m^{2}~Total)), 
                     labels = function(x) format(x, big.mark = ".", scientific = FALSE),
                     limits=c(0,6) )+ 
  xlab("Ano")

# Total de material para reciclagem ---------------------------------------

#Modelo 4

#Tabela geral com os indicadores que formam a TMR

tabTMR <- d %>% filter(ano %in% c("2019","2021")) %>% 
  select(mes,ano,dpa, dpl, dmt, dvd,cge,tmr) %>% 
  mutate(Mês = paste(mes,"/",ano,sep="")) %>% select(-mes,-ano) %>% 
  relocate(Mês,.before = dpa) %>% rename("Dpa" = `dpa`,"Dpl"=`dpl`,"Dmt"=`dmt`,
                                      "Dvd"=`dvd`,"Cge"=`cge`,"TMR"=`tmr`) %>% 
  flextable() %>% 
  colformat_double(j=c(2:7),big.mark=".", 
                decimal.mark = ",",digits=0)  %>% 
  set_caption( "Indicadores componentes\n9.6 - Total de material para reciclagem \n(Kg)") %>% 
  align(j=1, align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% 
  valign(valign = "center", part = "all") %>%  autofit()
tabTMR


#Tabela com a soma dos indicadores componentes do TMR

tabCompTMR <- d %>% filter(ano %in% 2016:2021) %>%  
  select(ano,dpa, dpl, dmt, dvd,cge, tmr) %>% group_by(ano) %>%
  summarise_all(.funs = sum) %>% 
  rename("Dpa" = `dpa`,"Dpl"=`dpl`,"Dmt"=`dmt`,
         "Dvd"=`dvd`,"Cge"=`cge`, "TMR" = `tmr`) %>% 
  flextable() %>% 
  colformat_double(j=c(2:7),big.mark=".",digits=0)  %>% 
  set_caption( "Indicadores de coletas de resíduos - Totais Anuais (kg)") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% bold(j = c(7),bold = TRUE) %>% 
  valign(valign = "center", part = "all") %>%  autofit()
tabCompTMR


#Gráfico do TMR

d %>% filter(ano %in% c("2019","2021")) %>% ggplot(aes(x=mes,y=tmr, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(tmr, big.mark = ".", digits = 0),
                       fill= ano, fontface="bold"), color="white",size=3.75)+
  scale_color_manual(values = asspe_colors[c(2,3)])+
  scale_fill_manual(values = asspe_colors[c(2,3)])+
  #Aqui retira o formato de número científico
  scale_y_continuous(name= "TMR\n(Kg)", 
                     labels = function(x) format(x, big.mark = ".", scientific = FALSE))+ 
  xlab("Mês")


# variação 2019-2021 
TMR <- FunVar(d=d,indic = tmr)



# Destinação de Resíduos de saúde ---------------------------------------



#Tabela
tabDrs <- d %>% select(ano, mes, drs) %>% spread(ano,drs) %>% 
  bind_rows(summarise(.,mes = "Total",
                      `2016` = sum(`2016`,na.rm = TRUE),
                      `2017` = sum(`2017`,na.rm = TRUE),
                      `2018` = sum(`2018`,na.rm = TRUE),
                      `2019` = sum(`2019`,na.rm = TRUE), 
                      `2020` =sum(`2020`,na.rm = TRUE),
                      `2021` =sum(`2021`,na.rm = TRUE))) %>% rename(Mês = `mes`)
tabDrs <- tabDrs %>% flextable() %>% 
  colformat_double(j=c(2:7),big.mark=".",digits=1)  %>% 
  set_caption( "Indicador 9.11 - Destinação de resíduos de saúde - Drs (Kg)") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% 
  valign( valign = "center", part = "all") %>%  
  bold(bold = TRUE, i  = nrow(tabDrs)) %>% 
  autofit()
tabDrs


#Gráfico do DRS

d %>% filter(ano %in% c("2019","2021")) %>% 
  ggplot(aes(x=mes,y=drs, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(drs, big.mark = ".", digits = 1),
                       fill= ano, fontface="bold"), 
                   color="white",size=3.75)+
  scale_color_manual(values = asspe_colors[c(2,3)])+
  scale_fill_manual(values = asspe_colors[c(2,3)])+
  #Aqui retira o formato de número científico
  scale_y_continuous(name= expression(~D[rs]~"(Kg)"), 
                     labels = function(x) format(x, big.mark = ".", scientific = FALSE))+ 
  xlab("Mês")

# variação 2019-2021 
DRS <- FunVar(d=d,indic = drs)


# Consumo relativo de álcool e gasolina  -----------------------------------------------


#Tabela dos Carros menos diesel


tabveic <- PLS_veic %>% filter(ano %in% c("2019","2021")) %>% select(ano,vg,vet,vf,vh) %>% 
  rename(Ano = `ano`, VG = `vg`, VEt = `vet`, VF = `vf`, VH = `vh`)
tabveic <- tabveic %>% flextable() %>% 
  colformat_num(j=c(2:5),big.mark=".")  %>% 
  set_caption( "Quantidade de veículos a álcool e/ou gasolina - TRE/SP") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% 
  valign( valign = "center", part = "all") %>%  
  autofit()
tabveic

#Tabela usando CRag
tabCrag <- d %>% select(ano, mes, crag) %>% 
    spread(ano,crag) %>%
    rename(Mês = `mes`)%>% 
    flextable() %>% 
    colformat_double(j=c(2:7),big.mark=".",digits = 1) %>% 
    set_caption( "Indicador 14.5 - Consumo relativo de álcool e gasolina - CRag (l/veículo)") %>% 
    align(align = "center",part = "all") %>% 
    bold(bold = TRUE, part = "header") %>% 
    valign( valign = "center", part = "all") %>%  
    autofit()
tabCrag


#graf 19/21

d %>% filter(ano %in% c("2019","2021")) %>% 
  ggplot(aes(x=mes,y=crag, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(crag, big.mark = ".", digits = 2),fill= ano, fontface="bold"), color="white")+
  scale_color_manual(values = asspe_colors[c(2,5)])+
  scale_fill_manual(values = asspe_colors[c(2,5)])+
  #Aqui retira o formato de número científico
  scale_y_continuous(name= expression(~ CR[ag] ~"\n(l/veículo)"), 
                     labels = function(x) format(x, big.mark = ".", scientific = FALSE))+ 
  xlab("Mês")

# variação 2019-2021 
CRAG <- FunVar(d=d,indic = crag,Soma = FALSE)



# Consumo relativo de diesel  -----------------------------------------------


#Tabela dos veículos a diesel

tabveicd <- PLS_veic %>% filter(ano %in% c("2019","2021")) %>% select(ano,vd) %>% 
  rename(Ano = `ano`, VD = `vd`)
tabveicd <- tabveicd %>% flextable() %>% 
  colformat_num(j=c(2),big.mark=".")  %>% 
  set_caption( "Quantidade de veículos a diesel - TRE/SP") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% 
  valign( valign = "center", part = "all") %>%  
  autofit()
tabveicd


#Tabela usando flextable
tabCRd <- d %>% select(ano, mes, crd) %>% 
  spread(ano,crd) %>% 
  rename(Mês = `mes`)%>% 
  flextable() %>% 
  colformat_double(j=c(2:6),big.mark=".",digits=1)  %>% 
  set_caption( "Indicador 14.5 - Consumo relativo de diesel- CRd (l/veículo)") %>% 
  align(align = "center",part = "all") %>% 
  bold(bold = TRUE, part = "header") %>% 
  valign( valign = "center", part = "all") %>%  
  autofit()
tabCRd


#Gráfico

d %>% filter(ano %in% c("2019","2021")) %>% 
  ggplot(aes(x=mes,y=crd, group = ano))+ theme_bw()+
  geom_line(aes(color=ano, linetype = ano),size = 1.1)+ 
  geom_point(aes(fill = ano),shape=21, size=3)+
  geom_label_repel(aes(label=format(crd, big.mark = ".", digits = 2, nsmall = 1),fill= ano, fontface="bold"), color="white")+
  scale_color_manual(values = asspe_colors[c(2,5)])+
  scale_fill_manual(values = asspe_colors[c(2,5)])+
  #Aqui retira o formato de número científico
  scale_y_continuous(name= expression(~ CR[d] ~"\n(l/veículo)"), 
                     labels = function(x) format(x, big.mark = ".", scientific = FALSE))+ 
  xlab("Mês")

# variação 2019-2021 
CRD <- FunVar(d=d,indic = crd,Soma = FALSE)

