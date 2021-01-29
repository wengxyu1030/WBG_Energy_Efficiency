.libPaths("C:/Users/wb500886/Downloads/library")
library(readxl)
library(openxlsx)
library(tidyverse)
library(readxl)
library(countrycode)
library(skimr)
library(dplyr)
library(xlsx)
library(lubridate)
library(imputeTS)

#Universe-----
operation_lend_raw <- read_excel("input/LENDING_PROGRAM_V2.xlsx")

operation_raw <- read_excel("input/PROJECT_MASTER_V2.xlsx", col_types = "text")

operation_0 <- read_excel("input/LENDING_PROGRAM_V2.xlsx") %>%
  select(PROJ_ID,TOT_CMT_USD_AMT) %>%
  dplyr::mutate(inst = "WB-Lend")

save(operation_0,file = "data/operation_0.rda")

theme_project <- read_excel("input/PROJECT_THEME_V2.xlsx", skip = 4)
theme <- read_excel("input/THEME.xlsx",skip = 4)

sct <- read_excel("input/SECTOR.xlsx",skip = 4)
sct_project <- read_excel("input/PROJECT_SECTOR_V2.xlsx", skip = 4)
sct_project_3 <- sct_project %>% 
  distinct(PROJ_ID,SECT_TEXT,SECT_PERCENTAGE) %>%
  group_by(PROJ_ID) %>%
  top_n(3,SECT_PERCENTAGE)%>%
  mutate(rank = order(SECT_PERCENTAGE))%>%
  filter(rank <= 3) %>% #select by alphabetic if all the same share of sector. 
  pivot_wider(-SECT_PERCENTAGE,names_from = rank, names_prefix = "sector_",values_from = SECT_TEXT) %>%
  rename(pid = PROJ_ID)


asa <- read_excel("input/ASA_PROJECT.xlsx")%>%
  mutate(inst = "WB-ASA") %>%
  dplyr::rename(pid = PROJ_ID, addfinancing = ADDTNL_FNCNG_IND, 
                ttl_upi =  TEAM_LEAD_UPI,
                amount = CMT_AMT,
                practice = LEAD_GP_NAME,ctr = CNTRY_SHORT_NAME, appfy = PROJ_APPRVL_FY) %>%
  mutate(amount_asa = as.numeric(amount),
         ttl_upi = as.character(ttl_upi),
         appfy = as.numeric(appfy)) %>%
  select(pid,amount_asa,inst) 

skim(asa) #no commitment from this data source. 

#wb ctr cmt-----
ctr_ptf <- operation_lend_raw %>%
  mutate(amount = as.numeric(TOT_CMT_USD_AMT),
         APPRVL_FY = as.numeric(APPRVL_FY))%>%
  filter(dplyr::between(APPRVL_FY,2011,2020))%>%
  group_by(CNTRY_SHORT_NAME) %>%
  dplyr::summarise(amount_wb = sum(amount, rm.na = TRUE),
                   n_wb = n())%>% 
  mutate(ctr_code = countrycode(CNTRY_SHORT_NAME, 
                                origin = 'country.name',
                                destination = 'iso3c',
                                custom_match = c('Kosovo'='KVO',
                                                 "Yemen, People's Democratic Republic of" = 'YEM'))) %>%
  distinct(ctr_code,amount_wb,n_wb)
save(ctr_ptf,file = "data/ctr_ptf.rda")

#fcv countries----
fcv <- operation_raw %>%
  distinct(FRAGILE_STATE_IND,CNTRY_SHORT_NAME,PROJ_APPRVL_FY) %>%
  mutate(PROJ_APPRVL_FY = round(as.numeric(PROJ_APPRVL_FY),digits = 0),
         PROJ_APPRVL_FY = as.character(PROJ_APPRVL_FY))

save(fcv,file = "data/fcv.rda")

#country_tag----
path_raw <- "C:/Users/wb500886/OneDrive - WBG/9_Decentralization/raw/R/Decentralization"
load(paste0(path_raw,"/data/income_ts.rda"))

ctr_tag <- read_excel("input/LENDING_PROGRAM_V2.xlsx", col_types = "text") %>%
  #import FCV, Income Group, Lending Group,country tag by fy.
  distinct(APPRVL_FY,FRAGILE_STATES_CODE,CNTRY_CODE,RGN_CODE,LNDNG_GRP_CODE,CNTRY_SHORT_NAME) %>%
  dplyr::rename( fy = APPRVL_FY,
                 ctr = CNTRY_SHORT_NAME,
                 rgn_proj = RGN_CODE) %>%
  mutate(fy = round(as.numeric(fy),0), #though it's appfy, it's not about approval but just FY ctr tag.
         fcv = ifelse(FRAGILE_STATES_CODE == "YES", 1, 0),
         iso3c = countrycode(ctr, 
                             origin = 'country.name',
                             destination = 'iso3c',
                             custom_match = c('Kosovo'='KVO',
                                              'Micronesia'= 'FSM',
                                              'Yugoslavia'= 'SRB',
                                              'Serbia and Montenegro' = 'SRB'))) %>%  #the converged updated iso3c list
  select(iso3c,fy,fcv,ctr) %>%
  filter(!is.na(iso3c),fy>1900) %>%
  
  #impute fcv
  pivot_wider( -ctr,names_from = iso3c, values_from = fcv)%>%
  arrange(fy) %>%
  as.ts() %>% 
  na_interpolation(option = "linear")%>%
  as.data.frame() %>%
  pivot_longer(-fy,names_to = "iso3c", values_to = "fcv") %>%
  
  #add income_n
  left_join(income_ts,by = c("fy","iso3c")) %>% #do not full join because only focus on Bank lending clients
  
  #clean data 
  group_by(iso3c)%>%
  mutate(income_n = ifelse(is.na(income_n),mean(income_n,na.rm = TRUE),income_n),
         fcv = ifelse(is.na(fcv),mean(fcv,na.rm = TRUE),fcv))%>%
  mutate(fcv = ifelse(is.na(fcv),0,fcv)) %>%
  mutate(`Income Level` = case_when( income_n < 1.5 ~ "LIC",
                                     dplyr::between(income_n,1.5,2.5) ~ "LMIC",
                                     dplyr::between(income_n,2.5,3.5) ~ "UMIC",
                                     income_n >= 3.5 ~ "UIC"),
         FCS = ifelse(fcv >0 , "Yes","No")) %>%
  ungroup 



skim(ctr_tag)
save(ctr_tag, file = "data/ctr_tag.rda")

#sector info-----
sector_raw <- read_excel("input/dbo_CDS_PROJECT_SECTOR Query.xlsx")
sector <- sector_raw %>%
  group_by(PROJECTID) %>%
  top_n(SECTOR_PCT,1) %>%
  distinct(PROJECTID,MAJORSECTOR_NAME)
save(sector, file = "data/sector.rda")

#indicator ----
rise_raw <- read_excel("input/Scores.xls",sheet = "Energy Efficiency")

rise <- rise_raw %>%
  cbind(score = rowSums(rise_raw
                        %>% select(where(is.numeric)))
  ) %>% 
  mutate(ctr_code = countrycode(Countries, 
                                origin = 'country.name',
                                destination = 'iso3c',
                                custom_match = c('Kosovo'='KVO',
                                                 "Yemen, People's Democratic Republic of" = 'YEM'))) %>%
  distinct(ctr_code,score)
save(rise,file = "data/rise.rda")

#cleaning up universe----- 

operation <- operation_raw %>%
  select(PROJ_ID,PROJ_SHORT_NAME,PROJ_STAT_NAME,TEAM_LEAD_UPI,PROD_LINE_NAME,INC_GRP_CODE,PREP_COST_AMT,
         CNTRY_SHORT_NAME,PROJ_APPRVL_FY,RGN_ABBR_CODE,LNDNG_INSTR_CODE,PROJ_DEV_OBJECTIVE_DESC,LNDNG_GRP_CODE,LEAD_GP_NAME) %>%
  filter(grepl("p",PROJ_ID,ignore.case = TRUE)) %>%
  left_join(operation_0,by="PROJ_ID") %>%
  mutate(inst = ifelse(is.na(inst),"WB-ASA",inst),
         #lending and asa amount (preparation cost as asa amount)
         amount = ifelse(inst == "WB-Lend",TOT_CMT_USD_AMT,PREP_COST_AMT),
         amount = as.numeric(amount),
         amount = ifelse(amount <0,0,amount))%>%
  dplyr::rename(rgn = RGN_ABBR_CODE,
                pid = PROJ_ID,
                ctr = CNTRY_SHORT_NAME,
                appfy = PROJ_APPRVL_FY,
                income = INC_GRP_CODE,
                lend_instr_wb = LNDNG_INSTR_CODE,
                lend_grp_wb = LNDNG_GRP_CODE,
                practice_wb = LEAD_GP_NAME,
                name = PROJ_SHORT_NAME,
                status = PROJ_STAT_NAME ) %>%
  mutate(iso3c = countrycode(ctr, 
                                      origin = 'country.name',
                                      destination = 'iso3c',
                                      custom_match = c('Kosovo'='KVO',
                                                       "Yemen, People's Democratic Republic of" = 'YEM',
                                                       'Micronesia'= 'FSM',
                                                       'Yugoslavia'= 'SRB',
                                                       'Yugoslavia, former'= 'SRB',
                                                       'Serbia and Montenegro' = 'SRB',
                                                       'Czechoslovakia' = "CZE")),
         appfy = round(as.numeric(appfy),0)) 

operation$inst %>% table()
operation %>% filter(inst == "WB-Lend") %>% skim(amount)
operation %>% filter(inst == "WB-ASA") %>% skim(amount)

save(operation,file = "data/operation.rda")



#EE Portfolio------

var_keep <- c("pid", "name","rgn", "ctr", "appfy", "income", "lend_instr_wb", "lend_grp_wb", "iso3c", "appfy", 
              "inst","amount(m)","practice_wb",
              "sector_1","sector_2","sector_3","status")

#lending and ASA from gp
ee_wb_raw <- read_excel("input/EE Port Review 051320.xlsx", sheet = "EE Portfolio",skip = 2) %>%
  filter(!is.na(`ID#`))

skim(ee_wb_raw %>% select(`ID#`,`Total EE Cost ($m)`))

ee_wb_gp <- ee_wb_raw %>%
  rename(pid = `ID#`,
         amount = `Total EE Cost ($m)` ) %>%
  select(pid, amount) %>%
  group_by(pid)%>%
  dplyr::summarise(amount_gp_ee = sum(amount))

anyDuplicated(ee_wb_gp$pid)
ee_wb_gp %>% skim()

save(ee_wb_gp,file = "data/ee_wb_gp.rda")

#asa from esmap
load("data/operation.rda")
esmap <- read_excel("input/ESMAP.xlsx")

esmap_asa_raw <- esmap %>% 
  select(`P#`,`Commitments USD`,`ESMAP Thematic and Cross-Cutting`,`TF #`,Priority) %>%
  rename(pid = `P#`,
         theme_esmap = `ESMAP Thematic and Cross-Cutting`) %>%
  left_join(operation,by = "pid") %>%
  filter(inst == "WB-ASA",
         !is.na(inst)) %>%
  distinct(pid, PROJ_LGL_NAME, `TF #`,theme_esmap,PROD_LINE_NAME,rgn,ctr,appfy,practice_wb,Priority) 

priority_ee <- c("Efficient City Services","Efficient and Clean Cooling","Efficient and Sustainable Buildings","Lighting Global")

esmap_asa <- esmap_asa_raw %>%
  filter(grepl(paste(priority_ee,collapse = "|"),Priority)) %>%
  distinct(pid)

esmap_asa$pid %>% anyDuplicated()
esmap_asa %>% distinct(pid) %>% count()
esmap_asa %>% count(Priority)%>% View()

esmap_asa%>%
  openxlsx::write.xlsx("output/esmap_asa.xlsx", sheetName = "ee_asa", firstRow = TRUE, firstCol = TRUE)

skim(esmap_asa)
View(esmap_asa)
save(esmap_asa,file = "data/esmap_asa.rda")


load("data/esmap_asa.rda")
ee_wb_esmap <- esmap_asa

ee_wb <- ee_wb_gp %>% #ee portfolio
  full_join(ee_wb_esmap, by = "pid") %>%
  #merge to universe
  left_join(operation, by = "pid") %>%
  left_join(sct_project_3, by = "pid") %>%
  mutate(`amount(m)` = amount/1e6, #amount_gp_ee is in million, amount is in dollar.
         `amount(m)` = ifelse(!is.na(amount_gp_ee),amount_gp_ee,`amount(m)`), 
         income = recode(income,
                         Low = "LIC",
                         High = "HIC",
                         `Lower Middle` = "LMIC",
                         `Upper Middle` = "UMIC"),
         status = toupper(status)) %>%
  select(var_keep)
  

anyDuplicated(ee_wb$pid)
skim(ee_wb)

ee_wb$inst %>% table()
ee_wb %>% filter(inst == "WB-Lend") %>% skim(`amount(m)`)
ee_wb %>% filter(inst == "WB-ASA") %>% skim(`amount(m)`)

ee_wb %>% filter(is.na(income)) %>% distinct(ctr)
ee_wb %>% filter(is.na(iso3c)) %>% distinct(ctr)
ee_wb %>% filter(is.na(lend_instr_wb)) %>% distinct(inst)


save(ee_wb,file = "data/ee_wb.rda")


ee_wb_raw %>%
  #filter(dplyr::between(`Approval FY`,2011,2020)) %>%
  nrow() #244 from GP list.(there's duplciation)

ee_wb_gp %>%
  nrow() #241 unique from GP list

ee_wb_gp %>%
  left_join(operation_0,,by=c("pid"="PROJ_ID")) %>%
  dplyr::count(inst) #232 of them are lending. 

ee_wb %>%
  mutate(appfy = as.numeric(appfy)) %>%
  #filter(dplyr::between(appfy,2011,2020)) %>%
  filter(grepl("WB-Lend",inst)) %>%
  nrow() #232 of them are lending. 


ee_wb %>%
  mutate(appfy = as.numeric(appfy)) %>%
  filter(dplyr::between(appfy,2011,2020)) %>%
  filter(grepl("WB-Lend",inst)) %>%
  nrow() #209 of them are in the evaluation period.  

