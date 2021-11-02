# Data cleaning -----------------------------------------------------------

packages<-c("tidyverse","readxl","rlist","rio","janitor","pdftools","openxlsx",
            "dint","lubridate","zoo","groupdata2","xts","tabulizer",
            "flextable")
#sapply(packages,install.packages,character.only=T)
sapply(packages,library,character.only=T)

#1.GDP
gdp_data<-read_excel(path = "HBS_Table_No._153___Quarterly_Estimates_of_Gross_Domestic_Product_at_Factor_Cost_(at_Current_Prices)_(New_Series)_(Base__2004-05).xlsx",
                     sheet = "NAS� 2011-12")
names(gdp_data)<-gdp_data[3,]
gdp_data<-gdp_data %>% clean_names()
gdp_data<-gdp_data %>% fill(year_industry)
gdp_data<-drop_na(gdp_data)
gdp_data<- gdp_data%>% slice(-c(1))

gdp_data<-gdp_data %>% separate(year_industry,into = c("Year",NA),
                                sep = "-",remove = T)
gdp_data<-gdp_data %>% unite(col = "Year/Quarter",c(Year,quarter),remove = T,
                             sep = "-")
gdp_data$`Year/Quarter`<-as.yearqtr(gdp_data$`Year/Quarter`,
                                    format = "%Y-Q%q")
gdp_data<-gdp_data %>% mutate(Date=as.Date(gdp_data$`Year/Quarter`)) %>% 
        mutate(Date=seq(as.Date(gdp_data$`Year/Quarter`[2]),length.out=nrow(gdp_data),
                        by="quarter")) %>% 
        select(Date,everything())

gdp_gva<-gdp_data %>% select(1:2,11)

#2. Trade Balance
tb_data<-read_excel("HBS_Table_No._190___India's_Foreign_Trade_-_Rupees.xlsx")
names(tb_data)<-tb_data[3,]
tb_data<-tb_data %>% clean_names()
tb_data<-drop_na(tb_data)
tb_data<- tb_data%>% slice(-c(1))
tb_data<-tb_data %>% filter(month!="Annual")
tb_data<-tb_data %>% separate(year,into = c("Year",NA),
                                sep = "-",remove = T)
tb_data<-tb_data %>% unite(col = "Year/Month",c(Year,month),remove = T,
                             sep = "-")
tb_data$`Year/Month`<-as.yearmon(tb_data$`Year/Month`,
                                    format = "%Y-%b")

tb_data<-tb_data %>% arrange(`Year/Month`)

qtr_function <- function(df, fn = "mean") {
        df2 <- group(df, floor(nrow(df) / 3), force_equal = T)
        df2_xts <- as.xts(df2, df2$`Year/Month`)
        df2_qtr <- apply.quarterly(df2_xts[, -c(1, ncol(df2_xts))], fn)
        yr <- year(index(df2_qtr))
        yr_qtr <- paste(yr, paste0("Q", quarter(yearqtr(index(df2_qtr)), fiscal_start = 4)))
        yr_qtr <- as.yearqtr(yr_qtr)
        
        df2_final <- data.frame(Period = yr_qtr, df2_qtr, row.names = NULL)
        if (fn == "sum") {
                df2_final <- df2_final[, -1] * 3
                df2_final <- data.frame(Period = yr_qtr, df2_final)
                return(df2_final)
        }
        
        
        
        
        return(df2_final)
}
tb_data<-qtr_function(tb_data) %>% arrange(Period)
tb_data<-tb_data %>% mutate(Date=as.Date(tb_data$Period)) %>% 
        mutate(Date=seq(as.Date(tb_data$Period[2]),length.out=nrow(tb_data),
                        by="quarter")) %>% 
        select(Date,everything())
tba_final<-tb_data %>% select(c(1:2,5))


#3.CAB
cab_data<-read_excel("HBS_Table_No._194___Indias_Overall_Balance_of_Payments_-_Quarterly_-_Rupees.xlsx",
                     sheet = "2000-01 Q1 Onwards")
names(cab_data)<-cab_data[4,]
cab_data<-cab_data %>% clean_names()
cab_data<-cab_data %>% remove_empty("rows")
cab_data <-cab_data %>% slice(-c(1:2,nrow(cab_data)))
cab_data<-cab_data %>% separate(year,into = c("Year",NA),
                                sep = "-",remove = T)
cab_data<-cab_data %>% unite(col = "Year/Quarter",c(Year,quarter),remove = T,
                             sep = "-")
cab_data$`Year/Quarter`<-as.yearqtr(cab_data$`Year/Quarter`,
                                    format = "%Y-Q%q")
cab_data<-cab_data %>% mutate(Date=as.Date(cab_data$`Year/Quarter`)) %>% 
        mutate(Date=rep(seq(as.Date(cab_data$`Year/Quarter`[4]),length.out=nrow(cab_data)/3,
                                  by="quarter"),each=3)) %>% 
        select(Date,everything())
cab_final<-cab_data %>% select(c(1:3,"a_current_account")) %>%
        filter(transaction_type=="Net") %>% select(-3,
                                                   CAB=a_current_account)


#4.Net international investment position
pdf_list<-list.files(pattern = "*.PDF")
tab_list<-map(pdf_list,extract_tables) %>% set_names(2015:2020)

niip_2018<-data.frame(
  stringsAsFactors = FALSE,
               Period = c("Mar-18(PR)",
                          "Jun-18(PR)","Sep-18(PR)","Dec-18(PR)","Mar-19(P)","Net IIP",
                          "-418.5","-407.5","-387.3","-426.9","-436.4",
                          "A. Assets","633.7","611.1","608.2","606.4","642.1",
                          "1. Direct Investment","157.4","160.9","163.5",
                          "166.6","170.0","2. Portfolio","Investment","3.6",
                          "3.1","2.6","2.7","4.7","2.1 Equity","Securities",
                          "2.1","1.9","1.8","1.4","0.6",
                          "2.2 Debt Securities","1.5","1.1","0.8","1.3","4.1",
                          "3. Other Investment","48.2","41.3","41.5","41.6","54.5",
                          "3.1 Trade Credits","1.7","1.4","0.9","0.3","0.9",
                          "3.2 Loans","8.2","7.0","7.1","6.6","9.9",
                          "3.3 Currency and","Deposits","20.8","16.3","16.6","17.2",
                          "25.2","3.4 Other Assets","17.5","16.7","16.9",
                          "17.5","18.5","4. Reserve Assets","424.5","405.7",
                          "400.5","395.6","412.9","B. Liabilities","1052.3",
                          "1018.5","995.5","1033.3","1078.5",
                          "1. Direct Investment","379.0","372.3","362.1","386.2","399.2",
                          "2. Portfolio","Investment","272.2","254.3",
                          "237.9","245.8","260.0","2.1 Equity","Securities",
                          "155.1","144.4","135.2","138.1","147.5",
                          "2.2 Debt securities","117.0","109.8","102.7","107.8","112.5",
                          "3. Other Investment","401.2","391.9","395.5",
                          "401.3","419.3","3.1 Trade Credits","103.2","99.6",
                          "104.3","103.6","105.2","3.2 Loans","159.8",
                          "156.5","157.6","160.5","168.1","3.3 Currency and",
                          "Deposits","126.5","124.5","122.1","126.0","130.6",
                          "3.4 Other Liabilities","11.7","11.3","11.5","11.2",
                          "15.2","Memo item: Assets to Liability Ratio (%)",
                          "60.2","60.0","61.1","58.7","59.5")
   )
niip_2018<-niip_2018[1:11,,drop=T]
niip_2018_b<-niip_2018[7:end(niip_2018)]
niip_2018_final<-rbind(c("Period",niip_2018[1:5]),c("Net IIP",niip_2018_b)) %>% 
        data.frame()
names(niip_2018_final)<-niip_2018_final[1,] 
niip_2018_final<-niip_2018_final %>% slice(-1)


niip2015<-tab_list$`2015`[[3]] %>% data.frame()
niip2016<-tab_list$`2016`[[2]] %>% data.frame()
niip2017<-tab_list$`2017`[[2]] %>% data.frame()
niip2019<-tab_list$`2019`[[2]]  %>% data.frame()
niip2020<-tab_list$`2020`[[2]] %>% data.frame()

niip_list<-ls(pattern = "niip*")[-c(1:4,7)]  
rm(niip_list)
niip_list<-map(niip_list,get) 
niip_a<-map(niip_list,slice,1:2) %>% map(function(x){
        names(x)<-x[1,]
        x<-x[-1,,drop=F]
        return(x)
}) %>% reduce(left_join)

names(niip_a)[1]<-names(niip_a)[2]
niip_a<-niip_a[,-2]
cbind(niip_a,niip_2018_final)

niip2019 <-niip2019%>% separate(col = X2,into = c("Mar","Jun","Sep","Dec","Mar20"),
                      sep = " ") %>% slice(1:2) %>% 
        select(-6) %>% slice(-1)
names_2019<-c("Mar-19 (R)", "Jun-19 (PR)", "Sep-19 (PR)", "Dec-19 (PR)", "Mar-20 (P)")

names(niip2019)<-c("Period",names_2019)
niip_b<-left_join(niip_a,niip_2018_final) %>% left_join(niip2019)

niip2015<-niip2015[,-2]
names(niip2015)<-c("Period",paste0(niip2015[1,2],
                                   niip2015[2,2]),
                   niip2015[1,-c(1:2)])
niip2015<-niip2015 %>% slice(3)

niip<-cbind(niip2015,niip_b)

niip<-niip %>% select(c(1,2:6,8:17,23:31,18:22)) %>% 
        pivot_longer(-Period,names_to = "Month/Year",
                     values_to = "NIIP") %>% select(-Period)
niip<-niip %>% slice(-c(5,10,15,20))
niip$`Month/Year`<-gdp_gva$`Year/Quarter`[16:40]
niip<-niip %>% mutate(
        Date=gdp_gva$Date[16:40]
) %>% select(Date,everything())

master_data<-list(niip,gdp_gva,cab_final,tba_final)
master_data<-master_data %>% reduce(left_join) %>% 
        select(Date,`Year/Quarter`,everything(),-c(`Month/Year`,Period)) %>% 
        slice(-1)
master_data[,-c(1:2)]<-lapply(master_data[,-c(1:2)],as.numeric)
master_data<-master_data %>% 
        mutate(
                NIIP=100*NIIP
        ) %>% 
        mutate("NIIP/GDP"=NIIP/total_gross_value_added_at_basic_price,
               "TB/GDP"=trade_balance/total_gross_value_added_at_basic_price,
               "CAB/GDP"=CAB/total_gross_value_added_at_basic_price)


#Plot
data_use<-master_data %>% select(`Year/Quarter`,7:9)
data_long<-data_use %>% pivot_longer(-`Year/Quarter`,
                                     names_to = "Series",
                                     values_to = "Values")
p1<-data_long %>% filter(Series!="TB/GDP") %>% 
        ggplot(aes(x=`Year/Quarter`,
                   y=Values,
                   color=Series))+
        geom_line(size=1.1)+
        theme_bw()+
        scale_x_yearqtr(format = "%Y-Q%q")+
        labs(y="Ratio")
ggsave(p1,dpi = 400,filename = "plot.jpg")
        
data_use<-data_use %>% 
        mutate(
                Difference=`CAB/GDP`-`TB/GDP`
        )
data_use[,-1]<-lapply(data_use[,-1],as.numeric)
#Table
ft<-flextable(data_use)
ft<-theme_booktabs(ft)
ft<-add_footer_lines(ft,"Difference refers to difference between ratio of Current Account Balance to GDP and Trade Balance to GDP (Difference= CAB/GDP - TB/GDP)")

#Word table
library(officer)
read_docx() %>% 
        body_add_par("Table: Data") %>% 
        body_add_flextable(value = ft) %>% 
        print(target = "table_word.docx")

#excel table
write.xlsx(data_use,"Data.xlsx",overwrite = T)
