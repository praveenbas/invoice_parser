require("xlsx")
require("docstring")

invoicelist<-list.files(path = "~/Desktop/invoice",pattern = "*.xls",all.files = T,full.names = T,recursive = T)
invocefilename<-list.files(path = "~/Desktop/invoice",pattern = "*.xls",all.files = T,full.names = F,recursive = T)
invoice.df<-data.frame(invoiceFilename=invocefilename,invoicePath=invoicelist)

# check if masterfile already exist 
if(file.exists("~/Desktop/invoice/Invoice_master_table.tsv")){  
  
  # if exist do not parse the file already parsed
  checkfile<-read.table("~/Desktop/invoice/Invoice_master_table.tsv",header = T,sep = "\t")
  checkfile$date<-as.Date(checkfile$date)
  invoice_toparse<-invoice.df[!invoice.df$invoiceFilename %in% checkfile$file,] 
  
}else{ #
  invoice_toparse<-invoice.df
  checkfile<- data.frame() # create a dummy df
}

if(nrow(invoice_toparse) > 0){
  invoice_toparse$invoicePath<-as.vector(invoice_toparse$invoicePath)
  invoice_toparse$invoiceFilename<-as.vector(invoice_toparse$invoiceFilename)
  toupdate<-data.frame()
  for(r in 1:nrow(invoice_toparse)){
    tempRech<-parse_invoice(invoiceFile = as.vector(invoice_toparse$invoicePath[r]),invoiceFname =as.vector(invoice_toparse$invoiceFilename[r]))
    #print(tempRech)
    toupdate<-rbind(toupdate,tempRech)
  }
  checkfile<-rbind(checkfile,toupdate) 
  write.table(x = checkfile,file = "~/Desktop/invoice/Invoice_master_table.tsv",append =F,quote = F,sep = "\t",row.names = F)
}

parse_invoice(invoiceFile = as.vector(invoice.df$invoicePath[3]),invoiceFname =as.vector(invoice.df$invoiceFilename[3]))





parse_invoice<-function(invoiceFile,invoiceFname){
  #' Parse xls files (invoice)
  #' 
  #' Function to parse the invoice xls files to extract information for reports
  #' @param invoiceFile Invoice file in xls format (qbic format xls files)
  #' @param invoiceFname Filename as string
  #' @return Data frame with extracted infos.
  #' @note This is a custom designed function and will not work with general xls files.
  
  f<-read.xlsx(invoiceFile,sheetIndex = 1)
  
  message(sprintf("Parsing :::: %s",invoiceFname))
  
  # convert factor columns to vectors
  for(x in 1:ncol(f)){
    class_obj<- class(f[,x])
    if(identical(class_obj,"factor")){
      f[,x]<-as.vector(f[,x])
    }
  }
  
  
  #sapply(1:ncol(f),FUN = function(x){sprintf("Col :: %s == class :::%s",colnames(f)[x],class(f[,x]))})
  
  ### grep specific i and j
  Fonds<-f[grep("Fonds",x = f[,"NA."]),"NA..5"]
  PSP.Element<-f[grep("PSP-Element:",x = f[,"NA."]),"NA..5"]
  rechnung<- paste(f[grep("^Rechnung$",x = f[,"NA."]),c("NA..19","NA..22","NA..31")],collapse = "/")
  Rech.NR<- f[grep("^Rechnung$",x = f[,"NA."]),"NA..31"]
  
  Brutto<- as.numeric(f[grep("Bruttobetrag",x= f[,"NA..22"]),"NA..29"])
  Brutto_index<-grep("Bruttobetrag",x= f[,"NA..22"])
  Betrag=NA
  vat =0
  Nettobetrag=0
  if(length(Brutto) == 0){
    Brutto= NA
    Betrag= as.numeric(f[grep("Betrag",x= f[,"NA..23"]),"NA..29"])
    betrag_index= grep("Betrag",x= f[,"NA..23"])[1]
    netto_index=betrag_index
  }else{
    Nettobetrag=  as.numeric(f[grep("Nettobetrag",x= f[,"NA..22"]),"NA..29"])
    netto_index= grep("Nettobetrag",x= f[,"NA..22"])[1]
    vat= as.numeric(f[grep(pattern = "%",f[,"NA..22"]),"NA..29"])
    vat_index= grep(pattern = "%",f[,"NA..22"])[1]
  }
  
  
  
  Address <- as.vector(f$NA.[c(20:28)])
  Address<-paste(x=Address[!is.na(Address)],collapse = " ; ")
  
  offer_search_term= paste(c("With reference","With regard"),collapse = "|")
  offer<- gsub(pattern = "\\)",replacement = "",x = gsub(x = grep(pattern = offer_search_term,x = f$NA..1,perl = T,value = T),pattern = "\\(",replacement = ""))
  
  if(length(offer) ==0){
    offer<- NA
  }
  
  date<- as.Date(f[grep("Tübingen, den",x= f[,"NA..36"]),"NA..42"])
  kostenstelle<-as.numeric(f[2,grep(pattern = "Kostenstelle",x=f[1,])])
  year <- as.numeric(gsub(pattern = " ",replacement = "",paste(f[grep("^Rechnung$",x = f[,"NA."]),c("NA..19")],collapse = "/")))
  
  seq_cost<-grep(pattern = "sequencing|Sequencing",f$NA..1,perl = T,value = T)
  project_mgmt_index= check_empty_betragcolum(indexes = grep(pattern = "Project|project|management|Management",f$NA..1,perl = T),table = f)
  Bio_cost<-grep(pattern ="Bioinformatics|differential|analysis",f$NA..1,perl = T)[1]
  
  if(length(seq_cost) > 0){
    seq_cost_in<-grep(pattern = "sequencing|Sequencing",f$NA..1,perl = T)[1]# get sequncing section start index
    
    if(length(project_mgmt_index) > 0){ # if project management section is in the invoice 
      seq_cost_end<-project_mgmt_index -1 # sequencing section ends
    }else if (length(project_mgmt_index) == 0 && length(Bio_cost) > 0){ ## if management cost is not there but bioinfo cost is there
      seq_cost_end=Bio_cost-1
    }else{ # if no bioinfo section is not included
      seq_cost_end<-  netto_index -1
    }
    ## estiamte ngs cost
    ngs_only_cost<-ifelse(length(seq_cost) > 0,
                          sum(f[c(seq_cost_in:seq_cost_end),"NA..29"],na.rm = T),
                          0)
  }else{
    ngs_only_cost=0
  }
  
  
  
  mgmt_cost<-ifelse(length(project_mgmt_index) > 0, 
                    sum(f[project_mgmt_index,"NA..29"],na.rm = T), 
                    0)
  
  
  Bioinfo_cost<-ifelse(length(Bio_cost) >0, 
                       sum(f[c(Bio_cost :(netto_index -1)),"NA..29"],na.rm = T),
                       0)
  
  ### if nettobetrag is present
  if(is.na(Betrag)){
    ### check total Betrag == Bioinfo_cost + ngs_only_cost +mgmt_cost
    amount_df<- total_sanity_check(Nettobetrag,Bioinfo_cost ,ngs_only_cost ,mgmt_cost)
  }else{
    amount_df<- total_sanity_check(Betrag,Bioinfo_cost ,ngs_only_cost ,mgmt_cost)
  }
  Bioinfo_cost<-amount_df[amount_df$Terms%in% "Bioinfo_cost","cost_list"]
  ngs_only_cost<-amount_df[amount_df$Terms%in% "ngs_only_cost","cost_list"]
  mgmt_cost<-amount_df[amount_df$Terms%in% "mgmt_cost","cost_list"]
  
  

  
  invoice_table<-data.frame(Rech.NR=Rech.NR,rechnung=rechnung,kostenstelle=kostenstelle,fonds=Fonds,
                            psp.Element=PSP.Element,date=date,bruttoBetrag=Brutto,betrag=Betrag,
                            NGS_cost=ngs_only_cost,Bioinfo_cost=Bioinfo_cost,project_mgmt=mgmt_cost,
                            year=year,vat=vat,nettobetrag=Nettobetrag,file=invoiceFname, Address=Address,Offer=offer)
  return(invoice_table)
}


check_empty_betragcolum<- function(table,indexes){
  #' check all index and return first index with betrag's field with some value.  
  returnindex=indexes
  for( k in 1:length(indexes)){
    returnindex[k]<-ifelse(is.na(table[indexes[k],"NA..29"]),FALSE,TRUE)
  }
  return(indexes[which(returnindex==TRUE)][1])
}

###  check if the total amount and sum of all section cost  are equal
### only handles if  total amount less than sum of all section
### have to include if betrag is greater than
total_sanity_check<-function(Betrag, Bioinfo_cost, ngs_only_cost,mgmt_cost){
  #' (internal function) Function to check if the amounts are matching 
  #' 
  #' check if the total amount and sum of all section cost  are equal and only handles if  total amount less than sum of all section. 
  #' @param Betrag total amount in the invoice
  #' @param Bioinfo_cost bioinformatics cost estimated from parsing the invoice
  #' @param ngs_only_cost sequencing cost
  #' @param mgmt_cost project management cost
  #' @section Updates needed:
  #'  This function needs to be updated for cases "betrag is greater than sum of sections"
  unequal=0
  
  cost_list= c(Bioinfo_cost,ngs_only_cost,mgmt_cost)
  cost_df<-data.frame(cost_list=cost_list,Terms=c("Bioinfo_cost","ngs_only_cost","mgmt_cost"))
  #cost_list1<-cost_list[cost_list !=0]
  if(!Betrag == sum(cost_list)){
    # total  is greater or lesser th
    unequal= ifelse(Betrag > sum(cost_list), "greater","less")
    gz<-data.frame(t(combn(cost_list,2)))
    if(unequal %in% "less"){
      gz$combn_score =rowSums(gz)
      gz1<-gz[gz$combn_score > Betrag,]
      gz2<-gz[gz$combn_score == Betrag,]
      
      common<-intersect(unname(unlist(gz1[,c(1,2)])),unname(unlist(gz2[,c(1,2)])))
      
      gz3<-unname(unlist(gz1[,c(1,2)]))
      
      cost_df[cost_df$cost_list == common,"cost_list"]<- common- gz3[gz3!= common]
      
    }
  }
  
  return(cost_df)
  
}




