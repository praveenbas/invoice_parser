require("xlsx")

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

#parse_invoice(invoiceFile = as.vector(invoice.df$invoicePath[7]),invoiceFname =as.vector(invoice.df$invoiceFilename[2]))





parse_invoice<-function(invoiceFile,invoiceFname){
  
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
  Betrag=NA
  if(length(Brutto) == 0){
    Brutto<-NA
    Betrag<- as.numeric(f[grep("Betrag",x= f[,"NA..23"]),"NA..29"])
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
  if(length(seq_cost) > 0){
    seq_cost_in<-grep(pattern = "sequencing|Sequencing",f$NA..1,perl = T)[1] # get sequncing section start index
    Bio_cost<-grep(pattern = "Bioinformatics",f$NA..1,perl = T)[1]
    if(length(Bio_cost) > 0){ # if bioinformatics section is in the invoice 
      seq_cost_end<-grep(pattern = "Bioinformatics",f$NA..1,perl = T)[1] -1 # get the Bioinfo section index
    }else{ # if no bioinfo section is not included
      seq_cost_end<- seq_cost_in + 4
    }
  }
  
  if(length(seq_cost) > 0){
    ngs_cost<-f[c(seq_cost_in:seq_cost_end),"NA..29"]
    ngs_datamgt_cost<-sum(ngs_cost[!is.na(ngs_cost )])
  }else{
    ngs_datamgt_cost<- 0
  }
  
  if(Brutto %in% NA){
    Bioinfo_cost<- Betrag - ngs_datamgt_cost
  }else{
    Bioinfo_cost<- Brutto - ngs_datamgt_cost
  }
  
  invoice_table<-data.frame(Rech.NR=Rech.NR,rechnung=rechnung,kostenstelle=kostenstelle,fonds=Fonds,
                            psp.Element=PSP.Element,date=date,bruttoBetrag=Brutto,betrag=Betrag,
                            NGS_dataMgmtcost=ngs_datamgt_cost,Bioinfo_cost=Bioinfo_cost,year=year,
                            file=invoiceFname, Address=Address,Offer=offer)
  return(invoice_table)
}