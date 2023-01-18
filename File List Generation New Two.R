Excel_file_list = Sys.glob(file.path("C:", "Users", "sanka", "Documents", "Company Finance Records Three", "*.xlsx"))

length_args = length(Excel_file_list)
page_count = ceiling(length_args/12)

library("readxl")
pdf("Test File Report Two.pdf", onefile = TRUE)

for (k in 1:page_count){
  par(mfrow = c(3, 4))
  
  for (i in 1:12){
    a = 12*(k - 1) + i
    
    if (a == length_args + 1) {
      #next
      break
    }
    
    file = try(Excel_file_list[a])
    
    if(grepl("~", file)){
      next
    }
    
    dataframe_iter = read_excel(file)
    #data_min = min(dataframe_iter[, 2:3])
    row_num = nrow(dataframe_iter)
    
    if(class(dataframe_iter)[3] == "try-error"){
      next
    }
    
    if(row_num == 0){
      next
    }
    
    #if(data_min < 0){
    #next
    #}
    
    ticker = strsplit(tail(strsplit(file, "/")[[1]], 1), " ")[[1]][1]
    
    if(ticker == "LHX"){
      p = as.numeric(tail(dataframe_iter$Year, 3)[1]) + .5
      q = as.numeric(tail(dataframe_iter$Year, 3)[1]) + 1
      dataframe_iter[row_num - 1, 1] = toString(p)
      dataframe_iter[row_num, 1] = toString(q)
      dataframe_iter$Year <- as.numeric(dataframe_iter$Year)
      
      row.names(dataframe_iter) <- dataframe_iter$Year
      dataframe_iter$`Total Revenue` <- as.numeric(dataframe_iter$`Total Revenue`)
      dataframe_iter$'Earnings Before Interest and Taxes' <- as.numeric(dataframe_iter$`Earnings Before Interest and Taxes`)
      dataframe_iter_new = dataframe_iter[, 2:3]
      print(dataframe_iter)
      #print(as.matrix(dataframe_iter))
      l = (i - 1) %% 4
      m = floor((i - 1)/4)
      #pdf("Files Report.pdf")
      par(fig=c((10/4)*l, (10/4)*(l+1), (10*(3-1)/3) - (10/3)*m, 10 - (10/3)*m)/10, new = TRUE)
      #title = paste(ticker, "Earnings and Revenue")
      dataframe_iter_new_T = t(as.matrix(dataframe_iter_new))
      colnames(dataframe_iter_new_T) <- dataframe_iter$Year
      barplot(t(as.matrix(dataframe_iter_new)), main = ticker, xlab = "Year", col = c("blue", "green"), beside = TRUE, names.arg=dataframe_iter$Year)
      legend("topleft", c("Revenue", "Earnings"), fill = c("blue","green"))
      
      next
    }
    
    g = as.numeric(tail(dataframe_iter$Year, 2)[1]) + 1
    dataframe_iter[row_num, 1] = "0"
    dataframe_iter$Year <- as.numeric(dataframe_iter$Year)
    dataframe_iter[row_num, 1] = g
    
    
    row.names(dataframe_iter) <- dataframe_iter$Year
    dataframe_iter$`Total Revenue` <- as.numeric(dataframe_iter$`Total Revenue`)
    dataframe_iter$'Earnings Before Interest and Taxes' <- as.numeric(dataframe_iter$`Earnings Before Interest and Taxes`)
    dataframe_iter_new = dataframe_iter[, 2:3]
    print(dataframe_iter)
    #print(as.matrix(dataframe_iter))
    l = (i - 1) %% 4
    m = floor((i - 1)/4)
    #pdf("Files Report.pdf")
    par(fig=c((10/4)*l, (10/4)*(l+1), (10*(3-1)/3) - (10/3)*m, 10 - (10/3)*m)/10, new = TRUE)
    #title = paste(ticker, "Earnings and Revenue")
    dataframe_iter_new_T = t(as.matrix(dataframe_iter_new))
    colnames(dataframe_iter_new_T) <- dataframe_iter$Year
    barplot(t(as.matrix(dataframe_iter_new)), main = ticker, xlab = "Year", col = c("blue", "green"), beside = TRUE, names.arg=dataframe_iter$Year)
    legend("topleft", c("Revenue", "Earnings"), fill = c("blue","green"))
    
  }
  
  if (k == page_count){
    next
  }
  
  grid::grid.newpage()
}

dev.off()