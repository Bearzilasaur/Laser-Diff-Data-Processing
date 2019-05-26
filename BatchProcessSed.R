#R script to batch process the standardized laser diff particle excel output
#
#Steps for processing
#
#1. User select files which comprise the batch to be processed
#   a. save paths to list variable 
#
#
#2. Method for iterating through each file, possibly a FOR loop or some other
#   function. Maybe pass a list to openxls argument
#
#3. Sanitize file name: remove string '.$ls' 
#
#
#3. Select raw data from within the excel file
#   a. Data should be in the range A75:C167
#   b. This range translates to rows(75:167) & cols(1:3)
#      ^^ NOTE: This is within excel, r may start lists from 0
#
#4. Save the selected data to a data.frame and add headers.
#   a. Saves data from range to variable framdat
#   b. Converts the mixed two columns into sorted 3 columns
#   c. applies the appropriate headers for each column
#
#
#5. Export relevant range using the same filename as original into a new
#   directory within the selected directory.
#   a. write.csv

#importing relevant libraries

library(readxl)
library(tcltk)




#----- 1. USER SELECT Files and The Directory -----
btchFiles <- tk_choose.files()



# Ask if need conversion from weird filetype or not
#yn<- menu(c('Yes', 'No'), title = 'Convert to .csv?',  graphics = TRUE )
yn<-tk_messageBox(type = 'yesno', 
                  message = 'Would you like to convert the .$ls.xls files to 
                             .csv files?')



#~~~~ Test to make sure that all files reside within the same folder~~~~~#
for (i in btchFiles){                                                            #iterates through the list of selected files
  
    if (isTRUE(dirname(btchFiles[i]) != dirname(btchFiles[1]))){                #checks to see if all files in list are equal NOTE THIS COULD BE BETTER
        
       tk_messageBox(type = 'ok', title('Error'), text("Error: All selected
                                                       files to be processed
                                                       must reside within the
                                                       same folder"))
      
      stop()
    }
}
#----- 2. LOOP through files and copy and save selected data -----
for (x in btchFiles){
  
  
  
  # 2a. FILE CHECKS Checks user wants to convert files -----
  #if (yn == "yes"){
    
    
  i <- gsub('.[.$]ls[.]xls', '.txt', x)
      
  tk_messageBox(type = 'ok', message = paste(i))
  tk_messageBox(type = 'ok', message = paste(x))
  
      



# 3. STORAGE the relevant data from workbook i in btchfiles in the -----
#    variable rawdat
  #if (grepl('*.$ls.xls', i) == FALSE){
    rawdat <- read.delim(x)
  
  
  
  # 4. CONVERSION the tab delim data to a data.frame format-----
    framdat<- data.frame(rawdat)
  
  
  
  # 4. b) Sorts data into appropriate format
    srtdDat<- data.frame(
        framdat[c(seq(69, nrow(framdat), by = 2)), 1],
        framdat[c(seq(69, nrow(framdat), by = 2)), 2],
        framdat[c(seq(70, nrow(framdat) + 1, by = 2)), 1]
        
        )
    
  
  
  # 4. c) Adds appropriate Headers to the datatable
    names(srtdDat)[c(1:3)]<- c('BinNumber', 
                               'BinMaxSedDiameter', 
                               'ProportionOfSample')
  
  
  
  # 5. WRITES the data.frame framdat to a folder called OUTPUT -----
    newdir<-as.character(paste(dirname(i), '/OUTPUT', sep = ""))
    dir.create(newdir, showWarnings = FALSE)
    setwd(newdir)
    
    filename = gsub(dirname(i), "", i)
    
    write.csv(srtdDat, file= paste(newdir, filename, sep = ""))
    
  
    next()
  
  #}else{
    
    #stop()
  #}
  
  
  

}
#-----Process Complete MessageBox -----
tk_messageBox(type = 'ok', title('Success!'), text('Batch Complete!'))















