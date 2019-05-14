#-------------------------------------------
# Functions.R - associated with ECCC_CaribouPM
# used by several child modules in ECCC_Caribou.PM
#-------------------------------------------

#' Read an Excel file
#'
#' This is a wrapper around openxlsx::read.xlsx
#'
#'
#' @param filename Character string indicating file path of Excel file
#' @param sheetname Character string indicating name of Excel sheet
#' @importFrom openxlsx read.xlsx
#' @importFrom data.table data.table
#' @export
readXlsx <- function(filename, sheetname) {
  if (!file.exists(filename)){
    stop(filename," does not exist.")
  }
  out <- data.table(read.xlsx(filename,
                              sheet=sheetname,
                              colNames=TRUE))
  if (!is.null(out$sPopIdList)) {
    out$sPopIdList <- as.character(out$sPopIdList)
  }
  if (!is.null(out$PopId)) {
    out$PopId <- as.character(out$PopId)
  }
  
  return(out)
  
}

#' Write an Excel file
#'
#' This is a wrapper around openxlsx::read.xlsx
#'
#'
#' @param filename Character string indicating file path of Excel file to write
#' @param inputObject R object to write to Excel
#' @importFrom openxlsx read.xlsx createWorkbook writeData addWorksheet saveWorkbook
#' @importFrom data.table data.table
#' @export
writeXlsx <- function(inputObject, filename) {
  # assumes library(openxlsx) has been loaded
  # designed here to use lists to write multiple objects to multiple sheets in the same workkbook
  # prior step 1: build the list of data.frames: df.list <- lapply(c("a.df","b.df"), function(x) get(x))
  # prior step 2: name each data.frame with the sheetname: names(df.list) <- paste0(c("name1","name2"))
  # inputObject must be a list of named objects (e.g., dataframes)
  # in the call to writeData(), the current list to be written is coerced to be a dataframe, so that colNames=TRUE works
  wb <- createWorkbook()
  # the if-else structure below is used only because I'm not sure how to write an lapply() function to deal with a list of length 1
  if (length(inputObject) > 1) {
    df.list <- inputObject
    wb <- createWorkbook()
    lapply(names(df.list),
           function(df){
             sheetname <- df
             addWorksheet(wb, sheetname)
             writeData(wb,
                       sheetname,
                       data.frame(df.list[[df]]),
                       rowNames=FALSE,
                       colNames=TRUE)
           } )
  } else {
    sheetname <- names(inputObject)
    addWorksheet(wb, sheetname)
    writeData(wb,
              sheetname,
              data.frame(inputObject),
              rowNames=FALSE,
              colNames=TRUE)
  }
  saveWorkbook(wb,filename, overwrite=TRUE)
}