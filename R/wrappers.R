



#' @name write.docx
#' @title write data.frame to docx table
#' @param x a data.frame
#' @param file file to save to
#' @param layout "portrait" or "landscape" page orientation
#' @param overwrite logical. Overwrite existing file.
#' @param colour A valid hex colour code starting with '#' eg. '#4f81BD'
#' @author Alexander Walker
#' @return Workbook object
#' @export
#' @import methods
#' @examples
#' write.docx(x = iris, "testing.docx", layout = "portrait", colour = NULL)
#' write.docx(x = iris, "testing1.docx", layout = "portrait", colour = "#4f81BD")
#' write.docx(x = iris, "testing2.docx", layout = "landscape", colour = "#C0504D")
write.docx <- function(x, file, layout = "portrait", overwrite = TRUE, colour = NULL){
  
  doc <- Doc$new()
  
  if(!"data.frame" %in% class(x))
    stop("x must be a data.frame")
  
  doc$buildTable(x, layout = layout, colour = colour)
  saveDoc(doc = doc, file = file, overwrite = overwrite)
  
  invisible(doc)
  
}



saveDoc <- function(doc, file, overwrite = FALSE){
  
  wd <- getwd()
  on.exit(setwd(wd), add = TRUE)
  
  ## increase scipen to avoid writing in scientific 
  exSciPen <- options("scipen")
  options("scipen" = 10000)
  on.exit(options("scipen" = exSciPen), add = TRUE)
  
  if(!"Doc" %in% class(doc))
    stop("First argument must be a Workbook.")
  
  if(!grepl("\\.docx", file))
    file <- paste0(file, ".docx")
  
  if(!is.logical(overwrite))
    overwrite = FALSE
  
  if(file.exists(file) & !overwrite)
    stop("File already exists!")
  
  tmp <- doc$saveWorkbook(quiet = TRUE)
  setwd(wd)
  
  file.copy(file.path(tmp$tmpDir, tmp$tmpFile), file, overwrite = overwrite)
  
  ## delete temporary dir
  unlink(tmp$tmpDir, force = TRUE, recursive = TRUE)
  
  invisible(1)
  
  
}