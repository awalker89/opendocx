
#' @useDynLib opendocx
#' @importFrom Rcpp sourceCpp 
Doc <- setRefClass("Doc", fields = c(".rels",
                                     "app",
                                     "Content_Types",
                                     "core",
                                     "fontTable",
                                     "settings",
                                     "styles",
                                     "stylesWithEffects",
                                     "theme",
                                     "document.xml.rels",
                                     "document",
                                     "webSettings")
)


Doc$methods(initialize = function(creator = Sys.info()[["login"]]){
  
  Content_Types <<- genBaseContent_Type()
  .rels <<- genBaseRels()
  app <<- genBaseApp()
  core <<- genBaseCore()
  document.xml.rels <<- genBaseDocument.xml.rels()
  theme <<- genBaseTheme()
  
  fontTable <<- genBaseFontTable()
  settings <<- genBaseSettings()
  styles <<- genBaseStyles()
  stylesWithEffects <<- genBaseStylesWithEffects()
  webSettings <<- genBaseWebSettings()
  document <<- NULL
  
})




Doc$methods(zipDoc = function(zipfile, files, flags = "-r1", extras = "", zip = Sys.getenv("R_ZIPCMD", "zip"), quiet = TRUE){ 
  
  ## code from utils::zip function (modified to not print)
  args <- c(flags, shQuote(path.expand(zipfile)), shQuote(files), extras)
  
  if(quiet){
    
    res <- invisible(suppressWarnings(system2(zip, args, stdout = NULL)))
    
  }else{
    if (.Platform$OS.type == "windows"){
      res <- invisible(suppressWarnings(system2(zip, args, invisible = TRUE)))
    }else{
      res <- invisible(suppressWarnings(system2(zip, args)))
    }
  }
  
  if(res != 0){
    stop("zipping up docx failed. Please make sure Rtools is installed or a zip application is available to R.
         Try installr::install.rtools() on Windows.", call. = FALSE)
  }
  
  invisible(res)
  
})




Doc$methods(saveWorkbook = function(quiet = TRUE){
    
  pxml <- function(x){
    paste(unique(unlist(x)), collapse = "")
  }
  
  ## temp directory to save XML files prior to compressing
  tmpDir <- file.path(tempfile(pattern="docTemp_"))
  
  if(file.exists(tmpDir))
    unlink(tmpDir, recursive = TRUE, force = TRUE)
  
  success <- dir.create(path = tmpDir, recursive = TRUE)
  if(!success)
    stop(sprintf("Failed to create temporary directory '%s'", tmpDir))
  
  relsDir <- file.path(tmpDir, "_rels")
  dir.create(path = relsDir, recursive = TRUE)
  
  docPropsDir <- file.path(tmpDir, "docProps")
  dir.create(path = docPropsDir, recursive = TRUE)
  
  wordDir <- file.path(tmpDir, "word")
  dir.create(path = wordDir, recursive = TRUE)
  
  wordRelsDir <- file.path(tmpDir, "word","_rels")
  dir.create(path = wordRelsDir, recursive = TRUE)
  
  wordthemeDir <- file.path(tmpDir, "word", "theme")
  dir.create(path = wordthemeDir, recursive = TRUE)
  
  

  ## Write content
  ## write .rels
  .Call("opendocx_writeFile", '', pxml(.rels), '', file.path(relsDir, ".rels"))
  .Call("opendocx_writeFile", '', pxml(document.xml.rels), '', file.path(wordRelsDir, "document.xml.rels"))
  
  ## write app.xml
  .Call("opendocx_writeFile", '', pxml(app), '', file.path(docPropsDir, "app.xml"))
  
  ## write core.xml
  .Call("opendocx_writeFile", '', pxml(core), '', file.path(docPropsDir, "core.xml"))
  
  ## will always have a theme
  lapply(1:1, function(i){
    con <- file(file.path(wordthemeDir, paste0("theme", i, ".xml")), open = "wb")
    writeBin(charToRaw(pxml(theme[[i]])), con)
    close(con)        
  })
  
  ## write [Content_type]       
  .Call("opendocx_writeFile", '', pxml(Content_Types), '', file.path(tmpDir, "[Content_Types].xml"))
  
  ## write styles.xml
  .Call("opendocx_writeFile",
        '<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">',
        pxml(styles),
        '</w:styles>',
        file.path(wordDir,"styles.xml"))
  
  .Call("opendocx_writeFile", '', pxml(stylesWithEffects), '', file.path(wordDir,"stylesWithEffects.xml"))

  .Call("opendocx_writeFile", '', pxml(fontTable), '', file.path(wordDir,"fontTable.xml"))
  .Call("opendocx_writeFile", '', pxml(settings), '', file.path(wordDir,"settings.xml"))
  .Call("opendocx_writeFile", '', pxml(webSettings), '', file.path(wordDir,"webSettings.xml"))
  
  .Call("opendocx_writeFile",
        '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:body>',
        pxml(document),
        '</w:body></w:document>',
        file.path(wordDir,"document.xml"))
  
  ## compress to docx
  setwd(tmpDir)
  tmpFile <- tempfile(tmpdir = tmpDir, fileext = ".docx")
  tmpFile <- basename(tmpFile)
  
  zipDoc(tmpFile, list.files(tmpDir, recursive = TRUE, include.dirs = TRUE, all.files=TRUE), quiet = quiet)
  
  invisible(list("tmpDir" = tmpDir, "tmpFile" = tmpFile))
  
})


Doc$methods(buildTable = function(x, layout = "portrait", colour = NULL){
  
  if(!is.null(colour)){
    
    # Table style
    borderColour <- colour
    headerFillColour <- colour
    headerFontColour <- "FFFFFF"
    
    styleId <- paste(sample(c(LETTERS[1:6], 1:6), size = 12), collapse = "")
      
    tableDef1 <- sprintf('<w:style w:type="table" w:styleId="%s"><w:name w:val="%s"/>', styleId, styleId)
    
    tableDef2 <- '<w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="60"/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    
    borderXML <- paste0('<w:tblBorders>',
                        sprintf('<w:top w:val="single" w:sz="8" w:space="0" w:color="%s"/>', borderColour),
                        sprintf('<w:left w:val="single" w:sz="8" w:space="0" w:color="%s"/>', borderColour),
                        sprintf('<w:bottom w:val="single" w:sz="8" w:space="0" w:color="%s"/>', borderColour),
                        sprintf('<w:right w:val="single" w:sz="8" w:space="0" w:color="%s"/>', borderColour),
                        '</w:tblBorders>')
    
    tcBorderXML <- gsub("tblBorders", "tcBorders", borderXML)
    
    ## table properties
    p4 <- paste('<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblInd w:w="0" w:type="dxa"/>', borderXML,
                '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>' , collapse = "")
    
    p5 <- sprintf('<w:tblStylePr w:type="firstRow">
                  <w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
                  </w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="%s"/></w:rPr><w:tblPr/><w:tcPr>
                  <w:shd w:val="clear" w:color="auto" w:fill="%s"/>
                  </w:tcPr>
                  </w:tblStylePr>', headerFontColour, headerFillColour)
    
    lastRow <- sprintf('<w:tblStylePr w:type="lastRow">
                  <w:pPr>
                  <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
                  </w:pPr>
                  <w:rPr><w:b/><w:bCs/></w:rPr><w:tblPr/>
                  <w:tcPr>%s</w:tcPr></w:tblStylePr>', tcBorderXML)
      
    band1Vert <- sprintf('<w:tblStylePr w:type="band1Vert">
                                 <w:tblPr/>
                                 <w:tcPr>
                                 %s
                                 </w:tcPr>
                                 </w:tblStylePr>', tcBorderXML)
    
    
    band1Horz <- sprintf('<w:tblStylePr w:type="band1Horz">
                                 <w:tblPr/>
                                 <w:tcPr>
                                 %s
                                 </w:tcPr>
                                 </w:tblStylePr>', tcBorderXML)
    
    styles <<- c(styles, paste(tableDef1, tableDef2, p4, p5, lastRow, band1Vert, band1Horz, '</w:style>', sep = ""))
    
  }else{
    styleId <- "TableGrid"
  }
  ## Document table element
  
  ## table properties
  tblPr <- sprintf('<w:tblPr>
  <w:tblStyle w:val="%s"/>
  <w:tblW w:w="0" w:type="auto"/>
  <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
  </w:tblPr>', styleId)
  
#   ## table grid
#   tblGrid <- '<w:tblGrid><w:gridCol w:w="6588"/><w:gridCol w:w="6588"/></w:tblGrid>'
#   tblGrid <- NULL
  
  ## table header
  rowXML1 <- sprintf('<w:tc><w:tcPr><w:cnfStyle w:val="001000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/>
                     <w:tcW w:w="6588" w:type="dxa"/></w:tcPr>
                     <w:p w:rsidR="001A42CF" w:rsidRDefault="001A42CF">
                     <w:r><w:t>%s</w:t></w:r>
                     </w:p>
                     </w:tc>', names(x))
  
  rowXML1 <- paste0('<w:tr w:rsidR="001A42CF" w:rsidTr="001A42CF">
                    <w:trPr><w:cnfStyle w:val="100000000000" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/>
                    </w:trPr>', paste(rowXML1, collapse = ""), '</w:tr>')    
  
  ## all other rows
  rowXML <- apply(x, 1, makeTableCell)
  
  rowXML <- paste0('<w:tr w:rsidR="001A42CF" w:rsidTr="001A42CF"><w:trPr>
                   <w:cnfStyle w:val="000000100000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="1" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/>
                   </w:trPr>', rowXML, '</w:tr>', collapse = "")
  
  
  document <<- c(document, paste('<w:tbl>', tblPr, rowXML1, rowXML, '</w:tbl>', addPage(layout = layout)))
  
  invisible(1)

})





addPage <- function(layout = "portrait"){
  
  layout <- tolower(layout)
  
  if(layout %in% "portrait"){
    
    xml <- '<w:p w:rsidR="001A42CF" w:rsidRDefault="001A42CF"/>
    <w:sectPr w:rsidR="001A42CF" w:rsidSect="001A42CF">
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:space="720"/>
    <w:docGrid w:linePitch="360"/>
    </w:sectPr>'
    
    
  }else if(layout %in% "landscape"){
    
    xml <- '<w:p w:rsidR="001A42CF" w:rsidRDefault="001A42CF"/>
    <w:sectPr w:rsidR="001A42CF" w:rsidSect="001A42CF">
    <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:space="720"/>
    <w:docGrid w:linePitch="360"/>
    </w:sectPr>'
    
    
  }else{
    stop("Invalid layout.")
  }
  
  return(xml)
  
}


