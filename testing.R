
## 616 base XML



library(opendocx)
doc <- Doc$new()
doc$document <- readLines("C:/Users/Alex/Desktop/doc1/word/document.xml", warn = FALSE)[[2]]



x <- iris[1:4,]
doc$document <- write.docx(x, layout = "portrait")
writeLines(xml, "c:/users/alex/desktop/testing.xml")


##   
install.packages("opendocx_1.0.5.tar.gz", repos = NULL, type = "source")

doc <- Doc$new()
doc$document <- doc$buildTable(x, layout = layout, colour = colour)
doc$styles <- c(doc$styles, makeTableStyle())
doc$stylesWithEffects <- ""

saveDoc(doc = doc, file = file, overwrite = overwrite)




x <- iris[1:5,1:5]
doc <- write.docx(x, "c:/users/alex/desktop/testing.docx", layout = "portrait", colour = NULL)
shell("c:/users/alex/desktop/testing.docx")

doc <- write.docx(x, "c:/users/alex/desktop/testing1.docx", layout = "portrait", colour = "#4f81BD")
shell("c:/users/alex/desktop/testing1.docx")

doc <- write.docx(x, "c:/users/alex/desktop/testing2.docx", layout = "landscape", colour = "#C0504D")
shell("c:/users/alex/desktop/testing2.docx")







doc$document




doc$document


cat(doc$document)


grepl('943634', doc$styles)



doc$document




write.docx(x, "c:/users/alex/desktop/testing2.docx", layout = "landscape")

shell("c:/users/alex/desktop/testing.docx")


## plain table
<w:style w:type="table" w:styleId="TableGrid">
  <w:name w:val="Table Grid"/>
  <w:basedOn w:val="TableNormal"/>
  <w:uiPriority w:val="59"/>
  <w:rsid w:val="00D9152C"/>
  <w:pPr>
  <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
  </w:pPr>
  <w:tblPr>
  <w:tblInd w:w="0" w:type="dxa"/>
  <w:tblBorders>
  <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  </w:tblBorders>
  <w:tblCellMar>
  <w:top w:w="0" w:type="dxa"/>
  <w:left w:w="108" w:type="dxa"/>
  <w:bottom w:w="0" w:type="dxa"/>
  <w:right w:w="108" w:type="dxa"/>
  </w:tblCellMar>
  </w:tblPr>
  </w:style>
  


