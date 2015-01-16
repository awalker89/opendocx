
opendocx
========
Write a data.frame to a docx document.

```
x <- head(iris)
doc <- write.docx(x, "test1.docx", layout = "portrait", colour = NULL)
doc <- write.docx(x, "test2.docx", layout = "portrait", colour = "#4f81BD")
doc <- write.docx(x, "test3.docx", layout = "landscape", colour = "#C0504D")

```

## Installation
```
install_github("awalker89/opendocx")
```

