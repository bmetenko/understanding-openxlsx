library(here)
# library(xlsx)
library(openxlsx)

setwd(here("data"))

# Following Alexander Walker's Introduction documents.
vignette("Introduction", package = "openxlsx")
vignette("formatting", package = "openxlsx")

# Basic write. Zero styling/
write.xlsx(iris, file = "test.iris.xlsx")

# Can be written as a table formatted data frame. Basic table format.
write.xlsx(iris, file = "test.table.iris.xlsx", asTable = T)

# Named lists as seperators for different worksheets.
l <- list("Iris" = iris, "Mtcars" = mtcars)

write.xlsx(l, file = "test.sheets.xlsx")
write.xlsx(l, file = "test.table.sheets.xlsx", asTable = T)

# Change default options on output.

options()[grep(names(options()), pattern = "openxlsx")] # odd??

options("openxlsx.borderColour" = "green") # hex input also.
options("openxlsx.borderStyle" = "thin") # others?
options("openxlsx.dateFormat" = "mm/dd/yyyy")
options("openxlsx.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
options("openxlsx.numFmt" = NULL) ## For default style rounding of numeric columns

# openxlsx options initialized after being specified.
## where are the defaults stored? Function parameters?

df <- data.frame("Date" = Sys.Date()-0:19, "LogicalT" = TRUE,
                 "Time" = Sys.time()-0:19*60*60,
                 "Cash" = paste("$",1:20), "Cash2" = 31:50,
                 "hLink" = "https://CRAN.R-project.org/",
                 "Percentage" = seq(0, 1, length.out=20),
                 "TinyNumbers" = runif(20) / 1E9, 
                 stringsAsFactors = FALSE)

# Class assignment as a determination of formatting.
class(df$Cash) <- "currency"
class(df$Cash2) <- "accounting"
class(df$hLink) <- "hyperlink"
class(df$Percentage) <- "percentage"
class(df$TinyNumbers) <- "scientific"


# Clean and simple, and uneffected by anything but style formatting through class designations.
write.xlsx(df, "test.df.custom.format.xlsx")

# Similar, except table formatting applied.
# Unclear if green border color was applied. TODO.
write.xlsx(df, file = "test.table.df.custom.format.xlsx", asTable = TRUE)

# Creating styles.
new_style <- createStyle(fontName = "Times",fontSize = 12, fontColour = "blue", bgFill = "gray", halign = "center", valign = "center", textDecoration = "Bold", border = "TopBottom")

str(new_style)

# Style is a custom S4 class with specific parameters.
# Methods belong to the instances of this class. Methods are contained in the instance, as well as data that can be accessed using the `$` operator.

attributes(new_style)


# Specifies borders on the whole dataset for the output. Does this override headerStyle?
write.xlsx(iris, file = "test.border.rows.new_style.xlsx", borders = "rows", headerStyle = new_style)

write.xlsx(iris, file = "test.border.cols.new_style.xlsx", borders = "columns", headerStyle = new_style)

write.xlsx(iris, "test.rotated.headers.xlsx", asTable = TRUE,
           headerStyle = createStyle(textRotation = 45)) 
# Table styling still applied first.
## Main Educational document idea: Styling and output >>> Mardown document with images.

# All list elements, read sheets, will get the style applied. Confirmed.

l <- list("IRIS" = iris, "colClasses" = df)
write.xlsx(l, file = "list.new_style.xlsx", borders = "columns", headerStyle = new_style)
write.xlsx(l, file = "list.table.style.Light2.xlsx", asTable = TRUE, tableStyle = "TableStyleLight2")

workbook_object <- write.xlsx(l, file = "list.new_style.xlsx", borders = "columns", headerStyle = new_style)

str(workbook_object)

# Deply nested object mapping of workbook.

workbook_object

object.size(workbook_object)
# Workbook object is modified and called in place, without the need for reassignment.
setColWidths(workbook_object, sheet = 1, cols = 1:5, widths = 20)


# > workbook_object$colWidths
# [[1]]
# 1    2    3    4    5 
# "20" "20" "20" "20" "20" 
# attr(,"hidden")
# [1] "0" "0" "0" "0" "0"
# 
# [[2]]
# list()

# write call can return a usable object variable and re-exported post processing.

# saveWorkbook(wb, "writeXLSX6.xlsx", overwrite = TRUE)

require(ggplot2)

wb <- createWorkbook()

options("openxlsx.borderColour" = "#4F80BD")
options("openxlsx.borderStyle" = "thin")

# Default style for font is assigned seperately.

modifyBaseFont(wb, fontSize = 10, fontName = "Arial Narrow")
# Modifies/appends wb$styles$fonts[[1]] with applicable xml tags list.

# Add two more worksheets. Gridlines?
addWorksheet(wb, sheetName = "Motor Trend Car Road Tests", gridLines = FALSE)
addWorksheet(wb, sheetName = "Iris", gridLines = FALSE)


freezePane(wb, sheet = 1, firstRow = TRUE, firstCol = TRUE) 
## freeze first row and column
writeDataTable(wb, sheet = 1, x = mtcars,
               colNames = TRUE, rowNames = TRUE,
               tableStyle = "TableStyleLight9")

setColWidths(wb, sheet = 1, cols = "A", widths = 18)

# Most methods can be called from instance.
# Example wb$setColWidths(sheet = 1, cols = "A", width = 18)


writeDataTable(wb, sheet = 2, iris, startCol = "K", startRow = 2)
# Write data table definition is quite dense, but write data is a method that can be called with wb$method

qplot(data=iris, x = Sepal.Length, y= Sepal.Width, colour = Species)

# Viewport rendered plot can be attached at a specific location
insertPlot(wb, 2, xy=c("B", 16)) ## insert plot at cell B16 in sheet 2, how is size decided? 4" x 6" default. Calls device copy, so don't change it before you do. xml based inserting of image? Stored in $drawings



means <- aggregate(x = iris[,-5], by = list(iris$Species), FUN = mean)
vars <- aggregate(x = iris[,-5], by = list(iris$Species), FUN = var)


## write group means ### Custom styled list object.
headSty <- createStyle(fgFill="#DCE6F1", halign="center", border = "TopBottomLeftRight")

writeData(wb, 2, x = "Iris dataset group means", startCol = 2, startRow = 2)

writeData(wb, 2, x = means, startCol = "B", startRow=3, borders="rows", headerStyle = headSty)

## write group variances
writeData(wb, 2, x = "Iris dataset group variances", startCol = 2, startRow = 9)

writeData(wb, 2, x= vars, startCol = "B", startRow=10, borders="columns", headerStyle = headSty)

setColWidths(wb, 2, cols=2:6, widths = 12) ## width is recycled for each col

setColWidths(wb, 2, cols=11:15, widths = 15)
# style mean & variance table headers

# Interesting that text decoration is c()'ed together.
s1 <- createStyle(fontSize=14, textDecoration=c("bold", "italic"))

addStyle(wb, 2, style = s1, rows=c(2,9), cols=c(2,2))

saveWorkbook(wb, "custom.with.plot.xlsx", overwrite = TRUE) ## save to working directory
# Looks very good. No grid theming is interesting especially considering the way you can style the rest of the excel file. Green borders worked this time.

# BELOW: ADD sizing checks on object sizes, and final excel size differences.
# Examples from vignette.
# Plot input and compression.
# Image outputs.


# Empty workbook overhead.

wb <- openxlsx::createWorkbook()
object.size(wb) # 688 bytes

wb2 <- xlsx::createWorkbook()
object.size(wb2) # 944 bytes


str(wb)
str(wb2)

# Profiling
profvis::profvis({
  wb <- openxlsx::createWorkbook()
  
  wb2 <- xlsx::createWorkbook()
})
