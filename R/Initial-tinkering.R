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

