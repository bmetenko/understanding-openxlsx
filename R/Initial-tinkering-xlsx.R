library(xlsx)

# Possibly include rJava configuration files.
# - # List of all functions in package.
ls("package:xlsx")

# [ADD INTO SHEET Functionality]
# - # - # - # - # - # - # - # - # - # - # 
# "addAutoFilter"
# xlsx::addAutoFilter -> function (sheet, cellRange) 
# """rJava::J class method call to util function CellRangeAddress.
# Adds to sheet through method of R class sheet == sheet$setAutoFilter(cellRangeAddress)"""
 

# "addDataFrame"          
# "addHyperlink"
# "addMergedRegion"       
# "addPicture"             

# // "Alignment"             
# [7] "autoSizeColumn"        

# [BORDER STYLING AND CELL BLOCK (CB) STYLING]
# - # - # - # - # - # - # - # - # - # - # 
# "Border"                
# [9] "BORDER_STYLES_"          "CB.setBorder"          
# [11] "CB.setColData"          "CB.setFill"            
# [13] "CB.setFont"             "CB.setMatrixData"      
# [15] "CB.setRowData"          "CELL_STYLES_"          
# [17] "CellBlock"              "CellProtection"        
# [19] "CellStyle"              

# [CREATE COMMANDS]
# - # - # - # - # - # - # - # - # - # - # 
# "createCell"            
# [21] "createCellComment"      "createFreezePane"      
# [23] "createRange"            "createRow"             
# [25] "createSheet"            "createSplitPane"       
# [27] "createWorkbook"         

# "DataFormat"            

# [FILL & FORCE]
# - # - # - # - # - # - # - # - # - # - # 
# [29] "Fill"                   "FILL_STYLES_"          
# [31] "Font"                   "forceFormulaRefresh"   
# [33] "forcePivotTableRefresh" 

# [GET COMMANDS]
# - # - # - # - # - # - # - # - # - # - # 
# "get_java_tmp_dir"      
# [35] "getCellComment"         "getCells"              
# [37] "getCellStyle"           "getCellValue"          
# [39] "getRanges"              "getRows"               
# [41] "getSheets"              

# "HALIGN_STYLES_"        
# [43] "INDEXED_COLORS_"        

# [IS.CLASSNAME] [Wrappers around inherits(x, classname)]
# - # - # - # - # - # - # - # - # - # - # 
# "is.Alignment"          
# [45] "is.Border"              "is.CellBlock"          
# [47] "is.CellProtection"      "is.CellStyle"          
# [49] "is.DataFormat"          "is.Fill"               
# [51] "is.Font"                

# [LOAD & READ]
# - # - # - # - # - # - # - # - # - # - # 
# "loadWorkbook"          
# [53] "printSetup"             "read.xlsx"             
# [55] "read.xlsx2"             "readColumns"           
# [57] "readRange"              "readRows"      

# [REMOVE COMMANDS]
# - # - # - # - # - # - # - # - # - # - # 
# [59] "removeCellComment"      "removeMergedRegion"    
# [61] "removeRow"              "removeSheet"           

# [SAVE & SET COMMANDS]
# - # - # - # - # - # - # - # - # - # - # 
# [63] "saveWorkbook"           "set_java_tmp_dir"      
# [65] "setCellStyle"           "setCellValue"          
# [67] "setColumnWidth"         "setPrintArea"          
# [69] "setRowHeight"           "setZoom"               

# [71] "VALIGN_STYLES_"   

# [WRITE COMMANDS]
# - # - # - # - # - # - # - # - # - # - # 
# "write.xlsx"            
# [73] "write.xlsx2"      