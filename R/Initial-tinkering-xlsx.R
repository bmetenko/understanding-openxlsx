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
# "autoSizeColumn"        

# [BORDER STYLING AND CELL BLOCK (CB) STYLING]
# - # - # - # - # - # - # - # - # - # - # 
# "Border"                
# "BORDER_STYLES_"        
# "CB.setBorder"          
# "CB.setColData"          
# "CB.setFill"            
# "CB.setFont"            
# "CB.setMatrixData"      
# "CB.setRowData"         
# "CELL_STYLES_"          
# "CellBlock"             
# "CellProtection"        
# "CellStyle"              

# [CREATE COMMANDS]
# - # - # - # - # - # - # - # - # - # - # 
# "createCell"            
# "createCellComment"
# "createFreezePane"      
# "createRange"            
# "createRow"             
# "createSheet"           
# "createSplitPane"       
# "createWorkbook"         

# "DataFormat"            

# [FILL & FORCE]
# - # - # - # - # - # - # - # - # - # - # 
# "Fill"                   
# "FILL_STYLES_"          
# "Font"                   
# "forceFormulaRefresh"   
# "forcePivotTableRefresh" 

# [GET COMMANDS]
# - # - # - # - # - # - # - # - # - # - # 
# "get_java_tmp_dir"
# "getCellComment"
# "getCells"              
# "getCellStyle"
# "getCellValue"          
# "getRanges"
# "getRows"               
# "getSheets"              

# "HALIGN_STYLES_"        
# "INDEXED_COLORS_"        

# [IS.CLASSNAME] [Wrappers around inherits(x, classname)]
# - # - # - # - # - # - # - # - # - # - # 
# "is.Alignment"          
# "is.Border"              
# "is.CellBlock"          
# "is.CellProtection"      
# "is.CellStyle"          
# "is.DataFormat"          
# "is.Fill"               
# "is.Font"                

# [LOAD & READ]
# - # - # - # - # - # - # - # - # - # - # 
# "loadWorkbook"          
# "printSetup"             
# "read.xlsx"             
# "read.xlsx2"             
# "readColumns"           
# "readRange"              
# "readRows"      

# [REMOVE COMMANDS]
# - # - # - # - # - # - # - # - # - # - # 
# "removeCellComment"
# "removeMergedRegion"    
# "removeRow"
# "removeSheet"           

# [SAVE & SET COMMANDS]
# - # - # - # - # - # - # - # - # - # - # 
# "saveWorkbook"           
# "set_java_tmp_dir"      
# "setCellStyle"
# "setCellValue"          
# "setColumnWidth"
# "setPrintArea"          
# "setRowHeight"
# "setZoom"               

# "VALIGN_STYLES_"   

# [WRITE COMMANDS]
# - # - # - # - # - # - # - # - # - # - # 
# "write.xlsx"            
# "write.xlsx2"      