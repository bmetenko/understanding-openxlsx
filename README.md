# understanding-openxlsx
Future home of an educational Shiny application for deciphering the use of the C++ dependent openxlsx R package, and comparing it to the rJava dependent xlsx R package.

These two package implementations are both able to read Microsoft Excel based files share the added benefit of being able to customize the resulting excel files. This isn't very common since most packages with excel file reading abilities tend to de-emphasize the resultant file's readability and aesthetics, in favor of simplicity and efficiency. The syntax associated between the openxlsx and xlsx packages share commonalities, but there isn't a clear guide on this matter. The following document and associated Shiny [work in progress] application aim to be an educational resource for anyone wanting to up their excel game programmatically in an R-based environment.

Later plans on using the learnr package for question and test logic.

**FUNCTION NAME SIMILARITIES**|  -
|--------|----------|                                        
| __xlsx__ | __openxlsx___ |
| ADD  | Functions |
| `addAutoFilter`, `addDataFrame`, `addHyperlink`, `addMergedRegion`, `addPicture` | `addCreator`, `addFilter`, `addStyle`, `addWorksheet` |