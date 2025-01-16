1.Class "exceloperate"
  读取（xlsx/xls）/将需求数据按格式（如单元格合并，底色，边框线条，以及字体样式（大小，颜色，对齐位置及其它）写入xlsx文档，并添加统计图表和图像
  Read (xlsx/xls)/Write the required data in formats such as cell merging, background color, border lines,
  and font styles (size, color, alignment, and others) into an xlsx document, and add statistical charts and images
2.Class "ExcelGrid"
  1.该ExcelGrid只支持xlsx文档、预览、编辑、保存、打印、缩放、数据格式化、冻结、大纲、公式计算、图表、脚本执行等功能。
  2.支持toolbar工具栏隐藏，以及隶属按钮显示控制。
  3.去掉Readonly只读属性后可模似Excel编辑单元格，但不考虑太多的界面单元格编辑逻辑工作，如单元格合并，底色，边框线条，以及字体样式（大小，颜色，对齐位置及其它），建议采用类“NetOffice.ExcelOperate”来对xlsx格式进行编辑。
  4.ExcelGrid支持Tab键焦点跳转及方法跳转。
  5.设计該ExcelGrid主要考虑目的是与xlsx文档界面交互应用以及预览,在VFP完成填写一些复杂的二维表类型，以直接引用xlsx编辑好的文档格式来填充数据，无非是一种更方便快捷低成本的报表应用方案,且文档通用。
  ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  Regarding:
  1. This ExcelGrid only supports functions such as xlsx documents, previewing, editing, saving, printing, scaling, data formatting, freezing, outlining, formula calculation, charts, script execution, etc.
  2. Support hiding the toolbar and controlling the display of subordinate buttons.
  3. After removing the Readonly read-only attribute, cells can be edited similar to Excel, but without considering too much interface cell editing logic work, such as cell merging, background color, border lines, and font style (size, color, alignment position, etc.). It is recommended to use the class "NetOffice. Excel Operate" to edit xlsx format.
  4. ExcelGrid supports Tab key focus jump and method jump.
  5. The main purpose of designing this ExcelGrid is to interact with and preview XLSX document interfaces. Completing some complex two-dimensional table types in VFP and directly referencing XLSX edited document formats to fill in is simply a more convenient, fast, and low-cost report application solution, and the document is universal.
