Hello.

This is a readme file of an ExcelTH project.

ExcelTH is a libraries or classes that are designed to help people generate and service excel tables 
with help of mechanisms that are provided by NetOffice and Gembox projects.

The main idea of ExcelTH is to lead user away from index control, inline format and structure handling.
Instead, user just describes headers, rows and related formats, then passes plain-text-data in 
(or plain-objects-data, but carefully). That's it.

ExcelTH structure can be described in a short list:
ExcelTableHandler / GemboxTableHandler - these are core, you should touch it ONLY if you want to modify internal mechanics.
CellFormats - rows and headers formats [generally painting, sizing, etc] are to be described here. Result of static function affects one row by default.
Utilities - utility file that you can use and that is being used.
Logger - is an abstract logger class that you may inherit for your project's needs; for example, in Program.cs it has realisation to write logs to file.
ReportData - example of generating or obtaining some plain-data for the table.
Program - example of workbook file preparations (sheets, names, save path, and so on).

ExcelTH project is free of use (CPOL). I will be glad if while using ExcelTH, you will save a small credits to an author in commentaries.

Also remember, that gratitude feels good, while currency supports good. :)

If you like ExcelTH so much or it did come handy and saved you a lot of time (nerves), you may provide some support for me (appreciate for anything)
with Qiwi:
+7(nine-zero-four)6543257

Thanks for visiting and reading. Have a good luck!
