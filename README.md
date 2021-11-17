# ToolsForParsingExcelData
A small group of tools to aid it extracting data from Excel Files.

The password for the Files is "clean"

There is a blog post that goes with 2 of these files located at https://pcsxcetrasupport3.wordpress.com/2021/11/16/excel-4-macro-code-obfuscation/

Tool named:: "Clean_And_Extract_Excel_Unicode_Shared_Strings" This tool is for use with the binary version of the shared strings.
  To use this tool: Find the binary file and open it in a hex editor copy and paste all of the bytes to the input and click the buooton.
  If it does not have a 0x13 in an odd location it should work properly. (Tool needs a rewrite for new information discoverd during the writing of the blog post)
  
Tool named:: "Parse_Excel4_SharedStrings" This tool is used with the xml version of the shared strings
  To use this tool: Navigate to the location of the xml version of the shared strings and select the file and then click the button.
  
Tool named:: "Parse_XLSM_MacroSheet_Value_Data" This tool is for use with xml version of the macrosheets It will extract only data from the value tag.
  To use this tool: Navigate to the location of the xml version of the macroshet select the file and then click the button.
  This tool was designed to extract the masive amounts of decimal values that end up converted to a HATA file.
  
Tool named:: "Parse_XLSM_MacroSheetData" was designed to extract the data from the macrosheet from this zloader sample.
  https://labs.inquest.net/dfi/sha256/00af74bbf7d5430b861d3fcd79afb184e352bbd27348613659e7769f0ef1207f
  To use this tool: Navigate to the xml version of the macrosheet and select file then click the button. 
  This tool will fail on sample like in the blogpost that use different layouts.
  
Tool named:: "Parse_Zloader_Excel4_W-FORMULA-FILL" was designed to work with macrosheets that use formulas. It will work on blogpost sample.
  To use this tool: Navigate to the xml version of the macrosheet of the type with formulas in it and select file and then click the button.


Tool named:: "ConvFromDecCharCode" was designed for converting decimal char codes to char. It allows you to pick a delimiter Char. 
  Default is * and may need to be changed.
  This version will Ignore "Null Bytes" when processing the char codes.
  To use this tool: Copy paste a decimal array into the input box. Input the correct delimeter. Click the button.
     
