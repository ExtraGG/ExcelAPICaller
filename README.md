## Description
This subroutine will perform four API Calls to get the value in € of the last trade that occurred at Kraken of Bitcoin, Ethereum, Ethereum classic and Litecoin in Microsoft Excel. These values are then passed to several cells in the sheet 'Cryptocurrency values'. It is easily extendable to include other cryptocurrencies or API Calls to other websites. 

## Requirements / Installation
You need the [developer tab](https://msdn.microsoft.com/nl-nl/library/bb608625.aspx) to access Microsoft Excel's Visual Basic functionality. 

#### Option 1:
When using the .xlsm file:

- [x] Enable Microsoft scripting Runtime setting. When inside a module in Visual Basic, click on Tools -> References, and scroll down to Microsoft Scripting Runtime.
It should work now. Press Ctrl + L to refresh the values when inside your workbook.

#### Option 2
- If you do not wish to use the .xlsm file you can implement the code yourself from the ExcelAPI.bas file. You still need to enable both settings listed above for it to work. Download the .bas file and import it with file -> import file when inside Visual Basic. It will be added under modules. 

- In this case you also need the JsonConverter module for it to work. This can be retrieved [here](https://github.com/VBA-tools/VBA-JSON).

You can link the macro to a shortcut if you want by adding a shortcut at Developer tab -> Macro's -> Edit.
