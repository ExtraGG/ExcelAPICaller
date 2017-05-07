## Description
This subroutine will perform four API Calls to get the value in € of the last trade that occurred at Kraken of Bitcoin, Ethereum, Ethereum classic and Litecoin in Microsoft Excel. These values are then passed to different cells in the sheet 'Cryptocurrency values'. It is easily extendable to include other cryptocurrencies or API Calls to other websites. You can easily link the macro to a shortcut by adding a shortcut at Developer tab -> Macro's -> Edit.

## Requirements / First configuration
When using the .xlsm file (easier):
- [x] [Enable developer tab on the ribbon](https://msdn.microsoft.com/nl-nl/library/bb608625.aspx)
- [x] Enable Microsoft scripting Runtime setting. When inside a module in Visual Basic, click on Tools -> References, and scroll down to Microsoft Scripting Runtime.

## Optional
If you do not wish to use the .xlsm file you can implement the code yourself from the APICallers.cls file. You still need to enable both settings listed above for it to work. Download the .cls file and import it with file -> import file when inside Visual Basic. It will be added under modules. 

In this case you also need the JsonConverter module for it to work. This can be retrieved [here](https://github.com/VBA-tools/VBA-JSON).


