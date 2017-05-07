#APICallers
This subroutine will perform four API Calls to get the value in € of the last trade that occurred at Kraken of Bitcoin, Ethereum, Ethereum classic and Litecoin. These values are then passed to different cells in the sheet 'Cryptocurrency values'. It is easily extendable to include other cryptocurrencies or API Calls to other websites. These values can be refreshed by pushing Ctrl + L when in the spreadsheet.

# Requirements / First configuration
You need to
[x] [Enable developer tab on the ribbon] (https://msdn.microsoft.com/nl-nl/library/bb608625.aspx)
[x] Enable Microsoft scripting Runtime setting. When inside a module in Visual Basic, click on Tools -> References, and scroll down to Microsoft Scripting Runtime)

# Optional
If you do not wish to use the .xlsm file you can implement the code yourself from APICallers.cls file.
You need the JsonConverter module for this to work. This can be retrieved @: https://github.com/VBA-tools/VBA-JSON


