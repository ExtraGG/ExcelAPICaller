Attribute VB_Name = "ExcelAPIModule"
'This subroutine will perform four subsequent API Calls. You need the JsonConverter module for this to work. This can be retrieved here: https://github.com/VBA-tools/VBA-JSON'

'Bitcoin API call'
Public Sub APICallers()
    Dim http As Object, BTC As Object, btcvalue As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://api.kraken.com/0/public/Ticker?pair=xbteur", False
    http.Send
    
    Set BTC = JsonConverter.ParseJson(http.responsetext)
    
    btcvalue = BTC("result")("XXBTZEUR")("c")(1)
    
    Sheets(1).Cells(2, 2).Value = (btcvalue)
    
    'Ethereum api call'
    Dim ethvalue As String, ETH As Object
    
    http.Open "GET", "https://api.kraken.com/0/public/Ticker?pair=etheur", False
    http.Send
    
    Set ETH = JsonConverter.ParseJson(http.responsetext)
    
    ethvalue = ETH("result")("XETHZEUR")("c")(1)
    
    Sheets(1).Cells(4, 2).Value = (ethvalue)
    
    'Ethereum classic api call'
    Dim etcvalue As String, ETC As Object
    
    http.Open "GET", "https://api.kraken.com/0/public/Ticker?pair=etceur", False
    http.Send
    
    Set ETC = JsonConverter.ParseJson(http.responsetext)
    
    etcvalue = ETC("result")("XETCZEUR")("c")(1)
    
    Sheets(1).Cells(6, 2).Value = (etcvalue)
    
    'Litecoin api call'
    Dim ltcvalue As String, LTC As Object

    http.Open "GET", "https://api.kraken.com/0/public/Ticker?pair=ltceur", False
    http.Send
    
    Set LTC = JsonConverter.ParseJson(http.responsetext)
    
    ltcvalue = LTC("result")("XLTCZEUR")("c")(1)
    
    Sheets(1).Cells(8, 2).Value = (ltcvalue)
End Sub

