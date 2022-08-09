Option Explicit

Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr
Dim testVal As Integer
Dim stocks(200) As String
Dim values(200, 2) As Double
Dim nextAvailableIndex As Integer


Function execShell(command As String, Optional ByRef exitCode As Long) As String
    Dim file As LongPtr
    file = popen(command, "r")
    
    If file = 0 Then
        MsgBox "exiting function"
        Exit Function
    End If
    
    While feof(file) = 0
       Dim chunk As String
        Dim read As Long
        chunk = Space(50)
        read = fread(chunk, 1, Len(chunk) - 1, file)
        execShell = execShell & chunk
        If read = 0 Then
            chunk = Left$(chunk, read)
        End If
    Wend
    
    exitCode = pclose(file)
End Function

Function HTTPGet(URL As String) As String

    Dim lExitCode As Long
    Dim cmd As String
    cmd = "curl " & URL
    HTTPGet = execShell(cmd, lExitCode)
    
End Function

Function getStockData(tickerID As String) As Double()
    Dim rawData As String
    Dim results(2) As Double
    rawData = HTTPGet("https://quotes-gw.webullfintech.com/api/stock/tickerRealTime/getQuote?tickerId=" & tickerID)
    Dim word As String
    Dim character As String
    Dim count As Integer
    Dim nextValue As String
    
    Dim i As Long
    For i = 1 To Len(rawData)
        character = Mid(rawData, i, 1)
        
        If character <> """" Then
            word = word & character
        Else
            word = Replace(word, " ", "")
            If nextValue <> "" Then
                count = count + 1
                If nextValue = "price" And count = 2 Then
                    results(0) = CDbl(word)
                    nextValue = ""
                End If
                If nextValue = "eps" And count = 2 Then
                    results(1) = CDbl(word)
                    nextValue = ""
                End If
                If nextValue = "dividend" And count = 2 Then
                    results(2) = CDbl(word)
                    nextValue = ""
                End If
            End If

            If word = "close" Then
                nextValue = "price"
                count = 0
            End If
            If word = "epsTtm" Then
                nextValue = "eps"
                count = 0
            End If
            If word = "dividend" Then
                nextValue = "dividend"
                count = 0
            End If
            word = ""
        End If
    Next i
    
    getStockData = results
End Function

Function getTickerId(symbol As String)
    Dim URL As String
    Dim searchResults As String
    Dim tickerID As String
    URL = """https://quotes-gw.webullfintech.com/api/search/pc/tickers?keyword=" & symbol & "&regionId=6&pageIndex=1&pageSize=1"""
    searchResults = HTTPGet(URL)
    Dim i As Long
    Dim word As String
    Dim character As String
    Dim save As Boolean
    Dim second As Boolean
    Dim symbolFound As String
    
    
    For i = 1 To Len(searchResults)
        character = Mid(searchResults, i, 1)
        
        If character <> """" Then
            word = word & character
            If word = "tickerId" Then
                tickerID = Mid(searchResults, i + 3, 9)
            End If
        Else
            If save Then
                symbolFound = word
                save = False
            End If
            
            If second Then
                save = True
                second = False
            End If
            
            If word = "symbol" Then
                second = True
            End If
            word = ""
        End If
        
    Next i
    
    If symbolFound = symbol Then
        getTickerId = tickerID
    Else
        getTickerId = "failed"
    End If
    
End Function

Function updateStocks()
    Dim sheet As Worksheet
    Set sheet = Worksheets("Sheet1")
    Dim tickerID As String
    Dim symbol As String
    Dim results() As Double
    Dim i As Long
    i = 1
    While Not IsEmpty(sheet.Cells(i, 1))
        symbol = sheet.Cells(i, 1).Value
        tickerID = getTickerId(symbol)
        
        If tickerID = "failed" Then
            i = i
        Else
            results = getStockData(tickerID)
            sheet.Cells(i, 2).Value = results(0)
            sheet.Cells(i, 3).Value = results(1)
            sheet.Cells(i, 4).Value = results(2)

        End If
        
        i = i + 1
    Wend
End Function

Function updateStock(i As Integer, Optional passedSymbol As String = "")
    Dim sheet As Worksheet
    Set sheet = Worksheets("Sheet1")
    Dim tickerID As String
    Dim symbol As String
    Dim results() As Double
    If i = 0 Then
        symbol = passedSymbol
    Else
        symbol = sheet.Cells(i, 1).Value
    End If
    tickerID = getTickerId(symbol)
        
    If tickerID <> "failed" Then
        results = getStockData(tickerID)
        stocks(nextAvailableIndex) = symbol
        values(nextAvailableIndex, 0) = results(0)
        values(nextAvailableIndex, 1) = results(1)
        values(nextAvailableIndex, 2) = results(2)
        nextAvailableIndex = nextAvailableIndex + 1
        sheet.Cells(i, 2).Formula = "=quote(A" + CStr(i) + ")"
        sheet.Cells(i, 3).Formula = "=eps(A" + CStr(i) + ")"
        sheet.Cells(i, 4).Formula = "=dividend(A" + CStr(i) + ")"
    End If
End Function

Function testing()
    testVal = testVal + 1
    testing = testVal
End Function

Function quote(symbol As String)
    Dim i As Long
    Dim found As Boolean
    For i = 0 To 200
        If stocks(i) = symbol Then
            quote = values(i, 0)
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        Dim tickerID As String
        tickerID = getTickerId(symbol)
        If tickerID <> "failed" Then
            quote = getStockData(tickerID)(0)
        End If
    End If
End Function

Function eps(symbol As String)
    Dim i As Long
    Dim found As Boolean
    For i = 0 To 200
        If stocks(i) = symbol Then
            eps = values(i, 1)
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        Dim tickerID As String
        tickerID = getTickerId(symbol)
        If tickerID <> "failed" Then
            eps = getStockData(tickerID)(1)
        End If
    End If
End Function

Function dividend(symbol As String)
    Dim i As Long
    Dim found As Boolean
    For i = 0 To 200
        If stocks(i) = symbol Then
            dividend = values(i, 2)
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        Dim tickerID As String
        tickerID = getTickerId(symbol)
        If tickerID <> "failed" Then
            dividend = getStockData(tickerID)(2)
        End If
    End If
End Function
