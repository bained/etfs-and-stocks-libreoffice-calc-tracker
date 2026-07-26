'=====================================================================
' LibreOffice Basic Macro — Optimized European ETF Tracking
'
' Architecture:
' - Sheet "etf_data" — local cache (row 1: tickers, row 2: prices, rows 3+:
'   historical data)
' - Sheet "ETFs" — user interface
' - Sheet "settings" — configuration (C2: Log, C3: Delay ms)
' - Sheet "etf_logs" — auto-created log sheet (optional)
' - Loader: graphical Shape (text box) on the active sheet
' - Rate limiting: from settings!C3 (ms) pause between requests
'
' Dependencies: Windows (MSXML2.XMLHTTP.6.0)
'====================================================================

' ============================================================
' CONFIGURATION (constants)
' ============================================================

Public Const SHEET_DATA     As String = "etf_data"
Public Const SHEET_LOGS     As String = "etf_logs"
Public Const SHEET_SETTINGS As String = "settings"
Public Const USER_AGENT     As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

' Yahoo endpoints
Private YAHOO_ENDPOINTS As Variant

' ============================================================
' SETTINGS FROM "settings" SHEET
' ============================================================

' Reads a setting from the settings sheet.
' If the sheet or cell is empty, returns the default value.
Private Function ReadSetting(ByVal cellAddr As String, ByVal defaultVal As String) As String
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oSettings As Object
    
    On Error GoTo SettingError
    
    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    oSettings = GetSheetByName(oSheets, SHEET_SETTINGS)
    
    If oSettings Is Nothing Then
        ReadSetting = defaultVal
        Exit Function
    End If
    
    ' Parse address like "C2" — column (0-based), row (0-based)
    Dim col As Long, row As Long
    col = Asc(UCase(Mid(cellAddr, 1, 1))) - Asc("A")
    row = CLng(Mid(cellAddr, 2)) - 1
    
    Dim val As String
    val = Trim(GetCellString(oSettings, row, col))
    
    If val = "" Then
        ReadSetting = defaultVal
    Else
        ReadSetting = val
    End If
    Exit Function
    
SettingError:
    ReadSetting = defaultVal
End Function

' Returns True if logging is enabled (settings!C2 = "Yes")
Private Function IsLoggingEnabled() As Boolean
    Dim setting As String
    setting = UCase(ReadSetting("C2", "No"))
    IsLoggingEnabled = (setting = "YES")
End Function

' Returns the delay in ms (settings!C3)
Private Function GetDelayMs() As Long
    Dim setting As String
    setting = ReadSetting("C3", "250")
    
    On Error Resume Next
    GetDelayMs = CLng(setting)
    If Err.Number <> 0 Or GetDelayMs < 0 Then
        GetDelayMs = 250
    End If
End Function

' ============================================================
' INITIALIZATION
' ============================================================

Private Sub InitConfig()
    YAHOO_ENDPOINTS = Array( _
        "https://query1.finance.yahoo.com/v8/finance/chart/", _
        "https://query2.finance.yahoo.com/v8/finance/chart/" _
    )
End Sub

' ============================================================
' LOADER — toggles visibility of an existing Shape on the sheet
' ============================================================

Private gLoaderShape As Object
Private gLoaderCreated As Boolean
Private Const LOADER_SHAPE_NAME As String = "ETF_LOADER_SHAPE"

' Shows the Loader Shape (finds it by name or creates a new one).
' @param title Initial text
Private Sub LoaderShow(ByVal title As String)
    On Error GoTo LoaderError
    
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oDrawPage As Object
    Dim found As Boolean
    Dim i As Long
    
    oDoc = ThisComponent
    oSheet = oDoc.getCurrentController().getActiveSheet()
    oDrawPage = oSheet.getDrawPage()
    
    gLoaderCreated = False
    found = False
    
    ' Look for an existing Shape by name
    For i = 0 To oDrawPage.Count - 1
        If oDrawPage(i).Name = LOADER_SHAPE_NAME Then
            Set gLoaderShape = oDrawPage(i)
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        ' Create a new Shape
        Dim oShape As Object
        oShape = oDoc.createInstance("com.sun.star.drawing.TextShape")
        oShape.Name = LOADER_SHAPE_NAME
        oShape.setPosition(CreatePoint(8000, 4000))
        oShape.setSize(CreateSize(6000, 2500))
        
        ' Shape properies
        Dim oProp As Object
        Set oProp = oShape
        oProp.FillColor = RGB(255, 255, 200)
        oProp.FillStyle = 1
        oProp.LineColor = RGB(0, 100, 200)
        oProp.LineWidth = 50
        oProp.CharAlignment = 1
        oProp.CharWeight = 150
        oProp.CharHeight = 14
        oProp.CharColor = RGB(0, 0, 0)
        oShape.setString(title)
        
        oDrawPage.add(oShape)
        Set gLoaderShape = oShape
        gLoaderCreated = True
    Else
        ' Existing Shape — just update the text
        gLoaderShape.setString(title)
    End If
    
    ' Make the Shape visible
    gLoaderShape.Visible = True
    oDoc.draw(False)
    Exit Sub
    
LoaderError:
End Sub

' Hides the Loader.
Private Sub LoaderClose()
    Dim oSheet As Object
    Dim oDrawPage As Object
    Dim i As Long
    
    On Error Resume Next
    
    ' Find the Shape if not in memory
    If gLoaderShape Is Nothing Then
        oSheet = ThisComponent.getCurrentController().getActiveSheet()
        oDrawPage = oSheet.getDrawPage()
        
        For i = 0 To oDrawPage.Count - 1
            If oDrawPage(i).Name = LOADER_SHAPE_NAME Then
                Set gLoaderShape = oDrawPage(i)
                Exit For
            End If
        Next i
    End If
    
    If gLoaderShape Is Nothing Then Exit Sub
    
    If gLoaderCreated Then
        ' We created it — remove from DrawPage
        oSheet = ThisComponent.getCurrentController().getActiveSheet()
        oDrawPage = oSheet.getDrawPage()
        oDrawPage.remove(gLoaderShape)
    Else
        ' It existed before us — just hide it
        gLoaderShape.Visible = False
    End If
    
    Set gLoaderShape = Nothing
    gLoaderCreated = False
End Sub

' Updates the Loader text.
Private Sub LoaderSetText(ByVal text As String)
    On Error Resume Next
    If Not gLoaderShape Is Nothing Then
        gLoaderShape.setString(text)
    End If
End Sub

' Helper functions for creating Point and Size objects
Private Function CreatePoint(ByVal x As Long, ByVal y As Long) As Object
    Dim pt As Object
    Set pt = CreateUnoService("com.sun.star.awt.Point")
    pt.X = x
    pt.Y = y
    Set CreatePoint = pt
End Function

Private Function CreateSize(ByVal width As Long, ByVal height As Long) As Object
    Dim sz As Object
    Set sz = CreateUnoService("com.sun.star.awt.Size")
    sz.Width = width
    sz.Height = height
    Set CreateSize = sz
End Function

' ============================================================
' PUBLIC FUNCTION — FOR THE BUTTON
' ============================================================

' Main function for updating all ETF data.
' Assign this macro to the button in the ETFs sheet.
Sub UpdateEtfData()
    Call InitConfig()
    
    Call LoaderShow("ETF Tracker: Loading data...")
    
    Dim logEnabled As Boolean
    logEnabled = IsLoggingEnabled()
    
    If logEnabled Then Call LogToSheet("INFO", "UpdateEtfData() started")
    
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oDataSheet As Object
    
    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    oDataSheet = GetSheetByName(oSheets, SHEET_DATA)
    
    If oDataSheet Is Nothing Then
        If logEnabled Then Call LogToSheet("ERROR", "Sheet """ & SHEET_DATA & """ not found!")
        Call LoaderClose()
        MsgBox "Error: sheet """ & SHEET_DATA & """ not found!", vbCritical, "Error"
        Exit Sub
    End If
    
    Dim lastCol As Long
    Dim tickers() As String
    Dim i As Long, j As Long
    Dim delayMs As Long
    
    lastCol = GetLastColumn(oDataSheet)
    If logEnabled Then Call LogToSheet("DEBUG", "GetLastColumn returned: " & lastCol)
    
    If lastCol < 1 Then
        If logEnabled Then Call LogToSheet("ERROR", "No tickers on row 1!")
        Call LoaderClose()
        MsgBox "Error: no tickers on row 1!", vbCritical, "Error"
        Exit Sub
    End If
    
    ReDim tickers(0 To lastCol - 1)
    Dim tickerCount As Long
    tickerCount = 0
    
    For i = 0 To lastCol - 1
        tickers(i) = Trim(GetCellString(oDataSheet, 0, i))
        If tickers(i) <> "" Then tickerCount = tickerCount + 1
        If logEnabled Then Call LogToSheet("DEBUG", "Ticker at column " & i & ": '" & tickers(i) & "'")
    Next i
    
    If tickerCount = 0 Then
        If logEnabled Then Call LogToSheet("ERROR", "No valid tickers on row 1!")
        Call LoaderClose()
        MsgBox "Error: no valid tickers on row 1!", vbCritical, "Error"
        Exit Sub
    End If
    
    delayMs = GetDelayMs()
    If logEnabled Then Call LogToSheet("INFO", "Delay: " & delayMs & "ms")
    
    Call LoaderSetText("Fetching data from Yahoo Finance...")
    
    Dim allData As Variant
    Dim maxLen As Long
    Dim result As Variant
    
    result = FetchAllTickers(tickers, logEnabled, delayMs)
    allData = result(0)
    maxLen = result(1)
    
    If logEnabled Then _
        Call LogToSheet("INFO", "Fetched data: " & (UBound(allData) + 1) & " columns, maxLen=" & maxLen)
    
    Call LoaderSetText("Processing data...")
    
    Dim hasData As Boolean
    hasData = False
    For i = LBound(allData) To UBound(allData)
        If IsArray(allData(i)) Then
            If UBound(allData(i)) >= 0 Then
                hasData = True
                Exit For
            End If
        End If
    Next i
    
    If Not hasData Then
        If logEnabled Then Call LogToSheet("ERROR", "No data was fetched!")
        Call LoaderClose()
        MsgBox "Error: no data was fetched! Check etf_logs.", vbCritical, "Error"
        Exit Sub
    End If
    
    Call ClearSheetArea(oDataSheet, 1, 0, maxLen, lastCol)
    
    Call LoaderSetText("Writing data to sheet...")
    
    Dim matrix() As Variant
    matrix = BuildMatrix(allData, maxLen, lastCol)
    
    ' Bulk write: setDataArray writes the entire 2D array at once
    ' Range: col=0..lastCol-1, row=1..maxLen (0-based: row 1 = second row in sheet)
    If UBound(matrix, 1) >= 0 And UBound(matrix, 2) >= 0 Then
        Dim range As Object
        Set range = oDataSheet.GetCellRangeByPosition(0, 1, lastCol - 1, maxLen)
        range.setDataArray(matrix)
    End If
    
    Call LoaderSetText("Done!")
    
    If logEnabled Then _
        Call LogToSheet("SUCCESS", "Data written successfully! " & _
                        lastCol & " tickers, " & maxLen & " rows.")
    
    Call WaitFor(300)
    Call LoaderClose()
    
    MsgBox "ETF data updated successfully! (" & lastCol & " tickers)", _
           vbInformation, "Done"
End Sub

' ============================================================
' HELPER FUNCTION (used from Calc cells)
' ============================================================

' Returns an array of historical close prices for charting.
' @customfunction
Function GetSparklineData(ByVal ticker As String, ByVal period As String) As Variant
    Dim colIndex As Long
    colIndex = FindTickerColumn(ticker)
    
    If colIndex = -1 Then
        GetSparklineData = Array(Array("N/A"))
        Exit Function
    End If
    
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oDataSheet As Object
    
    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    oDataSheet = GetSheetByName(oSheets, SHEET_DATA)
    
    If oDataSheet Is Nothing Then
        GetSparklineData = Array(Array("Error"))
        Exit Function
    End If
    
    Dim rowCount As Long
    rowCount = GetLastRow(oDataSheet) - 2
    If rowCount < 1 Then
        GetSparklineData = Array(Array("No data"))
        Exit Function
    End If
    
    Dim rowsNeeded As Long
    Select Case period
        Case "1w": rowsNeeded = Min2(5, rowCount)
        Case "1m": rowsNeeded = Min2(21, rowCount)
        Case Else: rowsNeeded = rowCount
    End Select
    
    Dim prices() As Double
    Dim r As Long
    Dim cell As Object
    
    ReDim prices(0 To rowsNeeded - 1)
    For r = 0 To rowsNeeded - 1
        Set cell = oDataSheet.GetCellByPosition(colIndex, r + 2)
        
        Dim v As Double
        v = cell.Value
        
        If v <> 0 Then
            prices(r) = v
        Else
            Dim txt As String
            txt = Trim(cell.getString())
            If txt <> "" And txt <> "N/A" And txt <> "Error" Then
                v = Val(txt)
                If v <> 0 Then
                    prices(r) = v
                Else
                    prices(r) = -1
                End If
            Else
                prices(r) = -1
            End If
        End If
    Next r
    
    Dim cleanPrices() As Double
    Dim cnt As Long
    cnt = 0
    
    For r = 0 To UBound(prices)
        If prices(r) >= 0 Then cnt = cnt + 1
    Next r
    
    If cnt = 0 Then
        GetSparklineData = Array(Array("No data"))
        Exit Function
    End If
    
    ReDim cleanPrices(0 To cnt - 1)
    cnt = 0
    For r = 0 To UBound(prices)
        If prices(r) >= 0 Then
            cleanPrices(cnt) = prices(r)
            cnt = cnt + 1
        End If
    Next r
    
    ' Reverse for chronological order (oldest first)
    Dim reversed() As Double
    ReDim reversed(0 To UBound(cleanPrices))
    For r = 0 To UBound(cleanPrices)
        reversed(r) = cleanPrices(UBound(cleanPrices) - r)
    Next r
    
    Dim outArr() As Variant
    ReDim outArr(0 To UBound(reversed), 0 To 0)
    For r = 0 To UBound(reversed)
        outArr(r, 0) = reversed(r)
    Next r
    
    GetSparklineData = outArr
End Function

' ============================================================
' INTERNAL HELPER FUNCTIONS
' ============================================================

' Finds the column for a given ticker in etf_data.
Private Function FindTickerColumn(ByVal ticker As String) As Long
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oDataSheet As Object
    
    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    oDataSheet = GetSheetByName(oSheets, SHEET_DATA)
    
    If oDataSheet Is Nothing Then
        FindTickerColumn = -1
        Exit Function
    End If
    
    Dim lastCol As Long
    lastCol = GetLastColumn(oDataSheet)
    If lastCol < 1 Then
        FindTickerColumn = -1
        Exit Function
    End If
    
    Dim i As Long
    Dim cellVal As String
    
    For i = 0 To lastCol - 1
        cellVal = Trim(GetCellString(oDataSheet, 0, i))
        If cellVal = ticker Then
            FindTickerColumn = i
            Exit Function
        End If
    Next i
    
    FindTickerColumn = -1
End Function

' Fetches data for all tickers from Yahoo Finance.
Private Function FetchAllTickers(ByRef tickers() As String, _
                                  ByVal logEnabled As Boolean, _
                                  ByVal delayMs As Long) As Variant
    Dim allData() As Variant
    Dim maxLen As Long
    Dim i As Long
    Dim colData As Variant
    Dim totalTickers As Long
    
    ReDim allData(0 To UBound(tickers))
    maxLen = 0
    
    ' Count real tickers (skip empty ones)
    totalTickers = 0
    For i = LBound(tickers) To UBound(tickers)
        If Trim(tickers(i)) <> "" Then totalTickers = totalTickers + 1
    Next i
    
    If totalTickers = 0 Then
        FetchAllTickers = Array(allData, 0)
        Exit Function
    End If
    
    Dim currentTicker As Long
    currentTicker = 0
    
    For i = LBound(tickers) To UBound(tickers)
        Dim ticker As String
        ticker = Trim(tickers(i))
        
        If ticker = "" Then
            allData(i) = Array("")
        Else
            currentTicker = currentTicker + 1
            
            Call LoaderSetText("Fetching " & ticker & " (" & currentTicker & "/" & totalTickers & ")...")
            
            If logEnabled Then _
                Call LogToSheet("FETCH", "Request to Yahoo for " & ticker & _
                                " (" & currentTicker & "/" & totalTickers & ")")
            
            colData = FetchSingleTicker(ticker, logEnabled)
            allData(i) = colData
            
            If IsArray(colData) Then
                If UBound(colData) + 1 > maxLen Then
                    maxLen = UBound(colData) + 1
                End If
            End If
        End If
        
        ' Rate limiting — pause before next request
        If i < UBound(tickers) And delayMs > 0 Then
            Call WaitFor(delayMs)
        End If
    Next i
    
    FetchAllTickers = Array(allData, maxLen)
End Function

' Fetches data for a single ticker from Yahoo Finance.
Private Function FetchSingleTicker(ByVal ticker As String, _
                                    ByVal logEnabled As Boolean) As Variant
    Dim endpoints As Variant
    Dim e As Long
    
    endpoints = Array( _
        YAHOO_ENDPOINTS(0) & EncodeUrl(ticker) & "?interval=1d&range=1y", _
        YAHOO_ENDPOINTS(1) & EncodeUrl(ticker) & "?interval=1d&range=1y" _
    )
    
    For e = 0 To UBound(endpoints)
        Dim url As String
        url = endpoints(e)
        
        Dim endpointName As String
        If InStr(url, "query1") > 0 Then
            endpointName = "query1.finance.yahoo.com"
        Else
            endpointName = "query2.finance.yahoo.com"
        End If
        
        ' HTTP GET
        Dim httpResponse As Variant
        httpResponse = HttpGet(url)
        
        ' Validate HTTP response
        If IsNull(httpResponse) Or Not IsArray(httpResponse) Then
            If logEnabled Then _
                Call LogToSheet("ERROR", ticker & " → " & endpointName & _
                                " HTTP request returned no result")
            GoTo ContinueEndpoint
        End If
        
        If UBound(httpResponse) < 1 Then
            If logEnabled Then _
                Call LogToSheet("ERROR", ticker & " → " & endpointName & _
                                " Incomplete HTTP response")
            GoTo ContinueEndpoint
        End If
        
        Dim httpCode As Long
        Dim responseText As String
        
        httpCode = CLng(httpResponse(0))
        responseText = CStr(httpResponse(1))
        
        If logEnabled Then _
            Call LogToSheet("HTTP", ticker & " → " & endpointName & " HTTP " & httpCode)
        
        ' Check for empty response
        If Len(responseText) = 0 Then
            If logEnabled Then _
                Call LogToSheet("WARN", ticker & " → " & endpointName & " empty response")
            GoTo ContinueEndpoint
        End If
        
        ' Check for HTML blocking page
        If Left(Trim(responseText), 5) = "<!DOC" Or Left(Trim(responseText), 5) = "<html" Then
            If logEnabled Then _
                Call LogToSheet("WARN", ticker & " → " & endpointName & " returned HTML (blocking?)")
            GoTo ContinueEndpoint
        End If
        
        If httpCode <> 200 Then
            If logEnabled Then _
                Call LogToSheet("WARN", ticker & " → HTTP " & httpCode & ", length: " & Len(responseText))
            GoTo ContinueEndpoint
        End If
        
        ' Parse JSON
        Dim priceArr As Variant
        priceArr = SimpleParseYahooJson(responseText)
        
        If IsNull(priceArr) Then
            If logEnabled Then _
                Call LogToSheet("WARN", ticker & " → invalid JSON or missing data from " & endpointName)
            GoTo ContinueEndpoint
        End If
        
        ' Check for empty array
        If Not IsArray(priceArr) Then
            If logEnabled Then _
                Call LogToSheet("WARN", ticker & " → empty array from " & endpointName)
            GoTo ContinueEndpoint
        End If
        
        If logEnabled Then _
            Call LogToSheet("OK", ticker & " → current price: " & priceArr(0) & _
                            ", historical records: " & (UBound(priceArr) - 1) & _
                            " (from " & endpointName & ")")
        
        FetchSingleTicker = priceArr
        Exit Function
        
ContinueEndpoint:
    Next e
    
    If logEnabled Then _
        Call LogToSheet("FAIL", ticker & " → all endpoints failed")
    FetchSingleTicker = Array("N/A")
End Function

' SIMPLE JSON PARSER for Yahoo Finance response.
Private Function SimpleParseYahooJson(ByVal text As String) As Variant
    On Error GoTo ParseError
    
    Dim priceStart As Long
    Dim currentPrice As Double
    Dim pricesStr As String
    
    ' 1. Find regularMarketPrice
    priceStart = InStr(text, """regularMarketPrice"":")
    If priceStart = 0 Then
        SimpleParseYahooJson = Null
        Exit Function
    End If
    
    priceStart = priceStart + Len("""regularMarketPrice"":")
    Dim tempStr As String
    tempStr = Mid(text, priceStart)
    tempStr = Trim(tempStr)
    
    ' Find the end of the number (comma, }, or end of string)
    Dim numEnd As Long
    numEnd = InStr(tempStr, ",")
    If numEnd = 0 Then numEnd = InStr(tempStr, "}")
    If numEnd = 0 Then numEnd = Len(tempStr) + 1
    
    ' Parse US number (with dot) manually
    Dim priceText As String
    priceText = Mid(tempStr, 1, numEnd - 1)
    currentPrice = ParseUsNumber(priceText)
    
    ' 2. Find the close array
    Dim closeStart As Long
    closeStart = InStr(text, """close"":[")
    If closeStart = 0 Then
        SimpleParseYahooJson = Null
        Exit Function
    End If
    
    closeStart = closeStart + Len("""close"":[")
    
    Dim closeEnd As Long
    closeEnd = InStr(closeStart, text, "]")
    If closeEnd = 0 Then
        SimpleParseYahooJson = Null
        Exit Function
    End If
    
    pricesStr = Mid(text, closeStart, closeEnd - closeStart)
    
    ' 3. Split by comma and filter nulls
    Dim parts() As String
    Dim cleanCount As Long
    Dim i As Long
    Dim val As Double
    
    parts = Split(pricesStr, ",")
    
    ' Count non-null values
    cleanCount = 0
    For i = 0 To UBound(parts)
        Dim p As String
        p = Trim(parts(i))
        If p <> "null" And p <> "" Then
            cleanCount = cleanCount + 1
        End If
    Next i
    
    If cleanCount = 0 Then
        SimpleParseYahooJson = Null
        Exit Function
    End If
    
    ' Create array
    Dim prices() As Variant
    ReDim prices(0 To cleanCount)
    prices(0) = currentPrice
    
    cleanCount = 1
    For i = 0 To UBound(parts)
        p = Trim(parts(i))
        If p <> "null" And p <> "" Then
            val = ParseUsNumber(p)
            prices(cleanCount) = val
            cleanCount = cleanCount + 1
        End If
    Next i
    
    ' Fail-safe check for zero/negative currentPrice
    If prices(0) <= 0 Then
        SimpleParseYahooJson = Null
        Exit Function
    End If
    
    ' Reverse historical prices (index 1 onwards)
    ' Yahoo: [oldest, ..., newest] -> we want [newest, ..., oldest]
    Dim reversed() As Variant
    ReDim reversed(0 To UBound(prices))
    reversed(0) = prices(0)
    
    Dim lastIdx As Long
    lastIdx = UBound(prices)
    For i = 1 To lastIdx
        reversed(i) = prices(lastIdx - i + 1)
    Next i
    
    SimpleParseYahooJson = reversed
    Exit Function
    
ParseError:
    SimpleParseYahooJson = Null
End Function

' Parses a US-format number (with dot) manually, locale-independent.
' Example: "704.6" -> 704.6, "1234" -> 1234
Private Function ParseUsNumber(ByVal text As String) As Double
    Dim result As Double
    Dim intPart As String
    Dim fracPart As String
    Dim dotPos As Long
    Dim i As Long
    Dim isNegative As Boolean
    
    text = Trim(text)
    If text = "" Then
        ParseUsNumber = 0
        Exit Function
    End If
    
    ' Handle negative numbers
    isNegative = False
    If Left(text, 1) = "-" Then
        isNegative = True
        text = Mid(text, 2)
    End If
    
    dotPos = InStr(text, ".")
    
    If dotPos = 0 Then
        ' Integer — Val() works fine
        result = Val(text)
    Else
        intPart = Mid(text, 1, dotPos - 1)
        fracPart = Mid(text, dotPos + 1)
        
        result = Val(intPart)
        
        Dim fracVal As Double
        Dim divisor As Double
        fracVal = Val(fracPart)
        divisor = 1
        For i = 1 To Len(fracPart)
            divisor = divisor * 10
        Next i
        
        If fracVal >= 0 Then
            result = result + (fracVal / divisor)
        End If
    End If
    
    If isNegative Then result = -result
    
    ParseUsNumber = result
End Function

' BuildMatrix — transposes columns into rows.
Private Function BuildMatrix(ByRef allData As Variant, ByVal maxLen As Long, _
                              ByVal numCols As Long) As Variant
    Dim matrix() As Variant
    Dim r As Long, c As Long
    Dim colArr As Variant
    
    If maxLen <= 0 Or numCols <= 0 Then
        Call LogToSheet("ERROR", "BuildMatrix: invalid dimensions maxLen=" & maxLen & ", numCols=" & numCols)
        BuildMatrix = Array()
        Exit Function
    End If
    
    ReDim matrix(0 To maxLen - 1, 0 To numCols - 1)
    
    For c = 0 To numCols - 1
        colArr = allData(c)
        
        If IsArray(colArr) Then
            For r = 0 To maxLen - 1
                If r <= UBound(colArr) Then
                    matrix(r, c) = colArr(r)
                Else
                    matrix(r, c) = ""
                End If
            Next r
        Else
            For r = 0 To maxLen - 1
                matrix(r, c) = ""
            Next r
        End If
    Next c
    
    BuildMatrix = matrix
End Function

' ============================================================
' HTTP REQUESTS (via MSXML2.XMLHTTP for Windows)
' ============================================================

' Executes an HTTP GET request.
Private Function HttpGet(ByVal url As String) As Variant
    Dim http As Object
    Dim result(0 To 1) As Variant
    
    result(0) = 0
    result(1) = ""
    
    On Error GoTo HttpError
    
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    If http Is Nothing Then
        result(1) = "HTTP Error: Cannot create MSXML2.XMLHTTP.6.0"
        HttpGet = result
        Exit Function
    End If
    
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", USER_AGENT
    
    ' 30-second timeout
    On Error Resume Next
    http.SetTimeouts 30000, 30000, 30000, 30000
    On Error GoTo HttpError
    
    http.Send ""
    
    result(0) = http.Status
    result(1) = http.ResponseText
    
    On Error Resume Next
    Set http = Nothing
    On Error GoTo 0
    
    HttpGet = result
    Exit Function
    
HttpError:
    result(0) = 0
    result(1) = "HTTP Error: " & Err.Description
    
    On Error Resume Next
    Set http = Nothing
    On Error GoTo 0
    
    HttpGet = result
End Function

' ============================================================
' LibreOffice Calc HELPER FUNCTIONS
' ============================================================

' Returns a sheet object by name.
Private Function GetSheetByName(ByRef sheets As Object, ByVal name As String) As Object
    Dim i As Long
    For i = 0 To sheets.Count - 1
        If sheets(i).Name = name Then
            Set GetSheetByName = sheets(i)
            Exit Function
        End If
    Next i
    Set GetSheetByName = Nothing
End Function

' Returns the string value of a cell.
Private Function GetCellString(ByRef sheet As Object, ByVal row As Long, _
                                ByVal col As Long) As String
    Dim cell As Object
    Set cell = sheet.GetCellByPosition(col, row)
    
    On Error Resume Next
    GetCellString = cell.getFormula()
    If Err.Number <> 0 Or GetCellString = "" Then
        Err.Clear()
        GetCellString = cell.getString()
    End If
    If Err.Number <> 0 Or GetCellString = "" Then
        Err.Clear()
        GetCellString = cell.String
    End If
End Function

' Sets a cell value (number or text).
Private Sub SetCellValue(ByRef sheet As Object, ByVal row As Long, _
                          ByVal col As Long, ByVal value As Variant)
    Dim cell As Object
    Set cell = sheet.GetCellByPosition(col, row)
    
    ' Skip null/empty values (leave cell empty)
    If IsNull(value) Or IsEmpty(value) Then
        Exit Sub
    End If
    
    If IsNumeric(value) Then
        cell.Value = CDbl(value)
    Else
        cell.String = CStr(value)
    End If
End Sub

' Returns the last used column (1-based, 0 if empty).
Private Function GetLastColumn(ByRef sheet As Object) As Long
    Dim cursor As Object
    Set cursor = sheet.CreateCursor()
    cursor.GotoEndOfUsedArea(False)
    GetLastColumn = cursor.RangeAddress.EndColumn + 1
End Function

' Returns the last used row (1-based, 0 if empty).
Private Function GetLastRow(ByRef sheet As Object) As Long
    Dim cursor As Object
    Set cursor = sheet.CreateCursor()
    cursor.GotoEndOfUsedArea(False)
    GetLastRow = cursor.RangeAddress.EndRow + 1
End Function

' Clears a rectangular area in the sheet.
Private Sub ClearSheetArea(ByRef sheet As Object, ByVal row As Long, _
                            ByVal col As Long, ByVal height As Long, _
                            ByVal width As Long)
    If height <= 0 Or width <= 0 Then Exit Sub
    
    ' Limit to actual number of rows in the sheet
    Dim maxSheetRows As Long
    maxSheetRows = sheet.Rows.Count
    
    If row + height > maxSheetRows Then
        height = maxSheetRows - row
    End If
    If height <= 0 Then Exit Sub
    
    Dim range As Object
    Set range = sheet.GetCellRangeByPosition(col, row, _
                col + width - 1, row + height - 1)
    range.ClearContents(7)
End Sub

' Minimum value helper.
Private Function Min2(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then Min2 = a Else Min2 = b
End Function

' URL encoding helper.
Private Function EncodeUrl(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace(result, " ", "%20")
    result = Replace(result, "+", "%2B")
    EncodeUrl = result
End Function

' ============================================================
' LOGGING TO "etf_logs" SHEET
' ============================================================

' Adds a log entry to the etf_logs sheet.
Private Sub LogToSheet(ByVal level As String, ByVal msg As String)
    On Error GoTo LogError
    
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oLogSheet As Object
    
    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    oLogSheet = GetSheetByName(oSheets, SHEET_LOGS)
    
    ' Create the sheet if it doesn't exist
    If oLogSheet Is Nothing Then
        oLogSheet = oSheets.insertNewByName(SHEET_LOGS, 0)
        SetCellValue oLogSheet, 0, 0, "Timestamp"
        SetCellValue oLogSheet, 0, 1, "Level"
        SetCellValue oLogSheet, 0, 2, "Message"
    End If
    
    Dim lastRow As Long
    lastRow = GetLastRow(oLogSheet)
    
    Dim nowStr As String
    nowStr = Format(Now(), "DD.MM.YYYY HH:MM:SS")
    
    SetCellValue oLogSheet, lastRow, 0, nowStr
    SetCellValue oLogSheet, lastRow, 1, level
    SetCellValue oLogSheet, lastRow, 2, msg
    
    Exit Sub
    
LogError:
End Sub

' ============================================================
' UTILITY FUNCTIONS
' ============================================================

' Pause execution for ms milliseconds.
Private Sub WaitFor(ByVal ms As Long)
    If ms <= 0 Then Exit Sub
    
    Dim startTime As Double
    startTime = Timer()
    Do While (Timer() - startTime) * 1000 < ms
        DoEvents
    Loop
End Sub

' ============================================================
' TEST
' ============================================================

' Test function.
Sub TestConfiguration()
    Call InitConfig()
    
    Call LogToSheet("TEST", "=== CONFIGURATION TEST ===")
    Call LogToSheet("TEST", "SHEET_DATA: " & SHEET_DATA)
    Call LogToSheet("TEST", "SHEET_LOGS: " & SHEET_LOGS)
    Call LogToSheet("TEST", "SHEET_SETTINGS: " & SHEET_SETTINGS)
    
    Dim logEnabled As Boolean
    logEnabled = IsLoggingEnabled()
    Call LogToSheet("TEST", "Logging enabled (settings!C2): " & IIf(logEnabled, "Yes", "No"))
    
    Dim delayMs As Long
    delayMs = GetDelayMs()
    Call LogToSheet("TEST", "Delay (settings!C3): " & delayMs & " ms")
    
    ' Test the Loader
    Call LoaderShow("ETF Tracker: Test Loading...")
    Call WaitFor(1500)
    Call LoaderSetText("Still working...")
    Call WaitFor(1500)
    Call LoaderSetText("Almost done!")
    Call WaitFor(1000)
    Call LoaderClose()
    
    MsgBox "Test completed. Check sheet """ & SHEET_LOGS & """.", _
           vbInformation, "Test"
End Sub