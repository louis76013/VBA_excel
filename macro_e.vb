 Public Enum Name
    AvePrice = 1
    ProfitAmout
    BColor
    BPercent
    BAmout
    BJColor
    BJAmout
    Setx
    ProfitPercent
    LivePercent
    RestaurantPercent
    ProcurementPrice
    AgentPercent
    AgentAmout
    AgentComboPercent
    AgentComboAmount
    Supplier
    Product
    Spec
    Weight
    MRSP
    RestaurantPrice
    LivePrice
    BJPercent
    BAdjust
    LAST
End Enum

Public strColor(8) As String
Public color(8) As Integer
Public strArray(LAST) As String
Public colArray(LAST) As Integer
Public valArray(LAST) As Double



Function initArray()
    strColor(1) = "紅"
    strColor(2) = "粉"
    strColor(3) = "黃"
    strColor(4) = "綠"
    strColor(5) = "藍"
    strColor(6) = "黑"
    strColor(7) = "紫"
    strColor(8) = "白"
    color(1) = 3
    color(2) = 22
    color(3) = 44
    color(4) = 43
    color(5) = 37
    color(6) = 1
    color(7) = 29
    color(8) = 2
    
    strArray(AvePrice) = "均單價"
    strArray(ProfitAmout) = "商品毛利"
    strArray(BColor) = "B色標"
    strArray(BPercent) = "B概%"
    strArray(BAmout) = "B利潤"
    strArray(BJColor) = "BJ色標"
    strArray(BJAmout) = "BJ抽成"
    strArray(Setx) = "組數"
    strArray(ProfitPercent) = "公司利潤%"
    strArray(LivePercent) = "直播折扣%"
    strArray(RestaurantPercent) = "餐?折扣%"
    strArray(ProcurementPrice) = "成本"
    strArray(AgentPercent) = "經紀單%"
    strArray(AgentAmout) = "經紀單"
    strArray(AgentComboPercent) = "經紀組合%"
    strArray(AgentComboAmount) = "經紀組合"
    strArray(Supplier) = "廠商"
    strArray(Product) = "品名"
    strArray(Spec) = "規格"
    strArray(Weight) = "重量"
    strArray(MRSP) = "市價"
    strArray(RestaurantPrice) = "餐?價"
    strArray(LivePrice) = "直播價"
    strArray(BAdjust) = "B抓%"
    strArray(BJPercent) = "BJ抓%"
End Function

Function indexCol()
    Dim i As Integer
    Dim strSearch As String
    Dim aCell As Range
    For i = 1 To LAST - 1
    strSearch = strArray(i)
        Set aCell = Sheet1.rows(1).Find(What:=strSearch, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)

        If Not aCell Is Nothing Then
        '    MsgBox aCell.Value & "Column Number is " & aCell.Column
            colArray(i) = aCell.Column
        End If
    Next i
    
End Function

Function initArg()
    Dim nRow, i As Integer
    Dim basePrice As Double
 
    nRow = Cells(rows.Count, 1).End(xlUp).Row
   
    For i = 2 To nRow
        valArray(Setx) = Cells(i, colArray(Setx)).Value
        If valArray(Setx) = 0 Then
            valArray(Setx) = 1
            Cells(i, colArray(Setx)).Value = 1
        End If
        If valArray(Setx) = 1 Then
            valArray(ProcurementPrice) = Cells(i, colArray(ProcurementPrice)).Value
            basePrice = valArray(ProcurementPrice)
        Else
            valArray(ProcurementPrice) = basePrice * valArray(Setx)
            Cells(i, colArray(ProcurementPrice)).Value = valArray(ProcurementPrice)
        End If
        Cells(i, colArray(ProcurementPrice)).NumberFormat = "[$NT$]#,###"
    
        If Cells(i, colArray(ProfitPercent)).Value = Empty Then
            If valArray(Setx) = 1 Then
                valArray(ProfitPercent) = 0.3
            ElseIf valArray(Setx) >= 5 Then
                valArray(ProfitPercent) = 0.21
            ElseIf valArray(Setx) >= 3 Then
                valArray(ProfitPercent) = 0.24
            End If
            Cells(i, colArray(ProfitPercent)).Value = valArray(ProfitPercent)
        Else
            valArray(ProfitPercent) = Cells(i, colArray(ProfitPercent)).Value
        End If
        Cells(i, colArray(ProfitPercent)).NumberFormat = "0%"
        
        If Cells(i, colArray(LivePercent)).Value = Empty Then
            If valArray(Setx) = 1 Then
                valArray(LivePercent) = 0.88
            ElseIf valArray(Setx) >= 5 Then
                valArray(LivePercent) = 0.76
            ElseIf valArray(Setx) >= 3 Then
                valArray(LivePercent) = 0.79
            End If
            Cells(i, colArray(LivePercent)).Value = valArray(LivePercent)
        Else
            valArray(LivePercent) = Cells(i, colArray(LivePercent)).Value
        End If
        Cells(i, colArray(LivePercent)).NumberFormat = "0%"
    
        If Cells(i, colArray(RestaurantPercent)).Value = Empty Then
            If valArray(Setx) = 1 Then
                valArray(RestaurantPercent) = 0.94
            ElseIf valArray(Setx) >= 5 Then
                valArray(RestaurantPercent) = 0.9
            ElseIf valArray(Setx) >= 3 Then
                valArray(RestaurantPercent) = 0.92
            End If
            Cells(i, colArray(RestaurantPercent)).Value = valArray(RestaurantPercent)
        Else
            valArray(RestaurantPercent) = Cells(i, colArray(RestaurantPercent)).Value
        End If
        Cells(i, colArray(RestaurantPercent)).NumberFormat = "0%"
    Next i
End Function
   

Sub AutoFill()
   
    Dim nRow As Integer
    Dim i As Integer
    Dim iColor As Integer
    Dim fontcolor As Long
    
    initArray
    indexCol
    initArg
  
    nRow = Cells(rows.Count, 1).End(xlUp).Row
    
    For i = 2 To nRow
        valArray(Setx) = Cells(i, colArray(Setx)).Value
        valArray(ProfitPercent) = Cells(i, colArray(ProfitPercent)).Value
        valArray(LivePercent) = Cells(i, colArray(LivePercent)).Value
        valArray(RestaurantPercent) = Cells(i, colArray(RestaurantPercent)).Value
        valArray(ProcurementPrice) = Cells(i, colArray(ProcurementPrice)).Value
        
        valArray(LivePrice) = valArray(ProcurementPrice) / (1 - valArray(ProfitPercent)) '直播價=成本價/(1-公司利潤%)
        valArray(AvePrice) = valArray(LivePrice) / valArray(Setx)   '均單價=直播價/組數
        valArray(MRSP) = valArray(LivePrice) / (valArray(LivePercent)) '市價=直播價/(直播折扣%)
        valArray(RestaurantPrice) = valArray(MRSP) * valArray(RestaurantPercent) '餐?價=市價*餐直播折扣%
        valArray(ProfitAmout) = valArray(LivePrice) - valArray(ProcurementPrice) '商品毛利=直播價-成本
        valArray(BPercent) = valArray(ProfitPercent) / 2  'B概%=公司利潤% / 2
        valArray(BAdjust) = (valArray(BPercent) * 100 \ 5) * 5 / 100 'B抓% = B概% 無條件捨去
        valArray(BAmout) = valArray(LivePrice) * valArray(BAdjust) 'B利潤=直播價*B抓%
        valArray(BJPercent) = valArray(BAdjust) - 0.05 'BJ抓%=B抓%-5
        valArray(BJAmout) = valArray(LivePrice) * valArray(BJPercent) 'BJ抽成=直播價*BJ抓%
        valArray(BColor) = valArray(BAdjust) * 100 / 5 'B色標=B摡%\5
        valArray(BJColor) = valArray(BColor) 'BJ色標=B色標
        
        Cells(i, colArray(MRSP)).Value = valArray(MRSP)
        Cells(i, colArray(MRSP)).NumberFormat = "[$NT$]#,###"
        Cells(i, colArray(LivePrice)).Value = valArray(LivePrice)
        Cells(i, colArray(LivePrice)).NumberFormat = "[$NT$]#,###"
        Cells(i, colArray(RestaurantPrice)).Value = valArray(RestaurantPrice)
        Cells(i, colArray(RestaurantPrice)).NumberFormat = "[$NT$]#,###"
        Cells(i, colArray(AvePrice)).Value = valArray(AvePrice)
        Cells(i, colArray(AvePrice)).NumberFormat = "[$NT$]#,###"
        Cells(i, colArray(ProfitAmout)).Value = valArray(ProfitAmout)
        Cells(i, colArray(ProfitAmout)).NumberFormat = "[$NT$]#,###"
        Cells(i, colArray(BPercent)).Value = valArray(BPercent)
        Cells(i, colArray(BPercent)).NumberFormat = "0%"
        Cells(i, colArray(BAdjust)).Value = valArray(BAdjust)
        Cells(i, colArray(BAdjust)).NumberFormat = "0%"
        Cells(i, colArray(BAmout)).Value = valArray(BAmout)
        Cells(i, colArray(BAmout)).NumberFormat = "[$NT$]#,###"
        Cells(i, colArray(BJPercent)).Value = valArray(BJPercent)
        Cells(i, colArray(BJPercent)).NumberFormat = "0%"
        Cells(i, colArray(BJAmout)).Value = valArray(BJAmout)
        Cells(i, colArray(BJAmout)).NumberFormat = "[$NT$]#,###"
        iColor = valArray(BColor)
        If iColor > 8 Then
            iColor = 8
        End If
        If iColor < 1 Then
            iColor = 1
        End If
    
        If iColor = 8 Then
            fontcolor = vbBlack
        Else
            fontcolor = vbWhite
        End If
    
        Cells(i, colArray(BColor)).Value = strColor(iColor) & (valArray(BAdjust) * 100) & "%"
        Cells(i, colArray(BColor)).Font.color = fontcolor
        Cells(i, colArray(BColor)).Interior.ColorIndex = color(iColor)
'       Cells(iR, colArray(BColor)).Font.Size = 16
            
        Cells(i, colArray(BJColor)).Value = strColor(iColor) & (valArray(BJPercent) * 100) & "%"
        Cells(i, colArray(BJColor)).Font.color = fontcolor
        Cells(i, colArray(BJColor)).Interior.ColorIndex = color(iColor)
'       Cells(iR, colArray(BJColor)).Font.Size = 16
    Next i
 
End Sub

Sub Dummy()
    Sheet1.Select
    
    Dim nRow As Integer
    Dim randn As Double
    Dim arrX As Variant
    Dim iR As Integer
    Dim itemNum As Integer
    Dim List(15) As Integer
    Dim success, duplicate As Boolean
    initArray
    indexCol
    
    nRow = Cells(rows.Count, 1).End(xlUp).Row
     
    arrX = ArrayFromCSV("d:\vba\sample_products.csv")
    
    itemNum = 10 + Rnd * 6  'to pick 10-15 items
    For i = 1 To itemNum
        success = False
        While success = False
            randn = Rnd * 200
            duplicate = False
            For k = 1 To i - 1
                If randn = List(k) Then
                    duplicate = True
                End If
            Next k
            If duplicate = False Then
                success = True
                List(i) = randn
            End If
        Wend
    Next i
    
    For i = 1 To itemNum
        For j = 1 To 3
        
            iR = nRow + (i - 1) * 3 + j
           ' Cells(iR, colArray(Supplier)).NumberFormat = "@"
            Cells(iR, colArray(Supplier)).Value = "dummy"
'            Cells(iR, cB).Font.Size = 16
          
            If j = 1 Then
                Cells(iR, colArray(Setx)).Value = 1
                Cells(iR, colArray(Product)).Value = arrX(List(i), 2) '商品名稱
                Cells(iR, colArray(ProfitPercent)).Value = 0.3 + Rnd * 50 / 100
                Cells(iR, colArray(LivePercent)).Value = 0.85 + Rnd * 5 / 100
                Cells(iR, colArray(RestaurantPercent)).Value = 0.93 + Rnd * 4 / 100
                Cells(iR, colArray(ProcurementPrice)).Value = arrX(List(i), 3)
                basePrice = arrX(List(i), 3)
            End If

            If j = 2 Then
                Cells(iR, colArray(Setx)).Value = 3 + Int(Rnd * 2)
                
                Cells(iR, colArray(ProfitPercent)).Value = 0.22 + Rnd * 4 / 100
                Cells(iR, colArray(LivePercent)).Value = 0.75 + Rnd * 5 / 100
                Cells(iR, colArray(RestaurantPercent)).Value = 0.9 + Rnd * 3 / 100
            End If
            If j = 3 Then
                Cells(iR, colArray(Setx)).Value = 5 + Int(Rnd * 6)
                Cells(iR, colArray(ProfitPercent)).Value = 0.18 + Rnd * 4 / 100
                Cells(iR, colArray(LivePercent)).Value = 0.72 + Rnd * 3 / 100
                Cells(iR, colArray(RestaurantPercent)).Value = 0.9 + Rnd * 3 / 100
            End If
            
            Cells(iR, colArray(ProcurementPrice)).NumberFormat = "[$NT$]#,###"
            Cells(iR, colArray(ProfitPercent)).NumberFormat = "0%"
            Cells(iR, colArray(LivePercent)).NumberFormat = "0%"
            Cells(iR, colArray(RestaurantPercent)).NumberFormat = "0%"
       
        Next
    Next
End Sub

Sub DeleteRows()
    Dim i As Integer
    Dim nRow As Integer
    Dim str As String
    initArray
    indexCol
    nRow = Cells(rows.Count, 1).End(xlUp).Row
    For i = nRow To 1 Step -1
        str = Cells(i, colArray(Supplier)).Value
        If (str = "dummy") Then
            rows(i).EntireRow.Delete
        End If
    Next
End Sub
   
'VBA function to open a CSV file in memory and parse it to a 2D
'array without ever touching a worksheet:

Function ArrayFromCSV(sFile$)
    Dim c&, i&, j&, p&, d$, s$, rows&, cols&, a, r, v
    Const Q = """", QQ = Q & Q
    Const ENQ = ""  'Chr(5)
    Const ESC = ""  'Chr(27)
    Const COM = ","
    
    d = OpenTextFile$(sFile)
    If LenB(d) Then
        r = Split(Trim(d), vbCrLf)
        rows = UBound(r) + 1
        cols = UBound(Split(r(0), ",")) + 1
        ReDim v(1 To rows, 1 To cols)
        For i = 1 To rows
            s = r(i - 1)
            If LenB(s) Then
                If InStrB(s, QQ) Then s = Replace(s, QQ, ENQ)
                For p = 1 To Len(s)
                    Select Case Mid(s, p, 1)
                        Case Q:   c = c + 1
                        Case COM: If c Mod 2 Then Mid(s, p, 1) = ESC
                    End Select
                Next
                If InStrB(s, Q) Then s = Replace(s, Q, "")
                a = Split(s, COM)
                For j = 1 To cols
                    s = a(j - 1)
                    If InStrB(s, ESC) Then s = Replace(s, ESC, COM)
                    If InStrB(s, ENQ) Then s = Replace(s, ENQ, Q)
                    v(i, j) = s
                Next
            End If
        Next
        ArrayFromCSV = v
    End If
End Function

Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
        .Close
    End With
End Function
