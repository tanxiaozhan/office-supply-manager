Attribute VB_Name = "modFunction"
'盘点
Public Sub balance()
    Dim txtMsg As String
    Dim balanceDate As String '盘点日期
    Dim lastBalanceDate, lastDate As String '上次盘点日期
    Dim lastStockSaledate As String
    Dim number As Double     '数量
    Dim price As Double     '金额
    Dim lastNumber As Double     '上期数量
    Dim LastPrice As Double     '上期金额
    Dim numberStock As Double     '进货数量
    Dim priceStock As Double     '进货金额
    Dim numberSale As Double     '销售数量
    Dim priceSale As Double     '销售金额
    Dim priceGoods As Single   '商品单价
    Dim profit As Double       '本期盈余
    
    Dim rs As ADODB.Recordset
    Dim rsgoods As ADODB.Recordset
    Dim rsChain As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    DBConnect
    rs.Open "select top 1 * from balanceDate order by bid desc", Conn, 1, 1
    lastBalanceDate = ""
    If rs.RecordCount > 0 Then
        lastDate = rs("bDate")
        txtMsg = "上次盘点日期是：" & rs("bDate") & "，确认进行盘点吗？"
        lastBalanceDate = " and bDate =#" & rs("bDate") & "#"
        lastStockSaledate = " and sDate>#" & rs("bDate") & "#"
    Else
        txtMsg = "确认进行盘点吗？"
    End If
    rs.Close
        
    If MsgBox(txtMsg, vbInformation + vbYesNo, "盘点结存") = vbNo Then Exit Sub
    
    balanceDate = InputBox("请输入盘点日期", "盘点日期", CStr(Date))
    If Not IsDate(balanceDate) Then
        If balanceDate = "" Then Exit Sub
        MsgBox "错误的日期格式", vbCritical, "盘点结存"
        Exit Sub
    End If
    
    If lastDate <> "" Then
        If CDate(balanceDate) <= CDate(lastDate) Then
            MsgBox "输入的日期早于或等于上次盘点日期，不能进行盘点操作！", vbCritical, "错误"
            Exit Sub
        Else
            If CDate(balanceDate) > CDate(Format(Now, "yyyy-mm-dd")) Then
                MsgBox "输入的日期大于当前日期，不能进行盘点操作！", vbCritical, "错误"
                Exit Sub
            End If
        
        End If
    End If
    
    
    currentDate = Now
    
    Set rsgoods = New ADODB.Recordset
    rsgoods.Open "select * from goods", Conn, 1, 1
        
    Set rsChain = New ADODB.Recordset
    rsChain.Open "select * from chain", Conn, adOpenDynamic, adLockBatchOptimistic
    Do While Not rsChain.EOF
        
        rsgoods.MoveFirst
        Do While Not rsgoods.EOF
            
            strSQL = "select bCurrent,bCurrentPrice from balance where bChain=" & rsChain("cid") & " and bGoods=" & rsgoods("gid") & lastBalanceDate
            rs.Open strSQL, Conn, 1, 1   '获取库存
            If rs.RecordCount = 1 Then
                lastNumber = IIf(IsNull(rs("bCurrent")), 0, rs("bCurrent"))
                LastPrice = IIf(IsNull(rs("bCurrentPrice")), 0, rs("bCurrentPrice"))
                libIsnullandStockIsnull = IsNull(rs("bcurrent"))
            End If
            rs.Close
         
            strSQL = "select sum(iif(isnull(sNumber),0,sNumber)) as StockNumber,sum(iif(isnull(sTotal),0,sTotal)) as StockTotal from stock where sChain=" & rsChain("cid") & " and  sGoods=" & rsgoods("gid") & lastStockSaledate
            rs.Open strSQL, Conn, 1, 1
            numberStock = IIf(IsNull(rs("StockNumber")), 0, rs("stocknumber"))
            priceStock = IIf(IsNull(rs("StockTotal")), 0, rs("stocktotal"))
            number = lastNumber + numberStock
            price = LastPrice + priceStock
            rs.Close
            
            If number = 0 Then GoTo continue    '商品的库存和进货为零，则无而盘点该商品
            
            priceGoods = price / number     '计算商品平均价格（单价）
            
         
            strSQL = "select sum(iif(isNull(sNumber),0,sNumber)) as SaleNumber,sum(iif(isnull(sTotal),0,sTotal)) as SaleTotal from sale where sChain=" & rsChain("cid") & " and  sGoods=" & rsgoods("gid") & lastStockSaledate
            rs.Open strSQL, Conn, 1, 1
            numberSale = IIf(IsNull(rs("SaleNumber")), 0, rs("SaleNumber"))
            priceSale = IIf(IsNull(rs("SaleTotal")), 0, rs("SaleTotal"))
            number = number - numberSale
            
            rs.Close
            
            profit = priceSale - price + number * priceGoods
                 
            strSQL = "insert into balance(bChain,bGoods,bStock,bStockPrice,bSale,bSalePrice,bLogicalCurrent,bCurrent,bCurrentPrice,bLast,bLastPrice,bPrice,bProfit,bDate) " & _
                             "values(" & rsChain("cid") & "," & rsgoods("gid") & "," & numberStock & "," & priceStock & "," & _
                             numberSale & "," & priceSale & "," & number & "," & number & "," & price & "," & lastNumber & "," & LastPrice & "," & priceGoods & "," & profit & ",'" & balanceDate & " " & Format(currentDate, "hh:mm:ss") & "')"
            Conn.Execute strSQL
                             
continue:
            
            rsgoods.MoveNext
        Loop
        
        
        rsChain.MoveNext
    
    Loop
    
    rsgoods.Close
    Set rsgoods = Nothing
    rsChain.Close
    Set rsChain = Nothing

    Conn.Execute "insert into balanceDate(bDate,bDescript) values('" & balanceDate & " " & Format(currentDate, "hh:mm:ss") & "','用户[" & curUserName & "]于" & currentDate & "盘点！')"

    MsgBox "盘点完毕。", vbInformation, "盘点结存"
    frmBalance.Show
End Sub


Public Function PrintListView(ByRef pobjListView As ListView, pstrHeading As String) As Boolean
    Dim objCol As ColumnHeader
    Dim objLI As ListItem
    Dim objILS As ImageList
    Dim objPic As Picture
    
    Dim dblXScale As Double
    Dim dblYScale As Double
    Dim sngFontSize As Single
    Dim lngX As Long
    Dim lngY As Long
    Dim lngX1 As Long
    Dim lngY1 As Long
    Dim lngX2 As Long
    Dim lngRows As Long
    Dim lngLeft As Long
    Dim lngPageNo As Long
    Dim lngEOP As Long
    Dim lngEnd As Long
    Dim lngWidth As Long
    Dim intCols As Integer
    Dim lngTop As Long
    Dim intOffset As Integer
    Dim px As Integer
    Dim py As Integer
    Dim intRowHeight As Integer
    Dim strText As String
    Dim strTextTrun As String
    
    '--------------------------------------------------------------------------
    'Establish print & screen metrics
    '--------------------------------------------------------------------------
    
    On Error GoTo Error_Handler
    
    Screen.MousePointer = vbHourglass
        
    For Each objCol In pobjListView.ColumnHeaders
        
        lngX = lngX + objCol.Width
    
    Next
    
    Set objILS = pobjListView.SmallIcons
    
    dblXScale = (Printer.Width * 0.9) / lngX
    dblYScale = Printer.Height / pobjListView.Height
    
    lngLeft = (Printer.Width - (Printer.Width * 0.95)) / 2
    
    sngFontSize = Printer.Font.Size
    
    If pstrHeading <> "" Then
    
        Printer.Font.Size = 12
        Printer.CurrentX = (Printer.Width / 2) - (Printer.TextWidth(pstrHeading) / 2)
        Printer.Font.Underline = True
        Printer.Print pstrHeading
        Printer.Font.Underline = False
        Printer.Font.Size = sngFontSize
        lngTop = Printer.CurrentY + Printer.CurrentY
        
    End If
    
    intRowHeight = (Screen.TwipsPerPixelY * 17)
    
    lngEOP = Printer.Height - (intRowHeight * 3)
    
    lngX = lngLeft
    lngY = lngTop
    
    lngY1 = lngTop + (Screen.TwipsPerPixelY * 17)
    
    Printer.CurrentY = lngY
    Printer.Font.Bold = True
    Printer.DrawMode = vbCopyPen
       
    px = Screen.TwipsPerPixelX
    py = Screen.TwipsPerPixelY
    
    '--------------------------------------------------------------------------
    'Print column headers with slight 3D effect
    '--------------------------------------------------------------------------
    
    For Each objCol In pobjListView.ColumnHeaders
        
        lngX1 = lngX + (objCol.Width * dblXScale)
        
        Printer.Line (lngX, lngY)-(lngX1, lngY1), vbButtonShadow, BF
        Printer.Line (lngX, lngY)-(lngX1 - px, lngY1), RGB(245, 245, 245), BF
        Printer.Line (lngX + px, lngY + py)-(lngX1, lngY1), vbButtonShadow, BF
        Printer.Line (lngX + px, lngY + py)-(lngX1 - px, lngY1 - py), vbButtonFace, BF
        
        Printer.CurrentY = lngY + ((intRowHeight - Printer.TextHeight(objCol.Text)) / 2) + py
        
        Select Case objCol.Alignment
               
            Case ListColumnAlignmentConstants.lvwColumnCenter
                   
                Printer.CurrentX = lngX + (((objCol.Width * dblXScale) - Printer.TextWidth(objCol.Text)) / 2)
               
            Case ListColumnAlignmentConstants.lvwColumnLeft
                
                Printer.CurrentX = lngX + (px * 5)
            
            Case ListColumnAlignmentConstants.lvwColumnRight
                
                Printer.CurrentX = lngX + ((objCol.Width * dblXScale) - Printer.TextWidth(objCol.Text)) - (px * 5)
                
        End Select
        
        Printer.Print objCol.Text
           
        lngX = lngX1
    
    Next
    
    lngEnd = lngX1 + px
    
    Printer.Font.Bold = False
    
    '--------------------------------------------------------------------------
    'Print list item data
    '--------------------------------------------------------------------------
    
    For Each objLI In pobjListView.ListItems
        
        If lngY1 > lngEOP - intRowHeight - intRowHeight Then
            
            '------------------------------------------------------------------
            'Print page number
            '------------------------------------------------------------------
            
            lngPageNo = lngPageNo + 1
            Printer.CurrentX = (Printer.Width / 2) - (Printer.TextWidth("Page " & lngPageNo) / 2)
            Printer.CurrentY = lngEOP - intRowHeight
            Printer.Print "Page " & lngPageNo
            Printer.NewPage
            Printer.CurrentY = lngTop
            lngY = lngTop
        
        Else
        
            lngY = lngY + intRowHeight
        
        End If
        
        lngX = lngLeft
        
        lngY1 = lngY + intRowHeight
            
        For Each objCol In pobjListView.ColumnHeaders
            
            '------------------------------------------------------------------
            'Print the icon if on col 1
            '------------------------------------------------------------------
            
            If objCol.Index > 1 Then
                
                strText = objLI.SubItems(objCol.Index - 1)
                
                intOffset = 0
                
            Else
                
                strText = objLI.Text
     
                If IsEmpty(objLI.SmallIcon) Then
                    
                    intOffset = 0
                
                Else
                    
                    Set objPic = objILS.Overlay(objLI.SmallIcon, objLI.SmallIcon)
                
                    Printer.PaintPicture objPic, lngX + px, lngY + (py / 2), 16 * px, 16 * py, , , , , vbSrcCopy
                    
                    intOffset = px * 16
                    
                End If
            
            End If
            
            '------------------------------------------------------------------
            'Make sure text fits
            '------------------------------------------------------------------
            
            lngWidth = (objCol.Width * dblXScale)
            
            lngX1 = lngX + lngWidth
            
            strTextTrun = strText
            
            Do Until Printer.TextWidth(strTextTrun) < lngWidth - (px * 5) - intOffset Or strText = ""
                
                strText = Left$(strText, Len(strText) - 1)
                
                strTextTrun = strText & "..."
            
            Loop
            
            Printer.Line (lngX, lngY)-(lngX1, lngY1), 1, B
            
            Printer.CurrentY = lngY + ((intRowHeight - Printer.TextHeight(strTextTrun)) / 2) + py
            
            Select Case objCol.Alignment
                   
                Case ListColumnAlignmentConstants.lvwColumnCenter
                    
                    Printer.CurrentX = lngX + intOffset + (((objCol.Width * dblXScale) - Printer.TextWidth(strTextTrun)) / 2)
                    
                Case ListColumnAlignmentConstants.lvwColumnLeft
                    
                    Printer.CurrentX = lngX + intOffset + (px * 5)
                
                Case ListColumnAlignmentConstants.lvwColumnRight
                    
                    Printer.CurrentX = lngX + ((objCol.Width * dblXScale) - intOffset - Printer.TextWidth(strTextTrun)) - (px * 5)
                    
            End Select
            
            '------------------------------------------------------------------
            'Print each colum
            '------------------------------------------------------------------
            
            Printer.Print strTextTrun
             
            lngX = lngX1
        
        Next
        
    Next
    
    '--------------------------------------------------------------------------
    'Print final page number
    '--------------------------------------------------------------------------
    
    lngPageNo = lngPageNo + 1
    
    Printer.CurrentX = (Printer.Width / 2) - (Printer.TextWidth("Page " & lngPageNo) / 2)
    Printer.CurrentY = lngEOP - intRowHeight
    Printer.Print "Page " & lngPageNo
    Printer.EndDoc
    
    gPrintListView = True
    
    Screen.MousePointer = vbDefault
    
    Set objCol = Nothing
    Set objILS = Nothing
    Set objLI = Nothing
    Set objPic = Nothing
    
    Exit Function
    
Error_Handler:
    
    Set objCol = Nothing
    Set objILS = Nothing
    Set objLI = Nothing
    Set objPic = Nothing
    
    Screen.MousePointer = vbDefault
    
    '--------------------------------------------------------------------------
    'Simple error message reporting
    '--------------------------------------------------------------------------
    
    MsgBox "gPrintListView() failed with the following error:-" & vbCrLf & vbCrLf & _
    "Error Number: " & Err.number & vbCrLf & "Description:" & Err.Description, vbExclamation
    
End Function
Sub SetCmdState(bState As Boolean)
    '不进行任何操作
End Sub
