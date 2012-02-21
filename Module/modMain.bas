Attribute VB_Name = "modMain"
Public GetApp As String '本地路径
Public curUserID As Integer     '用户ID
Public curUserName As String    '当前用户名
Public curUserLevel As Long    '当前用户类型
Public DataOperateState As String      '数据录入/编辑状态
Public mainID As Long   '选择的main数据表记录的ID号,用于编辑修改对应的记录
Public subID As Long
Public borrowID As Long   '选择的借支表的记录ID
Public incomeID As Long
Public curDOCType As Integer    '生成文档类型：1-结算单，2-项目确认单，3-项目借支单
Public dblBalace As Double      '借支余额
Public strContractType() As String    '合同类型
Public strMode() As String
Public curList1Index As Integer     '合同列表位置
Public curList2Index As Byte
Public curList3Index As Byte
Public curList4Index As Byte
Public curList5Index As Byte
Public bytAfterDec As Byte       '工作量文本框小数位数
Public color(2) As Long         '0-列表背景色，1-列表文本色，2-已结算文本色
Public blBorrow As Boolean     'true-借用, false-领用
Public blStocksShow As Boolean    '显示库存


'程序入口
Public Sub Main()
'On Error Resume Next
    
    If App.PrevInstance Then
        End
        Exit Sub
    End If
    '获得本地路径
    GetApp = App.Path: If Right$(GetApp, 1) <> "\" Then GetApp = GetApp & "\"
    
    GetItemInfo
    
    curUserID = 7
    
    bIsBorrow = True
    blblStocksShow = True
    
    'frmMain.Show
    frmLogin.Show
End Sub

Public Sub SetSB(ByVal i&, ByVal strText$)
    frmMain.LBSB(i).caption = strText
End Sub

Public Function GetINI(ByVal s1 As String, s2 As String)
On Error Resume Next
    GetINI = GetSetting("MySuperMarket", s1, s2)
End Function

Public Function GetLongINI(ByVal s1 As String, s2 As String, Optional Def As Long = 0) As Long  '获取INI中整数值
On Error GoTo aaaa
    Dim str As String
    str = GetINI(s1, s2)
    If str = "" Then
        GetLongINI = Def
    Else
        GetLongINI = CLng(str)
    End If
    Exit Function
aaaa:
    GetLongINI = Def
End Function

Public Sub SaveINI(ByVal s1 As String, s2 As String, s3 As String)
On Error Resume Next
    SaveSetting "MySuperMarket", s1, s2, s3
End Sub

Function FieldTypeIsChar(n As Long) As Boolean    '判断字段是否数字类型，用于插入记录时是否加引号
    Dim IsChar As Boolean
    
Select Case n
'case常量 值 说明
'Case 0x2000
' p = AdArray '（不适用于 ADOX。） 0x2000 一个标志值，通常与另一个数据类型常量组合，指示该数据类型的数组。
Case 20, 128, 14, 5, 3, 205, 131, 4, 2, 16, 21, 19, 18, 17, 204
    IsChar = False
Case 8, 136, 129, 6, 7, 133, 134, 135, 205, 203, 200, 202
    IsChar = True
End Select

FieldTypeIsChar = IsChar

End Function

Function GetID(id As String) As String
    GetID = Left(id, Len(id) - 1)
End Function

Function coverToChinese(money As String) As String
    Dim x As String, Y As String
    Const zimu = ".sbqwsbqysbqwsbq" '定义位置代码
    Const letter = "0123456789sbqwy.zjf" '定义汉字缩写
    Const upcase = "零壹贰叁肆伍陆柒捌玖拾佰仟万亿元整角分" '定义大写汉字
    Dim temp As String
    temp = money
    If InStr(temp, ".") > 0 Then temp = Left(temp, InStr(temp, ".") - 1)

    If Len(temp) > 16 Then MsgBox "数目太大，无法换算！请输入一亿亿以下的数字", vbCritical, "错误提示": Exit Function  '只能转换一亿亿元以下数目的货币！

    x = Format(money, "0.00") '格式化货币
    Y = ""
    For i = 1 To Len(x) - 3
        Y = Y & Mid(x, i, 1) & Mid(zimu, Len(x) - 2 - i, 1)
    Next
    If Right(x, 3) = ".00" Then
        Y = Y & "z"          '***元整
    Else
        Y = Y & Left(Right(x, 2), 1) & "j" & Right(x, 1) & "f"     '*元*角*分
    End If
    
    Y = Replace(Y, "0q", "0") '避免零千(如：40200肆萬零千零贰佰)
    Y = Replace(Y, "0b", "0") '避免零百(如：41000肆萬壹千零佰)
    Y = Replace(Y, "0s", "0") '避免零十(如：204贰佰零拾零肆)

    Do While Y <> Replace(Y, "00", "0")
        Y = Replace(Y, "00", "0") '避免双零(如：1004壹仟零零肆)
    Loop
    
    Y = Replace(Y, "0y", "y") '避免零億(如：210億     贰佰壹十零億)
    Y = Replace(Y, "0w", "w") '避免零萬(如：210萬     贰佰壹十零萬)
    Y = IIf(Len(x) = 5 And Left(Y, 1) = "1", Right(Y, Len(Y) - 1), Y) '避免壹十(如：14壹拾肆；10壹拾)
    Y = IIf(Len(x) = 4, Replace(Y, "0.", ""), Replace(Y, "0.", ".")) '避免零元(如：20.00贰拾零圆；0.12零圆壹角贰分)

    For i = 1 To 19
        Y = Replace(Y, Mid(letter, i, 1), Mid(upcase, i, 1)) '大写汉字
    Next
    coverToChinese = Y
    
End Function
  Public Function DirExists(ByVal strDirName As String) As Integer
          Const strWILDCARD$ = "*.*"
        
          Dim strDummy     As String
    
          On Error Resume Next
          If Trim(strDirName) = "" Then
                DirExists = 0
                Exit Function
          End If
          strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
          DirExists = Not (strDummy = vbNullString)
    
          Err = 0
  End Function
  Public Sub GetItemInfo()
  
    Exit Sub
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set rs = New ADODB.Recordset
    DBConnect
      
      '合同类型
    strSQL = "select * from ItemInfo where ItemType=1 order by ItemID"
    rs.Open strSQL, Conn, 1, 1
    ReDim strContractType(IIf(rs.RecordCount > 0, rs.RecordCount - 1, 0), 1)
    For i = 1 To rs.RecordCount
        strContractType(i - 1, 0) = rs("ItemName")
        strContractType(i - 1, 1) = rs("ItemID")
        rs.MoveNext
    Next
    
    rs.Close
    strSQL = "select * from ItemInfo where ItemType=2 order by ItemID"
    rs.Open strSQL, Conn, 1, 1
    ReDim strMode(IIf(rs.RecordCount > 0, rs.RecordCount - 1, 0), 1)
    For i = 1 To rs.RecordCount
        strMode(i - 1, 0) = rs("ItemName")
        strMode(i - 1, 1) = rs("ItemID")
        rs.MoveNext
    Next
    
    '小数位数
    Set rs = New ADODB.Recordset
    strSQL = "select ItemValue from ItemInfo where ItemType=3"
    rs.Open strSQL, Conn, 1, 1
    
    bytAfterDec = 3          '三位小数
    
    If Not rs.EOF Then
        If Not IsNull(rs("ItemValue")) Then bytAfterDec = rs("ItemValue")
    Else
        Conn.Execute "insert into ItemInfo(ItemType,ItemValue) values(3,3)"
    End If
    rs.Close
        
    '颜色
    strSQL = "select * from ItemInfo where ItemType=4 order by ItemID"
    rs.Open strSQL, Conn, 1, 1
    If rs.RecordCount <> 3 Then
        color(0) = "&Hfafafa"     '缺省颜色
        color(1) = "&H000000"
        color(2) = "&H008000"
    Else
        color(0) = rs("ItemValue")
        rs.MoveNext
        color(1) = rs("ItemValue")
        rs.MoveNext
        color(2) = rs("ItemValue")
    End If
    
    rs.Close
    Conn.Close
    Set rs = Nothing
    Set Conn = Nothing

  End Sub
