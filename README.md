#### 隐藏重复行
```vba
Sub HideDuplication()    
  Dim rowsCount As Integer    
  rowsCount = Me.UsedRange.Rows.Count        
  Range("D3:D" & rowsCount).AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _    
  Range("D3:D" & rowsCount), Unique:=True    ' Unique:=False 显示重复行
End Sub
```

####  快速删除重复记录
```vba
Sub RemoveDuplicateRecord() ' 快速删除重复记录
  Dim Row0 As Long, Row1 As Long
  With ActiveSheet
    Row0 = .Cells(1048576, 1).End(xlUp).Row
    .UsedRange.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    Row1 = .Cells(1048576, 1).End(xlUp).Row
  End With
  MsgBox "共删除" & Row0 - Row1 & "(" & Row0 & "-" & Row1 & ")" & "条记录", vbInformation, "提示"
End Sub
```

#### 读写文件
```vba
Public Function exportHtmlToFile(ByVal dtsId As String, ByVal html As String)
  If Not CreateObject("Scripting.FileSystemObject").FolderExists("D:\DTS_TXT") Then
    CreateObject("Scripting.FileSystemObject").CreateFolder ("D:\DTS_TXT")    
  End If        
  Dim filePath As String    
  filePath = "D:\DTS_TXT\TEST.TXT"    
  Open filePath For Output As #1        
    Print #1, html    
  Close #1    
End Function

Private Function getHtmlContentFromFile(ByVal filename As String) As String    
  Dim result As String    
  Dim line As String        
  Open filename For Input As #1        
  While Not EOF(1)            
    Line Input #1, line            
    result = result & line & vbCrLf        
    Wend   
  Close 1        
  getHtmlContentFromFile = result
End Function
```

#### 删除重复行
```vba
Sub RemoveDuplicateRecord()    
  Me.UsedRange.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
End Sub
```

