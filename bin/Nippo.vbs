'===============================================================================
' モジュール名  :Nippou
' 説明          :指定された日付の日報ファイルを作成する
' @param        :yyyymmdd
' @param        :コピー元日報ファイルパス
' @retun        :0=正常終了 , 1=異常終了
' [改定履歴]
' 20201225 shinya.katori 新規作成
'===============================================================================

Option Explicit
On Error Resume Next
Err.Clear

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Dim objArgs
Dim objFS
Dim objFile1
Dim objFile2
Dim objTSR
Dim objTSW
Dim yyyymmdd
Dim copyFromFilepath
Dim copyToFilepath
Dim line

Set objArgs = WScript.Arguments
Set objFS = CreateObject("Scripting.FileSystemObject")

If objArgs.Count <> 2 Then
  Err.Number = 59999
  Err.Description = "引数不正"
  Call errorMsg
  Call terminate
Else
  yyyymmdd = objArgs(0)
  copyFromFilepath = objArgs(1)
End If

Set objFile1 = objFS.GetFile(copyFromFilepath)
If Err.Number <> 0 Then
  Call errorMsg
End If
Set objTSR = objFile1.OpenAsTextStream(ForReading, TristateUseDefault)
If Err.Number <> 0 Then
  Call errorMsg
End If

copyToFilepath = Left(objFile1.Path, Len(objFile1.Path) - Len(objFile1.Name)) & yyyymmdd & ".txt"
Set objFile2 = objFS.CreateTextFile(copyToFilepath, True)
Set objFile2 = objFS.GetFile(copyToFilepath)
Set objTSW = objFile2.OpenAsTextStream(ForWriting, TristateUseDefault)

Dim i:i=0
Dim kinmuJikanLine:kinmuJikanLine = 0
Do While objTSR.AtEndOfStream <> True
    i = i + 1
    line = objTSR.ReadLine
    
    If line = "▼勤務時間" Then
      kinmuJikanLine = i + 1
    End if
    If i = kinmuJikanLine Then
      objTSW.WriteLine "08:40～"
    End If
    
    Select Case i
        Case 1
          objTSW.WriteLine "■" & yyyymmdd & "日報"
        Case kinmuJikanLine
          '書き込みなし
        Case Else
          objTSW.WriteLine line
    End Select
Loop


If Err.Number <> 0 Then
  Call errorMsg
End If

Call terminate

'-------------------------------------------------------------------------------
' Error
'-------------------------------------------------------------------------------
Sub errorMsg
  WScript.Echo Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source
  Call terminate
End Sub

'-------------------------------------------------------------------------------
' Terminate
'-------------------------------------------------------------------------------
Sub terminate

  objTSR.Close
  objTSW.Close
  objFS.Close
  Set objArgs = Nothing
  Set objFile1 = Nothing
  Set objFile2 = Nothing
  Set objTSR = Nothing
  Set objTSW = Nothing
  Set objFS = Nothing
  WScript.Quit 0

End Sub

