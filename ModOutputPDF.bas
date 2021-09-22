Attribute VB_Name = "ModOutputPDF"
Option Explicit

'OutputPDF・・・元場所：FukamiAddins3.ModFile

'------------------------------

'------------------------------


Sub OutputPDF(TargetSheet As Worksheet, Optional FolderPath$, Optional FileName$, _
              Optional MessageIrunaraTrue As Boolean = True)
'指定シートをPDF化する
'20210721

'TargetSheet・・・PDF化する対象のシート
'FolderPath ・・・出力先フォルダ 指定しない場合はブックと同じフォルダ
'FileName   ・・・出力PDFのファイル名 指定しない場合はシートの名前
    
    '引数チェック
    If FolderPath = "" Then
        FolderPath = TargetSheet.Parent.Path '指定がない場合は自ブックのフォルダパス
    End If
    
    If FileName = "" Then
        FileName = TargetSheet.Name '指定がない場合はシート名
    End If
    
    '出力先フォルダがない場合は作成する。
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If
    
    '出力するPDFのファイル名を作成する
    Dim OutputFileName$
    OutputFileName = FolderPath & "\" & FileName & ".pdf"
    
    'PDFで出力する
    TargetSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=OutputFileName
    
    '出力結果の確認メッセージ
    If MessageIrunaraTrue Then
        If MsgBox("「" & FileName & ".pdf" & "」" & vbLf & "を作成しました" & vbLf & _
            "出力先フォルダを起動しますか?", vbYesNo + vbQuestion) = vbYes Then
            Shell "C:\Windows\explorer.exe " & FolderPath, vbNormalFocus
        End If
    End If
    
End Sub


