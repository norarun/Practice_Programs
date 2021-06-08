Attribute VB_Name = "Module1"
'================================
'
'file:  test
'Date:  2021/06/09
'用途:  他シートから条件にある情報をまとめる
'
'ver. 2.0
'================================

'定数
Const SheetName As String = "Sheet4"            '<<ここに、出力したいシートの名前を入れる>>
Const require As Date = "2021/3/1  0:00:00"     '<<ここは、条件、Dateになってるので条件に合わせて変数宣言してください。>>
Const item As String = "date"                   '<<ここに、項目の名前(例えば"id"や"日付")を入れる>>

'***********************
'メイン
'***********************
Sub SampleMain_onClick()
On Error GoTo ERROR_
    Dim A() As String   '配列
    Dim B() As String   '配列
    Dim C() As String   '配列
    Dim D() As String   '配列
    
    Call syokika_onClick                        '初期化、ユーザーメソッド
    
    Application.ScreenUpdating = False          '画面処理STOP
    
    'シート1からスタート、連想配列
    For Each mySheet In Worksheets
        If Not mySheet.Name = SheetName Then                 'もし、出力指定シートなら飛ばす
        
            Max_i = Worksheets(mySheet.Name).Range("A" & rows.count).End(xlUp).Row              '選択してるシートの最大 行 数取得
            Worksheets(mySheet.Name).Activate
            
            '挿入
            For i = 1 To (Max_i - 1)
            
                If require <= Cells(1, IndexUpdate(mySheet.Name, item)).Value Then    '条件 エラーの場合 ""の中身が一致してるか確認してみて
                
                    If isArrayEx(A) = -1 Then               '初期化の判定、ユーザーメソッド
                        
                        ReDim Preserve A(0)                 '初期化されていないので初期化
                        ReDim Preserve B(0)                 '初期化されていないので初期化
                        ReDim Preserve C(0)                 '初期化されていないので初期化
                        ReDim Preserve D(0)                 '初期化されていないので初期化
                        
                    Else
                        
                        ReDim Preserve A(UBound(A) + 1)     '配列を増やす
                        ReDim Preserve B(UBound(B) + 1)     '配列を増やす
                        ReDim Preserve C(UBound(C) + 1)     '配列を増やす
                        ReDim Preserve D(UBound(D) + 1)     '配列を増やす
                        
                    End If
                    
                    A(UBound(A)) = Cells(i + 1, 1).Value    '値を挿入している
                    B(UBound(B)) = Cells(i + 1, 2).Value    '値を挿入している
                    C(UBound(C)) = Cells(i + 1, 3).Value    '値を挿入している
                    D(UBound(D)) = Cells(i + 1, 4).Value    '値を挿入している
                    
                End If
            Next i
        End If
    Next
    
    '表示
    Application.ScreenUpdating = True       '画面処理開始
    Worksheets(SheetName).Activate
            
    For i = 1 To (UBound(A))
        index = i
    
        Cells(index, 1).Value = A(i) '値をCellに挿入している
        Cells(index, 2).Value = B(i) '値をCellに挿入している
        Cells(index, 3).Value = C(i) '値をCellに挿入している
        Cells(index, 4).Value = D(i) '値をCellに挿入している
        
    Next i
    
ERROR_:
    If Err.Number = 1004 Then
        MsgBox "38行目:IFの指定項目が存在していない可能性があります。再度確認してください。", vbOKOnly
    Else
    'any error's
    End If
End Sub
'--------------------------------------------------------------
'
'syokika_onClick
'
'@param     void
'@return    void
'
'最初の実行時に一度だけ実行される。セルをまっさらな状態に戻す。
'
'--------------------------------------------------------------
Sub syokika_onClick()
    Application.ScreenUpdating = False          '画面処理STOP
    
    '初期化
    Max_i = Worksheets(SheetName).Range("A" & rows.count).End(xlUp).Row              '選択してるシートの最大 行 数取得
    Worksheets(SheetName).Activate
    
    For i = 1 To Max_i
        
        Cells(i, 1).Value = ""      '初期化している
        Cells(i, 2).Value = ""
        Cells(i, 3).Value = ""
        Cells(i, 4).Value = ""
        
    Next i
    
    Application.ScreenUpdating = True          '画面処理STOP
End Sub
'--------------------------------------------------------------
'
'IndexUpdate
'
'@param     String      text    //条件
'@return    int         i       //日付の位置を取得
'
'引数のtextと同じ項目があれば、そこの位置の値を返す。そうでなければエラーを出力する。
'
'--------------------------------------------------------------
Function IndexUpdate(mySheet As String, text As String) As Integer
    Max_c = Worksheets(mySheet).Cells(1, Columns.count).End(xlToLeft).Column               '選択してるシートの最大 列 数取得
    
    For C = 1 To Max_c
        If text = Cells(1, C).Value Then
            IndexUpdate = C
            Exit Function
        End If
    Next C
    
    IndexUpdate = 0
End Function

'--------------------------------------------------------------
'WEBサイトから引用
'詳しくは：https://zukucode.com/2019/08/vba-array-loop.html
'
'機能：引数が配列か判定し、配列の場合は空かどうかも判定する
'戻り値：判定結果（1:配列 / 0:空の配列 / -1:配列ではない
'--------------------------------------------------------------
Public Function isArrayEx(varArray As Variant) As Long
On Error GoTo ERROR_
    If IsArray(varArray) Then
        isArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        isArrayEx = -1
    End If

    Exit Function

ERROR_:
    If Err.Number = 9 Then
        isArrayEx = -1
    Else
        '想定外エラー
    End If
End Function
