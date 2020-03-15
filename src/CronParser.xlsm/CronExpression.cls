VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CronExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public IsError As Boolean   ' 設定値にエラーがあるかどうか
Public ErrorMessage As String

Private fields() As String   ' 単位ごとの設定値

Private CONSTRAINTS() As String ' 各フィールドの有効値
Private ALIASES() As String ' 各フィールドの別名
Private DAY_OF_MONTH() As Integer  ' 月ごとの日数

' 初期化処理
Private Sub Class_Initialize()
    IsError = False
    ErrorMessage = ""
    
    ' 固定値(もどき)の初期化 ----------
    ' 有効値
    CONSTRAINTS = Array( _
          "0-59" _
        , "0-59" _
        , "0-23" _
        , "1-31" _
        , "1-12" _
        , "1-7")
    
    ' 別名
    ALIASES = Array( _
         "" _
        , "" _
        , "" _
        , "" _
        , "jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12" _
        , "sun:1,mon:2,tue:3,wed:4,thu:5,fri:6,sat:7")

    ' 月ごとの日数
    DAY_OF_MONTH = Array( _
        31, _
        29, _
        31, _
        30, _
        31, _
        30, _
        31, _
        31, _
        30, _
        31, _
        30, _
        31)

End Sub

' パース
Public Sub Parse(source As String)
    Dim field As String
    Dim i As Integer
    
    IsError = False
    ErrorMessage = ""
    
    fields = Split(source, " ")
    
    ' 配列のサイズを6に固定する
    If UBound(fields) <> 5 Then
        ReDim Preserve fields(6)
    End If
    
    ' チェック：日と曜日のいずれかがワイルドカードになっているか
    If Not CheckDayAndWeek(fields(3), fields(5)) Then
        IsError = True
        ErrorMessage = "日と曜日のいずれかにワイルドカードを指定してください"
        Exit Sub
    End If
    
    ' フィールドごとのパース
    For i = LBound(fields) To UBound(fields)
        ParseField fields(i), i
    Next

End Sub

' フィールドごとのパース
Private Sub ParseField(field As String, fieldIndex As Integer)
    
    Dim result As String
    
    
    result = field
    
    ' ワイルドカードの置き換え
    If result = "*" Or result = "?" Then
        result = CONSTRAINTS(fieldIndex)
    End If
    
    ' 別名を置き換える
    result = ReplaceAlias(result, fieldIndex)
    
    ' リスト(,)のパース
    fields(fieldIndex) = ParseList(result)
    
    
End Sub

' リスト(,)のパース
Private Function ParseList(source As String) As String
    Dim atoms() As String
    Dim i As Integer
    Dim result As String
    
    
    result = ""
    atoms = Split(source, ",")
    
    For i = LBound(atoms) To UBound(atoms)
        If result <> "" Then
            result = result & ","
        End If
        
        result = result & ParseRepeat(atoms(i))
    Next
    
End Function

' 繰り返し(/)のパース
Private Function ParseRepeat(source As String) As String
    Dim atoms() As String
    
    atoms = Split(source, "/")
    
    If UBound(atoms) = 0 Then
        ' 繰り返しなし
        ParseRepeat = ParseRange(source, 1)
    Else
        ' 繰り返しあり
        ParseRepeat = ParseRange(atoms(0), CInt(atoms(1)))
    End If
End Function

' 範囲(-)のパース
'
' 引数：
'   source - パース前の文字列
'   interval - 繰り返し間隔
' 戻り値：パース後の文字列。カンマ区切り。
Private Function ParseRange(source As String, interval As Integer) As String
    
    Dim atoms() As String
    Dim startInt As Integer
    Dim endInt As Integer
    Dim i As Integer
    Dim result As String
    
    
    result = ""
    atoms = Split(source, "-")
    
    If UBound(atoms) = 0 Then
        ' 範囲なし
        ParseRange = source
        Exit Function
        
    ElseIf UBound(atoms) > 1 Then
        ' 不正な値
        ParseRange = source
        IsError = True
        Exit Function
    End If
    
    ' 範囲を解体してリストを作る
    ' ※とりあえず特殊文字のことは考慮なし
    startInt = atoms(0)
    endInt = atoms(1)
    
    For i = startInt To endInt Step interval
        If i > startInt Then
            result = result & ","
        End If
        result = result & i
    Next
    
    ParseRange = result
    
End Function

' 曜日の「sun」を「1」に置き換えるなど、別名があればそれを置き換える。
Private Function ReplaceAlias(field As String, fieldIndex As Integer) As String
    Dim aliasStr As String
    Dim arr1() As String
    Dim arr2() As String
    Dim i As Integer
    Dim alias As String
    Dim tobe As String
    Dim result As String
    
    
    ReplaceAlias = field
    aliasStr = ALIASES(fieldIndex)
    
    If aliasStr = "" Then
        Exit Function
    End If
    
    result = field
    
    ' 設定値を文字列で持っているので分割する
    arr1 = Split(aliasStr, ",")
    
    For i = LBound(arr1) To UBound(arr1)
        arr2 = Split(arr1(i), ":")
        alias = arr2(0)
        tobe = arr2(1)
        
        ' 大文字小文字を無視するために「vbTextCompare」を使って置換
        result = Replace(result, alias, tobe, 1, -1, vbTextCompare)
        
    Next
    
    ReplaceAlias = result
End Function

' 日と曜日のいずれかのフィールドにワイルドカードが指定されているかチェックする
Private Function CheckDayAndWeek(dayStr As String, weekStr As String) As Boolean
    CheckDayAndWeek = True
    
    If dayStr <> "*" And dayStr <> "?" And weekStr <> "*" And weekStr <> "?" Then
        CheckDayAndWeek = False
    End If
End Function
