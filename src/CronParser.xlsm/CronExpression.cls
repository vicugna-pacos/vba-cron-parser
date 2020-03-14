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
Private fields() As String   ' 単位ごとの設定値

Private CONSTRAINTS() As String ' 各フィールドの有効値
Private ALIASES() As String ' 各フィールドの別名
Private DAY_OF_MONTH() As Integer  ' 月ごとの日数

' 初期化処理
Private Sub Class_Initialize()
    IsError = False
    
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
    Dim atom As String
    
    IsError = False
    fields = Split(source, " ")
    
    ' 配列のサイズを6に固定する
    If UBound(fields) <> 5 Then
        ReDim Preserve fields(6)
    End If

    ' 有効値のチェック
    For Each field In fields
    Next

End Sub

' フィールドのパース
Private Function parseField(field As String, fieldIndex As Integer) As Boolean
    
    Dim atoms1() As String
    Dim atom1 As String
    Dim atoms2() As String
    Dim atom2 As String
    
    
    parseField = True
    atoms1 = Split(field, ",")
    
    For Each atom1 In atoms1
        atoms2 = Split(atom1, "-")
        
        ' 範囲指定の「-」が使えるのは1回だけ
        If UBound(atoms2) > 1 Then
            parseField = False
            Exit Function
        End If
        
        For Each atom2 In atoms2
            ' 別名があれば置き換える
            atom2 = ReplaceAlias(atom, fieldIndex)
        Next
        
        ' 範囲指定の場合は、その分要素を増やす
        If UBound(atoms2) = 1 Then
            
        End If
    Next
    
    
    
    
    
End Function

' 曜日の「sun」を「1」に置き換えるなど、別名があればそれを置き換える。
Private Function ReplaceAlias(atom As String, fieldIndex As Integer) As String
    Dim aliasStr As String
    Dim arr1() As String
    Dim arr2() As String
    Dim str1 As String
    Dim alias As String
    Dim tobe As String
    
    ReplaceAlias = atom
    aliasStr = ALIASES(fieldIndex)
    
    If aliasStr = "" Then
        Exit Function
    End If
    
    arr1 = Split(aliasStr, ",")
    
    For Each str1 In arr1
        arr2 = Split(str1, ":")
        alias = arr2(0)
        tobe = arr2(1)
        
        If LCase(atom) = LCase(alias) Then
            ReplaceAlias = tobe
            Exit Function
        End If
    Next
End Function
