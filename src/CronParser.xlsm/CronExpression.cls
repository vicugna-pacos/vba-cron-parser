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


Public IsError As Boolean   ' �ݒ�l�ɃG���[�����邩�ǂ���
Private fields() As String   ' �P�ʂ��Ƃ̐ݒ�l

Private CONSTRAINTS() As String ' �e�t�B�[���h�̗L���l
Private ALIASES() As String ' �e�t�B�[���h�̕ʖ�
Private DAY_OF_MONTH() As Integer  ' �����Ƃ̓���

' ����������
Private Sub Class_Initialize()
    IsError = False
    
    ' �Œ�l(���ǂ�)�̏����� ----------
    ' �L���l
    CONSTRAINTS = Array( _
          "0-59" _
        , "0-59" _
        , "0-23" _
        , "1-31" _
        , "1-12" _
        , "1-7")
    
    ' �ʖ�
    ALIASES = Array( _
         "" _
        , "" _
        , "" _
        , "" _
        , "jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12" _
        , "sun:1,mon:2,tue:3,wed:4,thu:5,fri:6,sat:7")

    ' �����Ƃ̓���
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

' �p�[�X
Public Sub Parse(source As String)
    Dim atom As String
    
    IsError = False
    fields = Split(source, " ")
    
    ' �z��̃T�C�Y��6�ɌŒ肷��
    If UBound(fields) <> 5 Then
        ReDim Preserve fields(6)
    End If

    ' �L���l�̃`�F�b�N
    For Each field In fields
    Next

End Sub

' �t�B�[���h�̃p�[�X
Private Function parseField(field As String, fieldIndex As Integer) As Boolean
    
    Dim atoms1() As String
    Dim atom1 As String
    Dim atoms2() As String
    Dim atom2 As String
    
    
    parseField = True
    atoms1 = Split(field, ",")
    
    For Each atom1 In atoms1
        atoms2 = Split(atom1, "-")
        
        ' �͈͎w��́u-�v���g����̂�1�񂾂�
        If UBound(atoms2) > 1 Then
            parseField = False
            Exit Function
        End If
        
        For Each atom2 In atoms2
            ' �ʖ�������Βu��������
            atom2 = ReplaceAlias(atom, fieldIndex)
        Next
        
        ' �͈͎w��̏ꍇ�́A���̕��v�f�𑝂₷
        If UBound(atoms2) = 1 Then
            
        End If
    Next
    
    
    
    
    
End Function

' �j���́usun�v���u1�v�ɒu��������ȂǁA�ʖ�������΂����u��������B
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