Attribute VB_Name = "bionic_reading"

Option Explicit

Sub bionic_reading_1()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Dim Coll As New Collection
    Dim myRange As Range
    Dim i As Long
    Dim myWord, mid_myWord As String
    Dim fnd As Variant
    
    Set myRange = Selection.Range
    'Set myRange = ActiveDocument.Range
    
    For i = 1 To myRange.Words.Count
        On Error Resume Next
        myWord = myRange.Words(i)
        mid_myWord = Left(myWord, Int(Len(myWord) / 2))
        Coll.Add mid_myWord, mid_myWord
        On Error GoTo 0
    Next
    
    'https://docs.microsoft.com/es-es/office/vba/api/word.find
    With myRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchPrefix = True
        For Each fnd In Coll
            On Error Resume Next
            .Text = fnd
            .Replacement.Text = "^&"
            .Execute Replace:=wdReplaceAll
            On Error GoTo 0
        Next fnd
    End With
                    
    Application.ScreenUpdating = True
End Sub

Sub bionic_reading_2()
    On Error Resume Next
    'Application.ScreenUpdating = False
        
    Dim myRange As Range
    Dim i As Long
    Dim myWord As String
    
    Set myRange = Selection.Range
    'Set myRange = ActiveDocument.Range

    For i = 1 To myRange.Words.Count
        On Error Resume Next
        With myRange.Words(i).Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Replacement.Font.Bold = True
            .Format = True
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .MatchPrefix = True
            myWord = myRange.Words(i)
            .Text = Left(myWord, Int(Len(myWord) / 2))
            .Replacement.Text = "^&"
            .Execute Replace:=wdReplaceAll
        End With
        On Error GoTo 0
    Next i

    'Application.ScreenUpdating = True
End Sub
