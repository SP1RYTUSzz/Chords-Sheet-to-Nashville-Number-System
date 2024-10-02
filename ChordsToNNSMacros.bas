Attribute VB_Name = "ChordtoNNSMacros"
Sub ChordToNashville()
Attribute ChordToNashville.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro
'
'
    Dim chordSheet As Variant
    Dim sharp As Boolean
    sharp = False
    Dim chordOffset As Integer
    chordOffset = 0
    chordFlat = Array("A", "Bb", "B", "C", "Db", "D", "Eb", "E", "F", "Gb", "G", "Ab")
    chordSharp = Array("A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#")
    
    Dim chordMap As Object
    Set chordMap = CreateObject("Scripting.Dictionary")
    Dim bracketNotations As Object
    Set bracketNotations = CreateObject("Scripting.Dictionary")
    
    chordPrefices = Array(" ", "^13")
    chordNumbers = Array(2, 4, 5, 6, 7, 9, 11, 13)

    '
    ' [Key detection here]
    '
    '
    Dim key As String
    key = FindKey("why")
    If key <> "Not found" Then
        MsgBox "Key is " & key & "."
    Else
        MsgBox "Key was not found"
    End If
    
    ' Find if sharp or flat key signatures by searching for sharps in the document
    sharp = FindTextInDocument("#")
    
    ' Find index of the key in the chordFlat or chordSharp
    For i = 0 To 11
        If chordFlat(i) = key Then
            chordOffset = i
            Exit For
        ElseIf chordSharp(i) = key Then
            chordOffset = i
            Exit For
        End If
    Next i

    semitoneFromTonic = chordOffset
    If sharp = True Then
        For j = 1 To 7
            If Not chordMap.Exists(chordSharp(semitoneFromTonic)) Then chordMap.Add chordSharp(semitoneFromTonic), j
            semitoneFromTonic = semitoneFromTonic + 2
            If j = 3 Then
                semitoneFromTonic = semitoneFromTonic - 1
            End If
            If semitoneFromTonic >= 12 Then
                semitoneFromTonic = semitoneFromTonic - 12
            End If
        Next j
    Else
        For j = 1 To 7
            If Not chordMap.Exists(chordFlat(semitoneFromTonic)) Then chordMap.Add chordFlat(semitoneFromTonic), j
            semitoneFromTonic = semitoneFromTonic + 2
            If j = 3 Then
                semitoneFromTonic = semitoneFromTonic - 1
            End If
            If semitoneFromTonic >= 12 Then
                semitoneFromTonic = semitoneFromTonic - 12
            End If
        Next j
    End If
    
    bracketNotations.Add "horus", "Chorus"
    bracketNotations.Add "ridge", "Bridge"
    
    For Each chord In chordMap.Keys
         ' Replace chord with number only into number in parentheses
         For Each chordNumber In chordNumbers:
             Selection.Find.ClearFormatting
             With Selection.Find.Font
                 .Bold = True
             End With
             Selection.Find.Replacement.ClearFormatting
             With Selection.Find
                 .Text = chord & chordNumber
                 .Replacement.Text = chordMap(chord) & "(" & chordNumber & ")"
                 .Forward = True
                 .Wrap = wdFindContinue
                 .Format = True
                 .MatchCase = False
                 .MatchWholeWord = False
                 .MatchAllWordForms = False
                 .MatchSoundsLike = False
                 .MatchWildcards = True
             End With
             Selection.Find.Execute Replace:=wdReplaceAll
         Next chordNumber

         ' Replace plain chords into Nashville Notation
         Selection.Find.ClearFormatting
         With Selection.Find.Font
             .Bold = True
         End With
         Selection.Find.Replacement.ClearFormatting
         With Selection.Find
             .Text = chord
             .Replacement.Text = chordMap(chord)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = True
             .MatchCase = False
             .MatchWholeWord = False
             .MatchAllWordForms = False
             .MatchSoundsLike = False
             .MatchWildcards = True
         End With
         Selection.Find.Execute Replace:=wdReplaceAll
    Next chord
    
    'Haha change the Bridge and Chorus back from it weird format
    For Each bracketNotation In bracketNotations
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Font.Bold = True
        With Selection.Find
            .Text = "\[" & "[1-7]" & bracketNotation
            .Replacement.Text = "[" & bracketNotations(bracketNotation)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next bracketNotation
End Sub
Function FindTextInDocument(ByVal searchText As String) As Boolean
    Dim findResult As Boolean
    Dim rng As Range
    
    ' Set the range to search the entire document
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .Text = searchText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        findResult = .Execute
    End With
    
    ' Return true if found, false otherwise
    FindTextInDocument = findResult
End Function
Function FindKey(ByVal keyText As String) As String
    Dim rng As Range
    Dim foundWord As String
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .Text = "Key: "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .Execute
    End With
    
    ' If the key text is found
    If rng.Find.Found Then
        ' Move the range to the word that comes after the keyText
        rng.MoveStart wdWord, 2
        rng.MoveEnd wdWord, 2
        foundWord = Trim(rng.Text)
    Else
        foundWord = "Not found"
    End If
    foundWord = Replace(foundWord, " ", "")
    foundWord = Replace(foundWord, Chr(13), "")
    FindKey = foundWord
End Function
