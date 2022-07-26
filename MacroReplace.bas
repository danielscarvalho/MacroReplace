
Sub MacroReplace()
'
'  
' Macro1 Macro
' Replace multiple DOC files
'
  ' Excel Macro - VBA
'  
' Atalho do teclado: Ctrl+Shift+R
'
' At Microsoft Visual Basic for Applications, selecionar menu:
' Ferramentas, Preferências e ativar 
' "Microsoft Office 16.0 Object Library"
' 

Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim fileCount As Integer
Dim replaceCount As Integer
Dim replaceRows As Integer
Dim currentPath As String

Set oFSO = CreateObject("Scripting.FileSystemObject")

currentPath = Application.ActiveWorkbook.Path

Set oFolder = oFSO.GetFolder(currentPath)

replaceRows = Range("A2", Range("A2").End(xlDown)).Rows.Count

fileCount = 1

For Each oFile In oFolder.Files

    Let Filename = oFile.Name
    
    If Right$(Filename, 4) = "docx" Or Right$(Filename, 3) = "doc" Then
    
        Set oDoc = CreateObject("word.Application")
        Dim oSheet1 As Object
              
              
        oDoc.Visible = True
        MsgBox currentPath & "\" & Filename
             
        Set oSheet1 = oDoc.documents.Open(currentPath & "\" & Filename)
    
        Range("A2:B2").Select
        
        For replaceCount = 2 To replaceRows + 1
             Dim replaceFrom As String
             Dim replaceTo As String
        
             'replaceFrom = ActiveCell.Value
             replaceFrom = Cells(replaceCount, 1).Value
             replaceTo = Cells(replaceCount, 2).Value
             
             MsgBox (fileCount & " - " & oFile.Name & " - " & replaceFrom + " >> " + replaceTo)
             ' Insert your code here.
             ' Selects cell down 1 row from active cell.
             'ActiveCell.Offset(1, 0).Select
     
             
             'With oSheet1.Content.Find
             '   .Text = replaceFrom
             '   .Replacement.Text = replaceTo
             '   .Wrap = wdFindContinue
             '   .Execute Replace:=wdReplaceAll
             'End With
             
            'Selection.ParagraphFormat.Reset
            'Selection.Find.ClearFormatting
            'Selection.Find.Replacement.ClearFormatting
            'With oSheet1.Content.Find 'Selection.Find
            '    .Text = replaceFrom
            '    .Replacement.Text = replaceTo
            '    .Forward = True
            '    .Wrap = wdFindContinue
            '    .Format = False
            '    .MatchCase = False
            '    .MatchWholeWord = False
            '    .MatchWildcards = False
            '    .MatchSoundsLike = False
            '    .MatchAllWordForms = False
            'End With
            
            
            
            
            
            'oSheet1.Content.Find.Execute Replace:=wdReplaceAll
     
     
                With oSheet1.Content.Find
                   .ClearFormatting
                   .Replacement.ClearFormatting
                   .MatchWildcards = False
                   .Wrap = wdFindContinue
                   .Text = replaceFrom
                   .Replacement.Text = replaceTo
                   .Forward = True
                   .Wrap = wdFindStop
                   .Format = False
                   .MatchCase = False
                   .Execute Replace:=wdReplaceAll
                End With

     


        
     
     
             
             Range("A" & replaceCount & ":B" & replaceCount).Select
        Next
    
        fileCount = fileCount + 1
        
        'oDoc.Save
        'oDoc.Close
    
    End If

Next oFile

End Sub

    
    



