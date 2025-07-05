VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clBENCHMARK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private prCollect As Collection
Private prListeXLSB As Variant

Property Get listerXLSB() As Variant
Dim a
Dim xlsb As clXLSB

If IsEmpty(prListeXLSB) Then
    ReDim a(1 To 1)
    For i = 1 To prCollect.count
        Set xlsb = prCollect(i)
        a(i) = xlsb.path
        If i < prCollect.count Then
            ReDim Preserve a(1 To i + 1)
        End If
    Next i
    
    prListeXLSB = a
Else
    listerXLSB = prListeXLSB
End If

End Property
Private Sub Class_Initialize()
    Set prCollect = New Collection
    
End Sub


Property Get XLSBs() As Collection
    Set XLSBs = prCollect
End Property

Property Set XLSBs(nvCollect As Collection)
    Set prCollect = nvCollect
End Property



Public Sub Refresh()
Dim ofso As FileSystemObject
Dim oFld As Folder
Dim oFile As File
Dim Table As Collection
Dim path As String

Set Table = New Collection


path = ActiveWorkbook.path

Set ofso = New FileSystemObject


For Each oFld In ofso.GetFolder(path).SubFolders
    If oFld.Name = "SAUVEGARDE" Then
        test = True
        Exit For
    End If
Next

If Not test Then ofso.GetFolder(path).SubFolders.add ("SAUVEGARDE")

Set oFld = ofso.GetFolder(path)

Call getFile(oFld, Table)

ReDim prListeXLSB(1 To Table.count)
For i = 1 To Table.count
    prListeXLSB(i) = Table(i)
Next i

End Sub


Private Sub getFile(xfolder As Folder, Table As Collection)
Dim oFld As Folder, oFile As File

For Each oFile In xfolder.Files
    If oFile.Name Like "*.xlsb*" Then
        Table.add oFile.path
    End If
Next oFile

For Each oFld In xfolder.SubFolders
    getFile oFld, Table
Next

End Sub

Public Sub add(path As String)
Dim XSLB As clXLSB

Set xlsb = New clXLSB
With xlsb
   
    .path = path
    .Load
    Set .Parent = Me

End With
prCollect.add xlsb, xlsb.Id

End Sub

Public Sub Remove(PathOrId As String)
Dim xlsb As clXLSB
    For i = 1 To prCollect.count
        Set xlsb = prCollect(i)
        If xlsb.Id = PathOrId Or xlsb.path = PathOrId Then
            prCollect.Remove i
            Exit Sub
        End If
    Next i
    
End Sub

Public Sub chargerAll()
Dim path As String
Dim xlsb As clXLSB
Dim PresentX As Boolean

patiente.Show 0
Sheets("Erreur").Range("A1").CurrentRegion.ClearContents

On Error Resume Next

For i = LBound(prListeXLSB) To UBound(prListeXLSB)
    path = prListeXLSB(i)
    If Not Existe(path) Then
        add path
    End If
    
    With patiente
        .Caption = "chargement de " & i & " sur " & UBound(prListeXLSB)
        .ProgressBar.Value = Round(i / UBound(prListeXLSB) * 100)
        DoEvents
    End With
    
Next i

For i = 1 To prCollect.count
    Set xlsb = prCollect(i)
    For j = LBound(prListeXLSB) To UBound(prListeXLSB)
        If prListeXLSB(j) = xlsb.path Then
            PresentX = True
            Exit For
        End If
    Next j
    
    If Not PresentX Then Remove xlsb.path
Next i

Unload patiente


End Sub


Public Function Existe(PathOrId) As Boolean
For i = 1 To prCollect.count
    If prCollect(i).path = PathOrId Or prCollect(i).Id = PathOrId Then
        Existe = True
        Exit Function
    End If
Next i

End Function
Public Sub Save()
Dim xlsb As clXLSB
Dim Fille As clFllle

Dim a, b, c, v, x

Dim Detail As Variant

patiente.Show 0

   

'Sauvegarder PrID
a = Empty
b = Empty
v = Empty

For i = 1 To prCollect.count
    Set xlsb = prCollect(i)
    If i = 1 Then
        a = xlsb.tableID
       
    Else
        b = xlsb.tableID
        a = Ajoute(a, b)
    End If
Next i

With Sheets("ID").Range("A1")
    .CurrentRegion.ClearContents
    .Resize(UBound(a, 1) + 1, UBound(a, 2) + 1).Value = a
    Sheets("ID").ListObjects.add(xlSrcRange, .Resize(UBound(a, 1) + 1, UBound(a, 2) + 1), , xlYes).Name = "TableID"
    Sheets("ID").ListObjects("TableID").TableStyle = "TableStyleLight9"
    
End With

With patiente
    .Caption = "construction de la base de 1 sur 5"
    .ProgressBar.Value = Round(1 / 5 * 100)
    DoEvents
End With



'Sauvegarder PrSynth

a = Empty
b = Empty
v = Empty


For i = 1 To prCollect.count
    Set xlsb = prCollect(i)
    For j = 1 To xlsb.Filles.count
        Set Fille = xlsb.Filles.item(j)
        If i = 1 And j = 1 Then
            a = Fille.tableSynth
           
        Else
            b = Fille.tableSynth
            
            a = Ajoute(a, b)
        End If
    Next j
Next i

With Sheets("SYNTH").Range("A1")
    .CurrentRegion.ClearContents
    .Resize(UBound(a, 1) + 1, UBound(a, 2) + 1).Value = a
    Sheets("SYNTH").ListObjects.add(xlSrcRange, .Resize(UBound(a, 1) + 1, UBound(a, 2) + 1), , xlYes).Name = "TableSynth"
    Sheets("SYNTH").ListObjects("TableSynth").TableStyle = "TableStyleLight9"
End With

With patiente
    .Caption = "construction de la base de 2 sur 5"
    .ProgressBar.Value = Round(2 / 5 * 100)
    DoEvents
End With


'Sauvegarder PrPRP

a = Empty
b = Empty
v = Empty


For i = 1 To prCollect.count
    Set xlsb = prCollect(i)
    For j = 1 To xlsb.Filles.count
        Set Fille = xlsb.Filles.item(j)
        If i = 1 And j = 1 Then
            a = Fille.tablePRP
            
        Else
            b = Fille.tablePRP
            
            a = Ajoute(a, b)
        End If
    Next j
Next i


With Sheets("PRP").Range("A1")
    .CurrentRegion.ClearContents
    .Resize(UBound(a, 1) + 1, UBound(a, 2) + 1).Value = a
    Sheets("PRP").ListObjects.add(xlSrcRange, .Resize(UBound(a, 1) + 1, UBound(a, 2) + 1), , xlYes).Name = "TablePRP"
    Sheets("PRP").ListObjects("TablePRP").TableStyle = "TableStyleLight9"
End With


With patiente
    .Caption = "construction de la base de 3 sur 5"
    .ProgressBar.Value = Round(3 / 5 * 100)
    DoEvents
End With

'Sauvegarder Financement

a = Empty
b = Empty
v = Empty


For i = 1 To prCollect.count
    Set xlsb = prCollect(i)
    For j = 1 To xlsb.Filles.count
        Set Fille = xlsb.Filles.item(j)
        If i = 1 And j = 1 Then
            a = Fille.Financement
            
        Else
            b = Fille.Financement
            
            a = Ajoute(a, b)
        End If
    Next j
Next i


With Sheets("Financement").Range("A1")
    .CurrentRegion.ClearContents
    .Resize(UBound(a, 1) + 1, UBound(a, 2) + 1).Value = a
    Sheets("Financement").ListObjects.add(xlSrcRange, .Resize(UBound(a, 1) + 1, UBound(a, 2) + 1), , xlYes).Name = "TableFinancement"
    Sheets("Financement").ListObjects("TableFinancement").TableStyle = "TableStyleLight9"
    Sheets("Financement").Cells(1, UBound(a, 2) + 2).Value = "ORGANISME"
    Sheets("Financement").Cells(2, UBound(a, 2) + 2).Formula2R1C1 = _
        "=IF([@Ligne]=""Fonds propres"",""Fonds propres"",getOrganisme([@Colonne1]))"
End With


With patiente
    .Caption = "construction de la base de 3 sur 5"
    .ProgressBar.Value = Round(3 / 5 * 100)
    DoEvents
End With


'Sauvegarder PrPRPDetail

a = Empty
b = Empty
v = Empty


For i = 1 To prCollect.count
    Set xlsb = prCollect(i)
    For j = 1 To xlsb.Filles.count
        Set Fille = xlsb.Filles.item(j)
        If i = 1 And j = 1 Then
            a = Fille.PRPDetail
           
        Else
            b = Fille.PRPDetail
            
            a = Ajoute(a, b)
        End If
    Next j
Next i


With Sheets("PRPDetail").Range("A1")
    .CurrentRegion.ClearContents
    .Resize(UBound(a, 1), UBound(a, 2)).Value = a
    Sheets("PRPDetail").ListObjects.add(xlSrcRange, .Resize(UBound(a, 1), UBound(a, 2)), , xlYes).Name = "TablePRPDetail"
    
   
    Sheets("PRPDetail").ListObjects("TablePRPDetail").TableStyle = "TableStyleLight9"
End With

With patiente
    .Caption = "construction de la base de 4 sur 5"
    .ProgressBar.Value = Round(4 / 5 * 100)
    DoEvents
End With

'Sauvegarder Locatif

a = Empty
b = Empty
v = Empty


For i = 1 To prCollect.count
    On Error Resume Next
    Set xlsb = prCollect(i)
    For j = 1 To xlsb.Filles.count
        Set Fille = xlsb.Filles.item(j)
        If i = 1 And j = 1 Then
            a = Fille.Locatif
            
        Else
            b = Fille.Locatif
            
            a = empiler(a, b)
           
        End If
Retour:
    Next j
Next i


With Sheets("Locatif").Range("A1")
    .CurrentRegion.ClearContents
    .Resize(UBound(a, 1) + 1, UBound(a, 2) + 1).Value = a
    Sheets("Locatif").ListObjects.add(xlSrcRange, .Resize(UBound(a, 1) + 1, UBound(a, 2) + 1), , xlYes).Name = "TableLocatif"
    Sheets("Locatif").ListObjects("TableLocatif").TableStyle = "TableStyleLight9"
End With


With patiente
    .Caption = "construction de la base de 5 sur 5"
    .ProgressBar.Value = Round(5 / 5 * 100)
    DoEvents
End With


Unload patiente

Exit Sub
suivant:
xlsb.Erreur "Save Locatif - " & xlsb.path
GoTo Retour

End Sub



Private Function Transpose(Table As Variant) As Variant

Dim a, b

ReDim b(LBound(Table, 2) To UBound(Table, 2), LBound(Table, 1) To UBound(Table, 1))

For i = LBound(Table, 1) To UBound(Table, 1)
    For j = LBound(Table, 2) To UBound(Table, 2)
        b(j, i) = Table(i, j)
    Next j
Next i

Transpose = b

End Function


Private Function Ajoute(Table As Variant, x As Variant)
Dim c

ReDim c(LBound(Table, 1) To UBound(Table, 1) + UBound(x, 1), LBound(Table, 2) To UBound(Table, 2))

For i = LBound(c, 1) To UBound(c, 1)
    For j = LBound(c, 2) To UBound(c, 2)
        If i <= UBound(Table, 1) Then
            c(i, j) = Table(i, j)
        Else
            c(i, j) = x(i - UBound(Table, 1), j)
        End If
    Next j
Next i

Ajoute = c

End Function


Public Sub Load()
Dim xlsb As clXLSB
Dim Fille As clFllle

Dim a, b, c, v, x

Dim Detail As Variant

patiente.Show 0

   
On Error GoTo suivant

'Load PrID
a = ThisWorkbook.Worksheets("ID").Range("A1").CurrentRegion.Value

For i = LBound(a, 1) + 1 To UBound(a, 1)
    Set xlsb = New clXLSB
    With xlsb
        .Id = a(i, 1)
        .path = a(i, 2)
        
        ReDim b(0 To 1, LBound(a, 2) - 1 To UBound(a, 2) - 1)
        
        For j = LBound(b, 2) To UBound(b, 2)
            b(0, j) = a(1, j + 1)
            b(1, j) = a(i, j + 1)
        Next j
        
        .tableID = b
        
    End With
    patiente.ProgressBar.Value = Round(i / UBound(a, 1) * 100)
    
    prCollect.add xlsb, xlsb.Id
Next i

'Load Synthese

a = ThisWorkbook.Worksheets("SYNTH").Range("A1").CurrentRegion.Value


For i = LBound(a, 1) + 1 To UBound(a, 1)
    Set xlsb = prCollect(a(i, 1))
    Set Fille = xlsb.getFille(a(i, 2))
    

     ReDim b(0 To 1, LBound(a, 2) - 1 To UBound(a, 2) - 1)
        
    For j = LBound(b, 2) To UBound(b, 2)
        b(0, j) = a(1, j + 1)
        b(1, j) = a(i, j + 1)
    Next j
        
        
    Fille.tableSynth = b
    patiente.ProgressBar.Value = Round(i / UBound(a, 1) * 100)
         
Next i

'Load PRP
a = ThisWorkbook.Worksheets("PRP").Range("A1").CurrentRegion.Value


For i = LBound(a, 1) + 1 To UBound(a, 1)
    Set xlsb = prCollect(a(i, 1))
    Set Fille = xlsb.getFille(a(i, 2))
    

   ReDim b(0 To 1, LBound(a, 2) - 1 To UBound(a, 2) - 1)
        
        For j = LBound(b, 2) To UBound(b, 2)
            b(0, j) = a(1, j + 1)
            b(1, j) = a(i, j + 1)
        Next j
        
        
    Fille.tablePRP = b
    patiente.ProgressBar.Value = Round(i / UBound(a, 1) * 100)
         
Next i

'Load PRPDetail
a = ThisWorkbook.Worksheets("PRPDetail").Range("A1").CurrentRegion.Value


For i = LBound(a, 1) + 1 To UBound(a, 1)
    If Existe(a(i, 1)) Then
        Set xlsb = prCollect(a(i, 1))
        Set Fille = xlsb.getFille(a(i, 2))
        
    
       ReDim b(0 To 1, LBound(a, 2) - 1 To UBound(a, 2) - 1)
            
            For j = LBound(b, 2) To UBound(b, 2)
                b(0, j) = a(1, j + 1)
                b(1, j) = a(i, j + 1)
            Next j
            
            
        c = Fille.PRPDetail
        If IsEmpty(c) Then
            Fille.PRPDetail = b
        Else
            Fille.PRPDetail = Ajoute(c, b)
        End If
        
        patiente.ProgressBar.Value = Round(i / UBound(a, 1) * 100)
    End If
         
Next i

'Load Locatif
a = ThisWorkbook.Worksheets("Locatif").Range("A1").CurrentRegion.Value


For i = LBound(a, 1) + 1 To UBound(a, 1)
    Set xlsb = prCollect(a(i, 1))
    Set Fille = xlsb.getFille(a(i, 2))
    

   ReDim b(0 To 1, LBound(a, 2) - 1 To UBound(a, 2) - 1)
        
        For j = LBound(b, 2) To UBound(b, 2)
            b(0, j) = a(1, j + 1)
            b(1, j) = a(i, j + 1)
        Next j
        
        
    Fille.Locatif = b
    patiente.ProgressBar.Value = Round(i / UBound(a, 1) * 100)
         
Next i

'Load Financement
a = ThisWorkbook.Worksheets("Financement").Range("A1").CurrentRegion.Value


For i = LBound(a, 1) + 1 To UBound(a, 1)
    Set xlsb = prCollect(a(i, 1))
    Set Fille = xlsb.getFille(a(i, 2))
    

   ReDim b(0 To 1, LBound(a, 2) - 1 To UBound(a, 2) - 1)
        
        For j = LBound(b, 2) To UBound(b, 2)
            b(0, j) = a(1, j + 1)
            b(1, j) = a(i, j + 1)
        Next j
        
        
    Fille.Financement = b
    patiente.ProgressBar.Value = Round(i / UBound(a, 1) * 100)
         
Next i



Unload patiente

Exit Sub
suivant:
xlsb.Erreur "Save Locatif - " & xlsb.path


End Sub




