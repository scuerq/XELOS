VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clXLSB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private prID As String
Private prPath As String
Private prFilles As Collection
Private prParent As clBENCHMARK

Private prtableID As Variant


Private Sub Class_Initialize()
    Set prFilles = New Collection
End Sub

Property Get Filles() As Collection
    Set Filles = prFilles
End Property

Property Set Filles(nvFilles As Collection)
    Set prFilles = nvFilles
End Property
Property Get Id() As String
    Id = prID
End Property

Property Let Id(nvID As String)
    prID = nvID
End Property


Property Get Parent() As clBENCHMARK
    Set Parent = prParent
    
End Property

Property Set Parent(nvParent As clBENCHMARK)
    Set prParent = nvParent
End Property
Property Get path() As String
    path = prPath
End Property

Property Let path(nvPath As String)
    prPath = nvPath
End Property


Property Get tableID() As Variant
    tableID = prtableID
End Property

Property Let tableID(nvtableId As Variant)
    prtableID = nvtableId
End Property


Property Let AllDatas(nvAllDatas As Variant)
    prAllDatas = nvAllDatas
End Property

Public Sub Erreur(valeur As String)
Dim i As Integer

With Sheets("ERREUR")
    .Range("A1").Value = "Erreurs"
    i = .Range("A1").CurrentRegion.Rows.count
    .Range("A1").Offset(i, 0).Value = valeur
End With

End Sub



Public Sub Load()
Dim Classeur As Workbook
Dim a

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False ' Désactive l'exécution des macros
    
Set Classeur = GetObject(prPath)

ReDim a(0 To 1, 0 To 0)
a(0, 0) = "path"
a(1, 0) = prPath
prtableID = a

'On Error GoTo Suivant_2

MapID Classeur
MapSynth Classeur
MapPRP Classeur
MapFinancement Classeur

mapPRPDetail Classeur
MapLocatif Classeur

Retour_2:
Classeur.Close False

Application.EnableEvents = True
Application.DisplayAlerts = True
Application.ScreenUpdating = True

Exit Sub

Suivant_2:
If prParent.Existe(prPath) Then prParent.Remove (prPath)
GoTo Retour_2
End Sub

Private Sub MapSynth(xlClass As Workbook)

Dim Feuille As Worksheet
Dim Fille As clFllle
Set Feuille = xlClass.Worksheets("Fiche_Synthèse")
Dim Table

On Error GoTo suivant

For i = 1 To prFilles.count
    Set Fille = prFilles.item(i)
    Table = Empty
    
    Call addItemTable(Table, Feuille.Range("F18:F28"), Feuille.Range("F18:F28").Offset(0, Fille.Num))
    Call addItemTable(Table, Feuille.Range("D29"), Feuille.Range("D29").Offset(0, 2 + Fille.Num))
    Call addItemTable(Table, Feuille.Range("F31:F34"), Feuille.Range("F31:F34").Offset(0, Fille.Num))
    Call addItemTable(Table, Feuille.Range("D46"), Feuille.Range("D46").Offset(0, Fille.Num + 2))
    Call addItemTable(Table, Feuille.Range("D136:D138"), Feuille.Range("D136:D138").Offset(0, Fille.Num + 2))
    Call addItemTable(Table, Feuille.Range("D144:D146"), Feuille.Range("D144:D146").Offset(0, Fille.Num + 2))
    
    
    Table = add_id_to_table("Table", "synth", Table)
    Table = add_id_to_table("Id2", Fille.Id, Table)
    Table = add_id_to_table("Id", Fille.Parent.Id, Table)
    Fille.tableSynth = Table
Next i


Retour:

Exit Sub

suivant:
Erreur Me.path & " - " & "Synth"
GoTo Retour

End Sub

Private Sub MapLocatif(xlClass As Workbook)

Dim Feuille As Worksheet
Dim Fille As clFllle
Dim Plage As Range
Dim Table, table2, a
Set Feuille = xlClass.Worksheets("LoyersEtCharges")
Dim indice As Integer

On Error GoTo suivant

Table = Feuille.Range("C15:T55").Value

For k = 1 To prFilles.count
    Set Fille = prFilles(k)
    indice = 0
    a = Empty
    table2 = Empty
    For i = LBound(Table, 1) To UBound(Table, 1)
        If Fille.Id Like "*" & Table(i, 1) & "*" Then
            indice = indice + 1
            
            If IsEmpty(table2) Then ReDim table2(LBound(Table, 2) - 1 To UBound(Table, 2) - 1, 0 To 0)
                If Table(i, 1) > 0 Then
                    ReDim Preserve table2(LBound(table2, 1) To UBound(table2, 1), LBound(table2, 2) To UBound(table2, 2) + indice)
                        
                    For j = LBound(table2, 1) To UBound(table2, 1)
                        
                        table2(j, 0) = Table(1, j + 1)
                        table2(j, indice) = Table(i, j + 1)
                        
                    Next j
                    table2 = Transpose(table2)
                    table2 = add_id_to_table("Table", "locatif", table2)
                    table2 = add_id_to_table("Id2", Fille.Id, table2)
                    table2 = add_id_to_table("Id", Fille.Parent.Id, table2)
                End If
            
                If indice = 1 Then
                    a = table2
                Else
                    a = empiler(a, table2)
                End If
                    
                table2 = Empty
                
            End If
    Next i
    
    
    Fille.Locatif = a
Next k

Retour:
Exit Sub

suivant:
Erreur Me.path & " - " & "maplocatif"
GoTo Retour

End Sub

Private Sub mapPRPDetail(xlClass As Workbook)

Dim Feuille As Worksheet
Dim Fille As clFllle
Dim Plage As Range
Dim Table
Dim indice As Integer


On Error GoTo suivant

For i = 1 To prFilles.count
    Set Fille = prFilles.item(i)
    indice = Split(Fille.Id, " - ")(0)
    Set Feuille = xlClass.Worksheets("PRP CE SE_" & indice)
    
    Table = Empty
    
    Set Plage = Feuille.Range("D12:D15, D18:D22, D25:D29, D32:D37, D40:D45, D48:D51, D54:D55")
    Call addItemTableNum(Table, Plage, Plage.Offset(0, 6))
    
    
    Set Plage = Feuille.Range("D59:D84")
    Call addItemTableNum(Table, Plage, Plage.Offset(0, 6))
    
    
    Set Plage = Feuille.Range("D87:D103")
    Call addItemTableNum(Table, Plage, Plage.Offset(0, 6))
    
    
    Set Plage = Feuille.Range("D106:D112")
    Call addItemTableNum(Table, Plage, Plage.Offset(0, 6))
   
    
    Set Plage = Feuille.Range("D123:D127")
    Call addItemTableNum(Table, Plage, Plage.Offset(0, 6))
    
    

    Set Plage = Feuille.Range("D130:D132")
    Call addItemTableNum(Table, Plage, Plage.Offset(0, 6))
    
      
    
   
    Table = add_id_to_table("Table", "PRPDetail", Table)
    Table = add_id_to_table("Id2", Fille.Id, Table)
    Table = add_id_to_table("Id", Fille.Parent.Id, Table)
    
    TransposePRPDetail Table, Sheets("MAP").Range("Tableau6")
    
    Fille.PRPDetail = Table
   
Retour:
Next i

Exit Sub

suivant:
Erreur Me.path & " - " & "PRPDetail"
GoTo Retour


End Sub


Private Sub MapPRP(xlClass As Workbook)

Dim Feuille As Worksheet
Dim Fille As clFllle
Set Feuille = xlClass.Worksheets("PRP IKOS")
Dim Table

On Error GoTo suivant

For i = 1 To prFilles.count
    Set Fille = prFilles.item(i)
    Table = Empty
    Call addItemTable(Table, Feuille.Range("D52"), Feuille.Range("D52").Offset(0, 10 + Fille.Num))
    Call addItemTable(Table, Feuille.Range("D80"), Feuille.Range("D80").Offset(0, 10 + Fille.Num))
    Call addItemTable(Table, Feuille.Range("D99"), Feuille.Range("D99").Offset(0, 10 + Fille.Num))
    Call addItemTable(Table, Feuille.Range("D108"), Feuille.Range("D108").Offset(0, 10 + Fille.Num))
    Call addItemTable(Table, Feuille.Range("D119"), Feuille.Range("D119").Offset(0, 10 + Fille.Num))
    Call addItemTable(Table, Feuille.Range("D124"), Feuille.Range("D124").Offset(0, 10 + Fille.Num))
    Call addItemTable(Table, Feuille.Range("D136"), Feuille.Range("D136").Offset(0, 10 + Fille.Num))
    Call addItemTable(Table, Feuille.Range("D138"), Feuille.Range("D138").Offset(0, 10 + Fille.Num))
    
    
    Table = add_id_to_table("Table", "PRP", Table)
    Table = add_id_to_table("Id2", Fille.Id, Table)
    Table = add_id_to_table("Id", Fille.Parent.Id, Table)
    
    Fille.tablePRP = Table
    
Next i
Retour:
Exit Sub

suivant:
Erreur Me.path & " - " & "PRP"
GoTo Retour

End Sub

Private Sub MapFinancement(xlClass As Workbook)

Dim Feuille As Worksheet
Dim Fille As clFllle
Set Feuille = xlClass.Worksheets("Financement")
Dim a, b, c

On Error GoTo suivant

For i = 1 To prFilles.count
    Set Fille = prFilles.item(i)
    Table = Empty
    a = Empty
    b = Empty
    c = Empty
    
    Call addItemTable(a, Feuille.Range("E13:E33"), Feuille.Range("E13:E33").Offset(0, 4 + Fille.Num))
    a = Transpose(a)
    a = add_id_to_table("Ligne", "Subvention", a)
    
    Call addItemTable(b, Feuille.Range("F35:F45"), Feuille.Range("F35:F45").Offset(0, 3 + Fille.Num))
    b = Transpose(b)
    b = add_id_to_table("Ligne", "Fonds propres", b)
    
    Call addItemTable(c, Feuille.Range("F46:F71"), Feuille.Range("F46:F71").Offset(0, 3 + Fille.Num))
    c = Transpose(c)
    c = add_id_to_table("Ligne", "Prets", c)
    
    
    Table = empiler(a, b)
    Table = empiler(Table, c)
    
    
    Table = add_id_to_table("Table", "Financement", Table)
    Table = add_id_to_table("Id2", Fille.Id, Table)
    Table = add_id_to_table("Id", Fille.Parent.Id, Table)
    
    Fille.Financement = Table
    
Retour:
Next i

Exit Sub

suivant:
Erreur Me.path & " - " & "financement"
GoTo Retour

End Sub


Private Sub MapID(xlClass As Workbook)

Dim Feuille As Worksheet
Dim Fille As clFllle
Set Feuille = xlClass.Worksheets("Identif")

On Error GoTo suivant

prID = xlClass.path & xlClass.Name

Call addItemTable(prtableID, Feuille.Range("C7:C21"), Feuille.Range("D7:D21"))
Call addItemTable(prtableID, Feuille.Range("M19"), Feuille.Range("O19"))
Call addItemTable(prtableID, Feuille.Range("M22"), Feuille.Range("O22"))
Call addItemTable(prtableID, Feuille.Range("G85"), Feuille.Range("M85"))
Call addItemTable(prtableID, Feuille.Range("C86"), Feuille.Range("E86"))

prtableID = add_id_to_table("Id", prID, prtableID)

For i = 1 To 5
    If Feuille.Range("F57:J57").Cells(1, i).Value <> "" Then
        Set Fille = New clFllle
        Fille.Id = Feuille.Range("F57:J57").Cells(1, i).Value
        Fille.Num = i
        prFilles.add Fille
        Set Fille.Parent = Me
        
    End If
Next i

Retour:

Exit Sub

suivant:
Erreur Me.path & " - " & "mapId"
GoTo Retour


End Sub


Public Function getFille(id2) As clFllle
Dim Fille As clFllle

For i = 1 To prFilles.count
    If prFilles(i).Id = id2 Then
        Set getFille = prFilles(i)
        Exit Function
    End If
Next i

Set Fille = New clFllle
Fille.Id = id2
prFilles.add Fille

Set getFille = Fille


End Function



Private Function addItemTable(ByRef Table As Variant, plageChamps As Range, plageValeurs As Range)

    Dim i As Integer
    Dim nbChamps As Integer
    Dim nbLignes As Integer
    
    nbChamps = plageChamps.Cells.count

    ' Vérifier si Table est vide et la créer si nécessaire
    If IsEmpty(Table) Or IsNull(Table) Then
        ReDim Table(1, nbChamps - 1) ' 2 lignes (0 et 1) et autant de colonnes que de champs
        nbLignes = 0
    Else
        ' Déterminer le nombre actuel de lignes de données (Table est basé sur une 2D array)
        nbLignes = UBound(Table, 2) + 1
        ReDim Preserve Table(1, nbLignes + nbChamps - 1) ' Ajouter une nouvelle ligne
    End If

    ' Ajouter les valeurs des champs et des valeurs à la dernière ligne
    For i = nbLignes To nbLignes + nbChamps - 1
        j = i - nbLignes
        Table(0, i) = plageChamps.Cells(j + 1, 1).Value ' Toujours garder les en-têtes en ligne 0
        Table(1, i) = plageValeurs.Cells(j + 1, 1).Value ' Ajouter les valeurs en nouvelle ligne
    Next i

End Function

Private Function addItemTableNum(ByRef Table As Variant, plageChamps As Range, plageValeurs As Range, Optional plageTVA As Range)

    Dim i As Integer
    Dim nbChamps As Integer
    Dim nbLignes As Integer
    Dim Cell As Range
    Dim TVA As Double
    
    nbChamps = plageChamps.Cells.count

    ' Vérifier si Table est vide et la créer si nécessaire
    If IsEmpty(Table) Or IsNull(Table) Then
        ReDim Table(1, nbChamps - 1) ' 2 lignes (0 et 1) et autant de colonnes que de champs
        nbLignes = 0
    Else
        ' Déterminer le nombre actuel de lignes de données (Table est basé sur une 2D array)
        nbLignes = UBound(Table, 2) + 1
        ReDim Preserve Table(1, nbLignes + nbChamps - 1) ' Ajouter une nouvelle ligne
    End If

    ' Ajouter les valeurs des champs et des valeurs à la dernière ligne
    i = nbLignes - 1
    For Each Cell In plageChamps.Cells
        i = i + 1
       
        Table(0, i) = Cell.Row ' Toujours garder les en-têtes en ligne 0
        If Not (plageTVA Is Nothing) Then
            If IsNumeric(Cell.Offset(0, plageTVA.Column - plageChamps.Column).Value) Then
                TVA = 1 + Cell.Offset(0, plageTVA.Column - plageChamps.Column).Value
            Else
                TVA = 1
            End If
        Else: TVA = 1
        End If
        
        Table(1, i) = Cell.Offset(0, plageValeurs.Column - plageChamps.Column).Value * TVA ' Ajouter les valeurs en nouvelle ligne
    Next Cell

End Function

Private Function renseigneItemTable(ByRef Table As Variant, Attribut As String, valeur As String)

    Dim i As Integer
    Dim nbLignes As Integer
    
    
    ' Déterminer le nombre actuel de lignes de données (Table est basé sur une 2D array)
     nbLignes = UBound(Table, 2) + 1
   
   

    ' Ajouter les valeurs des champs et des valeurs à la dernière ligne
    For i = nbLignes To nbLignes + nbChamps - 1
        j = i - nbLignes
        If Table(0, i) = Attribut Then ' Toujours garder les en-têtes en ligne 0
            Table(1, i) = valeur ' Ajouter les valeurs en nouvelle ligne
        End If
    Next i

End Function


Private Function nommerTableDetail(ByRef Table As Variant, nom As String, plageChamps As Range)
Dim i As Integer
    Dim nbChamps As Integer
    Dim nbLignes As Integer
    
    nbChamps = plageChamps.Cells.count

    ' Vérifier si Table est vide et la créer si nécessaire
    If IsEmpty(Table) Or IsNull(Table) Then
        ReDim Table(1, nbChamps - 1) ' 2 lignes (0 et 1) et autant de colonnes que de champs
        nbLignes = 0
    Else
        ' Déterminer le nombre actuel de lignes de données (Table est basé sur une 2D array)
        nbLignes = UBound(Table, 2) + 1
        ReDim Preserve Table(1, nbLignes + nbChamps - 1) ' Ajouter une nouvelle ligne
    End If

    ' Ajouter les valeurs des champs et des valeurs à la dernière ligne
    For i = nbLignes To nbLignes + nbChamps - 1
        j = i - nbLignes
        Table(1, i) = plageChamps.Cells(j + 1, 1).Value ' Toujours garder les en-têtes en ligne 0
        Table(0, i) = nom ' Ajouter les valeurs en nouvelle ligne
    Next i
End Function



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

Public Function getValeurFromID(Champs As String) As String
    For i = LBound(prtableID, 2) To UBound(prtableID, 2)
        If prtableID(LBound(prtableID, 1), i) = Champs Then
            getValeurFromID = prtableID(UBound(prtableID, 1), i)
            Exit Function
        End If
    Next i
    
End Function

Private Sub TransposePRPDetail(tPrp, tMap As Range)

Dim a, b, c, d
Dim collect As Collection

Set collect = New Collection

On Error GoTo suivant
a = tPrp
b = tMap.Value

ReDim c(1 To 7)
c(1) = a(0, 0)
c(2) = a(0, 1)
c(3) = a(0, 2)
c(4) = "POSTE_1"
c(5) = "POSTE_2"
c(6) = "POSTE_3"
c(7) = "MONTANT¨POSTE"

collect.add c

For i = LBound(a, 2) + 3 To UBound(a, 2)
    ReDim c(1 To 7)
   
    c(1) = a(1, 0)
    c(2) = a(1, 1)
    c(3) = a(1, 2)
    For k = LBound(b, 1) To UBound(b, 1)
        If b(k, 1) = a(0, i) Then
            c(4) = b(k, 4)
            c(5) = b(k, 3)
            c(6) = b(k, 2)
            Exit For
        End If
    Next k
    
    c(7) = a(1, i)
    collect.add c
    
    
Next i

ReDim d(1 To collect.count, 1 To 7)

For i = 1 To collect.count
    For j = 1 To 7
        d(i, j) = collect(i)(j)
    Next j
Next i

tPrp = d
 
Exit Sub

suivant:
MsgBox ("Stop")

        
End Sub
