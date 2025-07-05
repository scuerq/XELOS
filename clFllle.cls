VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clFllle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private prID As String
Private prtableSynth As Variant
Private prTablePRP As Variant
Private prPRPDetail As Variant
Private prFinancement As Variant
Private prPRPDetailTitre As Variant
Private prLocatif As Variant
Private prAllDatas As Variant
Private prAllPRP As Variant
Private prParent As clXLSB
Public Num As Integer


Property Get Parent() As clXLSB
    Set Parent = prParent
End Property

Property Set Parent(nvParent As clXLSB)
    Set prParent = nvParent
End Property

Property Get Id() As String
    Id = prID
End Property

Property Let Id(nvID As String)
    prID = nvID
End Property

Property Get tableSynth() As Variant
    tableSynth = prtableSynth
End Property

Property Let tableSynth(nvtableSynth As Variant)
    prtableSynth = nvtableSynth
End Property

Property Get tablePRP() As Variant
    tablePRP = prTablePRP
End Property

Property Let tablePRP(nvtablePRP As Variant)
    prTablePRP = nvtablePRP
End Property

Property Get Financement() As Variant
    Financement = prFinancement
End Property

Property Let Financement(nvFinancement As Variant)
    prFinancement = nvFinancement
End Property
Property Get Locatif() As Variant
    Locatif = prLocatif
End Property

Property Let Locatif(nvLocatif As Variant)
    prLocatif = nvLocatif
End Property

Property Get PRPDetail() As Variant
    PRPDetail = prPRPDetail
End Property

Property Let PRPDetail(nvPRPDetail As Variant)
    prPRPDetail = nvPRPDetail
End Property

Property Get PRPDetailTitre() As Variant
    PRPDetailTitre = prPRPDetailTitre
End Property

Property Let PRPDetailTitre(nvPRPDetailTitre As Variant)
    prPRPDetailTitre = nvPRPDetailTitre
End Property



Property Get AllPRP() As Variant
    Dim a, b, c
    
    a = Transpose(prPRPDetail)
    b = Transpose(prPRPDetailTitre)
    ReDim c(LBound(a, 1) To UBound(a, 1), 0 To 2)
    
    For i = LBound(a, 1) To UBound(a, 1)
        For j = 0 To 2
            If j = 0 Then c(i, j) = a(i, j)
            If j = 1 Then c(i, j) = a(i, j)
            If j = 2 Then c(i, j) = b(i, j - 2)
        Next j
    Next i
    
    ReDim a(LBound(c, 1) To UBound(c, 1), LBound(c, 2) To UBound(c, 2) + 4)
    
    For i = LBound(a, 1) To UBound(a, 1)
        For j = LBound(a, 2) To UBound(a, 2)
            If j = 0 Then
                a(i, j) = prID
            ElseIf j = 1 Then
                a(i, j) = prtableSynth(1, 4)
            ElseIf j = 2 Then
                a(i, j) = prtableSynth(1, 6)
            ElseIf j = 3 Then
                a(i, j) = prtableSynth(1, 7)
            Else
                a(i, j) = c(i, j - 4)
            End If
        Next j
    Next i
    
    a(0, 0) = "ID"
    a(0, 1) = "NB"
    a(0, 2) = "SHAB"
    a(0, 3) = "SHABDOM"
    a(0, 4) = "POSTE"
    a(0, 5) = "VALEUR"
    a(0, 6) = "REGROUPEMENT"
    
    prAllPRP = a
    AllPRP = prAllPRP
End Property

Private Function Concataine(a As Variant, b As Variant)
Dim c

If UBound(a, 2) - LBound(a, 2) = 1 Then a = Transpose(a)
If UBound(b, 2) - LBound(b, 2) = 1 Then b = Transpose(b)

ReDim c(LBound(a, 1) To UBound(a, 1), LBound(a, 2) To UBound(a, 2) + UBound(b, 2) + 1)

For i = LBound(c, 1) To UBound(c, 1)
    For j = LBound(c, 2) To UBound(c, 2)
        If j <= UBound(a, 2) Then
            c(i, j) = a(i, j)
        Else
            c(i, j) = b(i, j - UBound(a, 2) - 1)
        End If
    Next j
Next i

Concataine = c

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


Private Sub Class_Initialize()
    Dim Table(0 To 1, 0 To 0)
    
End Sub
