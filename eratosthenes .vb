Dim Dimension As Integer ' Correspond au dimension entrer par l'utilisateur
Dim Cell As Range ' Définir la cellule "témoin"
Public Sub Construction_Tableau()
'
' Construction_Tableau Macro
'
Set Cell = Range("A1")
Dimension = InputBox("Entrez le nombre de chiffe du tableau", "Eratosthène", 100)

' Construire tableau
For i = 1 To Dimension
    Cell.Offset(i).Value = i
    Cell.Offset(i).Interior.Color = RGB(0, 102, 0) ' Vert
Next

' Mise en forme des bordures
With Range("A1:A" & Dimension).CurrentRegion
    'bordure de gauche
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    'bordure du dessus
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    'bordure du dessous
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    'bordure de droite
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    'bordure des interlignes
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    'bordure des inter colonnes
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End With

' Démarrage du timer: temps execution
Dim sngChrono As Single
sngChrono = Timer

' Crible d'Erathostène
i = 2
Dim j As Integer

' Met par défault la 1er cellule en rouge
Cell.Offset(1).Interior.Color = RGB(255, 0, 0)

Do While (i ^ 2 < Dimension)
    If (Cell.Offset(i).Interior.Color = RGB(0, 102, 0)) Then
        j = 2 * i
        Do While (j < Dimension)
            Cell.Offset(j).Interior.Color = RGB(255, 0, 0)
            j = j + i
        Loop
    End If
    i = i + 1
Loop

' Calcule du temps d'execution
sngChrono = Timer - sngChrono
MsgBox "Temps d'execution du code en : " & CStr(sngChrono * 1000) & " ms"

End Sub
Public Sub Detruire_Tableau()
'
' Détruire tableau
'
Set Cell = Range("A1") ' Sélectionne la cellule "témoin"
For i = 1 To Dimension
    Cell.Offset(i).Value = "" ' Remise zero des valeurs
    Cell.Offset(i).Interior.Color = RGB(255, 255, 255) ' Passe en blanc
    Cell.Offset(i).Interior.Pattern = xlNone ' Bordure par défault style excel
Next

' Mise en forme des bordures
Range("A1:A" & Dimension).Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
Selection.Borders(xlEdgeLeft).LineStyle = xlNone
Selection.Borders(xlEdgeTop).LineStyle = xlNone
Selection.Borders(xlEdgeBottom).LineStyle = xlNone
Selection.Borders(xlEdgeRight).LineStyle = xlNone
Selection.Borders(xlInsideVertical).LineStyle = xlNone
Selection.Borders(xlInsideHorizontal).LineStyle = xlNoneEnd

End Sub

' Enregistre une feuille
Public Sub EnregistrerUneFeuille()
Dim numero As Integer
Dim nom As String
Dim alerte As String
numero = Val(InputBox("Numéro de la feuille à enregistrer", "Numéro"))
If numero = 0 Then
    alerte = MsgBox("Saisissez un nombre supérieur à 0 !", vbCritical, "Attention")
    Exit Sub
End If

If numero > Sheets.Count Then
    alerte = MsgBox("Le nombre saisi est supérieur au nombre de feuilles du classeur !", vbCritical, "Attention")
    Exit Sub
End If

nom = ThisWorkbook.Path & "\" & Sheets(numero).Name
Sheets(numero).Copy
ActiveWorkbook.SaveAs Filename:=nom

End Sub