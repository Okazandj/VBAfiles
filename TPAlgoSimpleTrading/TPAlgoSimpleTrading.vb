Public Portefeuille
Public ValeurAction
Private Sub Valider_Click()

' Variable portefeuille
If (IsNumeric(TextBox1.Value)) Then
    Portefeuille = TextBox1.Value
Else
    MsgBox "Veuillez entrer un nombre", vbCritical, "Type incorrect"
End If

End Sub
Private Sub DebutTrading_Click()
On Error GoTo erreur

' Selection de la feuille
Sheets("CAC40").Select
' Suppression des anciens graphiques
ActiveSheet.DrawingObjects.Delete

If (Portefeuille <> 0) Then
    MsgBox "Début du trading avec un portefeuille de " & Portefeuille
    ' Selection du produit
    Sheets("CAC40").Select
    Range("I1").Value = "Acheter"
    Range("J1").Value = "Vendre"
    ValeurAction = 0
    ' Début de notre algorithme de trading
    
    ' Parcoure les données du tableau de la plus ancienne à la plus récente
    For i = 2 To 66 Step (1)
        ' Prise de décision
    
        ' Si on possède une action, on vend
        If ValeurAction <> 0 Then
            Range("J" & i).Value = Range("B" & i).Value
            Portefeuille = Portefeuille - Range("B" & i).Value
            ValeurAction = 0
        End If
        
        ' Si on a assez d'argent pour acheter l'action et que la valeur de l'action d'hier était plus chère => on veut acheter
        If Range("B" & i).Value < Portefeuille And Range("B" & i - 1).Value < Range("B" & i).Value Then
            Range("I" & i).Value = Range("B" & i).Value
            Portefeuille = Portefeuille + Range("B" & i).Value
            ValeurAction = Range("B" & i).Value
        End If
        
        ' N'affiche pas les transactions d'achat et vente en même temps (Graphique)
        If Range("I" & i).Value <> 0 And Range("J" & i).Value <> 0 Then
            Range("I" & i).Value = ""
            Range("J" & i).Value = ""
        End If
        
    Next
    
    ' Bilan
    
    ' Indique la valeur du portefeuille après la procédure
    MsgBox "Valeur actuelle du portefeuille: " & Portefeuille
    
    ' Indique les pertes ou les gains de la procédure
    If Portefeuille < TextBox1.Value Then
        MsgBox "Gain de " & Portefeuille - TextBox1.Value
    Else
        MsgBox "Perte de " & TextBox1.Value - Portefeuille
    End If
    
    ' S'il possède une action, on lui indique sa valeur
    If ValeurAction <> 0 Then
        MsgBox "Vous possèdez encore 1 action CAC40 d'une valeur actuelle de " & ValeurAction
    End If
    
Else
    MsgBox "Veuillez entrer votre montant dans le portefeuille", vbCritical, "Valeur inconnu"
End If

Exit Sub

erreur:
    MsgBox "Erreur", vbExclamation, ""
End Sub
Private Sub FinTrading_Click()
    End
End Sub
Private Sub Affichedoc_Click()
' Possible d'avoir des erreurs à cause du Path de l'utilisateur (si ' ' ou char. spéc)
On Error GoTo erreur
    Sheets("CAC40").Select
    ' Création du fichier
    Open ThisWorkbook.Path & "\document.txt" For Output As #1
    ' Ecriture dans le fichier
    For i = 2 To 67 Step (1)
        If Range("I" & i).Value <> 0 Then
            Print #1, Range("A" & i).Value & " Achat d'une action CAC40 d'une valeur de: " & Range("I" & i).Value
        ElseIf Range("J" & i).Value <> 0 Then
            Print #1, Range("A" & i).Value & " Vente d'une action CAC40 d'une valeur de: " & Range("J" & i).Value
        End If
    Next
    ' Fermeture du fichier
    Close #1
    Set sh = CreateObject("WScript.Shell")
    sh.Run (ThisWorkbook.Path & "\document.")

Exit Sub

' Message d'erreur
erreur:
    MsgBox "Attention erreur: Le chemin de votre dossier comprend surement des charactères spéciaux. Impossible de l'ouvrir.", vbExclamation, "Erreur !"
End Sub
Private Sub Graphique_Click()
    ' Selection de la feuille
    Sheets("CAC40").Select
    ' Suppression des anciens graphiques
    ActiveSheet.DrawingObjects.Delete
    ' Construction du tableau
    Range("A1:B67,I1:J67").Select
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
    ActiveChart.SetSourceData Source:=Range("CAC40!$A$1:$B$67,CAC40!$I$1:$J$67" _
        )
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = _
        "Graphique représentant les actions du processus de trading"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 58).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 58).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    
    
End Sub
Private Sub Envoyer_Click()
' Attention Outlook est obligatoire pour envoyer les mail
    ActiveWorkbook.SendMail "kazandji@yahoo.fr", "Tp VBA Olivier Kazandji Nicolas Huang-Dubois ESILV S7", True
End Sub
Private Sub Imprimer_Click()
' Imprime la page actuelle CAC40 en utilisant l'imprimante par défault
    Sheets("CAC40").Select
    ActiveSheet.PrintOut
End Sub
