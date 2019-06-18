Attribute VB_Name = "LangtonAntAlgo"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Langton()
Attribute Langton.VB_ProcData.VB_Invoke_Func = "l\n14"
    Application.ScreenUpdating = True 'on s'assure que chaque mouvement de la fourmi est affich� imm�diatement
    'LANCER LA MACRO AVEC F5
    Dim Fourmi As Range 'La fourmi est une cellule Excel
    ActiveWindow.Zoom = 40 'on d�zoome
    Cells.Clear 'on efface les cellules et le trac� pr�c�dent
    Cells.Interior.Color = vbWhite 'on met les cellules en noir pour �viter de voir le quadrillage
    Cells.Font.Color = vbRed 'on met la police des cellules en rouge pour la rendre facilement lisible
    
    'on rend les cellules � peu pr�s carr�es
    Columns.ColumnWidth = 2
    Rows.RowHeight = 14
    
    Set Fourmi = Range("Z30") 'On part d'une cellule Excel un peu au hasard, la "Z30"
    DirectionFourmi = "Gauche" 'on part d'une des 4 directions possibles
    
    For i = 1 To 12000 'Nombre de d�placements de la souris
        Fourmi.Value = Fourmi.Value + 1 'On �crit dans la cellule sur laquelle est la fourmi le nombre d�j� �crit + 1 (pour savoir combien de fois la fourmi passe dans cette cellule)
    
        Select Case Fourmi.Interior.Color 'on regarde la couleur de la cellule sur laquelle est la fourmi
            Case vbWhite 'si la cellule est blanche
            Fourmi.Interior.Color = vbBlack 'alors elle devient noire
            Select Case DirectionFourmi 'on regarde �galement la direction
                Case "Gauche" 'si elle allait vers la gauche
                    DirectionFourmi = "Bas" 'alors elle va maintenant vers le bas
                    Set Fourmi = Fourmi.Offset(-1, 0) 'et elle se d�place d'une case vers le bas
                Case "Droite"
                    DirectionFourmi = "Haut"
                    Set Fourmi = Fourmi.Offset(1, 0)
                Case "Haut"
                    DirectionFourmi = "Gauche"
                    Set Fourmi = Fourmi.Offset(0, -1)
                Case "Bas" 'si elle allait vers le bas
                    DirectionFourmi = "Droite" 'alors elle va maintenant vers la droite
                    Set Fourmi = Fourmi.Offset(0, 1) 'et elle se d�place d'une case vers la droite
            End Select
    
            Case vbBlack 'si la cellule n'est pas blanche mais est noire, on fait la m�me chose que dans le cas pr�c�dent, mais en tournant vers la droite
            Fourmi.Interior.Color = vbWhite
            Select Case DirectionFourmi
                Case "Gauche"
                    DirectionFourmi = "Haut"
                    Set Fourmi = Fourmi.Offset(1, 0)
                Case "Droite"
                    DirectionFourmi = "Bas"
                    Set Fourmi = Fourmi.Offset(-1, 0)
                Case "Haut"
                    DirectionFourmi = "Droite"
                    Set Fourmi = Fourmi.Offset(0, 1)
                Case "Bas"
                    DirectionFourmi = "Gauche"
                    Set Fourmi = Fourmi.Offset(0, -1)
            End Select
        End Select
    If i < 30 Then
        Sleep 1000
    ElseIf i < 100 Then
        Sleep 500
    ElseIf i < 350 Then
        Sleep 100
    ElseIf i < 1000 Then
        Sleep 10
    Else
        Sleep 0
    End If
    Application.StatusBar = "It�ration " & i
    Next i 'on relance la boucle jusqu'� atteindre le nombre d'it�rations sp�cifi�
End Sub

