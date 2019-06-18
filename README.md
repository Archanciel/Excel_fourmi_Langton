# excel_Langton
Implémentation dans Excel - VBA de l'algorithme de la fourmi de Langton  

## Youtube video
https://www.youtube.com/watch?v=qZRYGxF6D3w

## Auteur du code et son commentaire
Quentin Domon - 2016

Si certains veulent essayer, il vous suffit d'avoir Excel, d'ouvrir l'éditeur Visual Basic 
avec ALT+F11 une fois sur Excel, de copier le texte ci-dessous (de Sub... à End Sub) 
dans la page de code "ThisWorkbook" et de la lire en appuyant sur F5

(on pourrait mieux écrire ça, mais je l'ai fait en 2 minutes dès que j'ai vu le début de la 
vidéo pour voir moi-même et j'ai été impressionné par l'"autoroute" avant même de voir le 
phénomène tel que décrit dans la vidéo :))

## VBA code
```
Sub Langton()
    Application.ScreenUpdating = True 'on s'assure que chaque mouvement de la fourmi est affiché immédiatement
    'LANCER LA MACRO AVEC F5
    Dim Fourmi As Range 'La fourmi est une cellule Excel
    ActiveWindow.Zoom = 40 'on dézoome
    Cells.Clear 'on efface les cellules et le tracé précédent
    Cells.Interior.Color = vbWhite 'on met les cellules en noir pour éviter de voir le quadrillage
    Cells.Font.Color = vbRed 'on met la police des cellules en rouge pour la rendre facilement lisible
    
    'on rend les cellules à peu près carrées
    Columns.ColumnWidth = 2
    Rows.RowHeight = 14
    
    Set Fourmi = Range("Z30") 'On part d'une cellule Excel un peu au hasard, la "Z30"
    DirectionFourmi = "Gauche" 'on part d'une des 4 directions possibles
    
    For i = 1 To 20000 'Nombre de déplacements de la souris
        Fourmi.Value = Fourmi.Value + 1 'On écrit dans la cellule sur laquelle est la fourmi le nombre déjà écrit + 1 (pour savoir combien de fois la fourmi passe dans cette cellule)
    
        Select Case Fourmi.Interior.Color 'on regarde la couleur de la cellule sur laquelle est la fourmi
            Case vbWhite 'si la cellule est blanche
            Fourmi.Interior.Color = vbBlack 'alors elle devient noire
            Select Case DirectionFourmi 'on regarde également la direction
                Case "Gauche" 'si elle allait vers la gauche
                    DirectionFourmi = "Bas" 'alors elle va maintenant vers le bas
                    Set Fourmi = Fourmi.Offset(-1, 0) 'et elle se déplace d'une case vers le bas
                Case "Droite"
                    DirectionFourmi = "Haut"
                    Set Fourmi = Fourmi.Offset(1, 0)
                Case "Haut"
                    DirectionFourmi = "Gauche"
                    Set Fourmi = Fourmi.Offset(0, -1)
                Case "Bas" 'si elle allait vers le bas
                    DirectionFourmi = "Droite" 'alors elle va maintenant vers la droite
                    Set Fourmi = Fourmi.Offset(0, 1) 'et elle se déplace d'une case vers la droite
            End Select
    
            Case vbBlack 'si la cellule n'est pas blanche mais est noire, on fait la même chose que dans le cas précédent, mais en tournant vers la droite
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
    Next i 'on relance la boucle jusqu'à atteindre le nombre d'itérations spécifié
End Sub
```
