Option Explicit

Const PREMIERE_COLONNE_CHOIX = 8
Const NB_POSTES_DANS_SEMAINE = 4

Sub ResetPoste()
'
' Keyboard Shortcut: Ctrl+d
'
    Dim activeColonne As Integer
    Dim activeLigne As Integer
    Dim indicePosteDansSemaine As Integer
    Dim refColonne As Integer
    Dim refCellAdresse As String
    activeColonne = ActiveCell.Column
    If activeColonne < PREMIERE_COLONNE_CHOIX Then Exit Sub
    'MsgBox activeColonne
    activeLigne = ActiveCell.Row
    indicePosteDansSemaine = activeColonne Mod NB_POSTES_DANS_SEMAINE
    refColonne = PREMIERE_COLONNE_CHOIX - NB_POSTES_DANS_SEMAINE + indicePosteDansSemaine
    
    refCellAdresse = "RC" & refColonne
    ActiveCell.FormulaR1C1 = "=IF(" & refCellAdresse & "="""",""""," & refCellAdresse & ")"
End Sub
