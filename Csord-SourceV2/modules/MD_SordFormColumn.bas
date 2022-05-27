Attribute VB_Name = "MD_SordFormColumn"
' ------------------------------------------------------
' Name:    MD_SordFormColumn
' Kind:    Module
' Purpose: Module pour l'utilisation de la classe CSordFormColumn.
' Author:  Laurent
' Date:    26/05/2022 11:42
' DateMod: 26/05/2022 21:51
' ------------------------------------------------------
Option Compare Database
Option Explicit

'//::::::::::::::::::::::::::::::::::    VARIABLES      ::::::::::::::::::::::::::::::::::
    Private CSordForm   As CsordFormColumn
    Private oldFrmName  As String
'//:::::::::::::::::::::::::::::::::: END VARIABLES ::::::::::::::::::::::::::::::::::::::

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' ----------------------------------------------------------------
' Procedure Nom:    SordColumn
' Sujet:            Initialisation de la classe, lance le tri sur la colonne.
' Procedure Kind:   Function
' Procedure Access: Public
' Références:       Initialisation de la classe, lance le tri sur la colonne.
'
'=== Paramètres ===
'==================
'
' Return Type: Boolean
'
' Author:  Laurent
' Date:    26/05/2022 - 11:41
' DateMod:
' ----------------------------------------------------------------
Public Function SordColumn() As Boolean

On Error GoTo ERR_SordColumn

    Dim bRet As Boolean

    DoCmd.Echo False

    If (oldFrmName <> Screen.ActiveForm.Name) Then Set CSordForm = Nothing
    '// Initialisation de la classe, on peut indiquer, si besoin, le préfixe et/ou le suffixe (nb de car).
    '// Init class and defined suffix (the class cuts automatically the button name for extact field name)
    If (CSordForm Is Nothing) Then
        Set CSordForm = New CsordFormColumn
        With CSordForm
            '// Applique images défini lors de la création du form.
            .PicturePath = "PICFOLDER"
            .PictureASC = "PICIMGASC"
            .PictureDESC = "PICIMGDESC"
            oldFrmName = Screen.ActiveForm.Name
        End With
    End If

    bRet = CSordForm.SordNow()      '// Execute le tri, retour TRUE if ok,

    If (bRet = False) Then
        '// Your code here
        '// Your code here
    End If

SORTIE_SordColumn:
    DoCmd.Echo True
    Exit Function

ERR_SordColumn:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  CSord.MD_SordFormColumn.SordColumn, ligne " & Erl & "."
    Resume SORTIE_SordColumn
End Function

Public Function CloseSordColumn()
    Set CSordForm = Nothing
    oldFrmName = vbNullString
End Function
'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ END PUB. SUB/FUNC \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
