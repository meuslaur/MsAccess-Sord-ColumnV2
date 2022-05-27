VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsordFormColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Utilisation")
' ------------------------------------------------------
' Name:     CSordFormColumn
' Kind:     Module de classe
' Purpose:  Tri sur les colonnes d'un form en continu.
' Author:   Laurent
' Date:     18/04/2022
' DateMod:  02/05/2022-10:04
'
'   Utilisation dans le Formulaire :
'   ==============================
'
'   Event Form_Close=CloseSordColumn()    '// Recommended, clean mem.
'
'
'   Utilisation dans le Controle :
'   ============================
'
'   Event On_Click=SordColumn()
'
'   Function in MD_SordFormColumn
'   ==============================
'
'Private Function SordColumn() As Boolean
'
'    Dim bRet As Boolean
'
'    '// Initialisation de la classe.
'    If (m_CSordForm Is Nothing) Then
'        Set m_CSordForm = New CsordFormColumn
'        .PicturePath = "C:\MyApp\Res\"     '// dossier des images)
'        .PictureASC = "PicAsc"             '// Img Sord ASC
'        .PictureDESC = "PicDesc"           '// Img sord DESC
'
'        bRet = .SordNow()                  '// Execute le tri, retour TRUE if ok.
'
'    End With
'
'    If (bRet = False) Then
'        '// Your code here
'        '// Your code here
'    End If
'
'End Function
' ------------------------------------------------------
Option Compare Database
Option Explicit

    '// Defaut paramètres.
    Private Const SORD_ASC      As String = " ASC"
    Private Const SORD_DESC     As String = " DESC"
    Private Const CMB_PREFIXE   As String = "cmb_"

    Private Enum T_Sord
        eAsc = 1
        eDESC = 2
    End Enum
    Private m_eCurrentSord      As T_Sord           '// Sens du tri en cours (1 asc, 2 desc).

    Private m_bCheckOk          As Boolean          '// Flag if error occured.

    Private m_objForm           As Form             '// Form parent du control(sur lequel appliquer le tri).

    Private m_objButton         As CommandButton    '// Control bouton cliquer à utiliser.
    Private m_sButtonOldName    As String           '// Nom du control précedement modifier.

    Private m_sFieldName        As String           '// Champ sur lequel appliquer le tri.

    '// Paramètre pour m_objButton.
    Private m_sPicturePath      As String
    Private m_sPictureAsc       As String           '// Image tri pour ASC.
    Private m_sPictureDesc      As String           '// Image tri pour DESC.
    Private m_sPictureSord      As String           '// Image a afficher suivant le tri en cours (ASC ou DESC).

    '// Sauvegarde des propriétés du form/button avant modification.
    Private m_sSvFrmSord        As String           '// Sauvegarde les paramètres de tri d'origine.
    Private m_bSvFrmSortOn      As Boolean          '//             ""

    Private m_eSvPicArr         As AcPictureCaptionArrangement
    Private m_SvPicType         As Byte             '// 0 incorporée, 1 liée
    Private m_vSvPicData        As Variant          '// Pour PicType = 0 (pictureData)
    Private m_sSvPicFile        As String           '// Pour PicType = 1 (Path and picture file)

'//----------------------------------       INIT/TERM       ----------------------------------
Private Sub Class_Initialize()

    '// TOUT est OK par défaut.
    m_bCheckOk = True

End Sub

Private Sub Class_Terminate()

    '// Restaure les filtres du form.
    If (Not m_objForm Is Nothing) Then
        m_objForm.OrderBy = m_sSvFrmSord
        m_objForm.OrderByOn = m_bSvFrmSortOn
    End If

    Set m_objButton = Nothing
    Set m_objForm = Nothing

End Sub
'//---------------------------------------------------------------------------------------

'//==================================       PROP        ==================================

'// Picture Path
Public Property Get PicturePath() As String
    PicturePath = m_sPicturePath
End Property
Public Property Let PicturePath(NewPath As String)

    '// Evite les messages redondants.
    If (m_bCheckOk = False) Then Exit Property
    If (NewPath = vbNullString) Then Exit Property

    '// Le dossier existe ?
    If (CheckFolderExists(NewPath) = False) Then
        MsgBox "Sous-Dossier '" & NewPath & "', non trouver", vbExclamation + vbOKOnly, "PicturePath"
        m_bCheckOk = False
        Exit Property
    End If

    '// Vérifier si fini bien par '\'.
    If (Right$(NewPath, 1) <> "\") Then NewPath = NewPath & "\"

    m_sPicturePath = NewPath

End Property

'// Picture for sord asc
Public Property Get PictureASC() As String
    PictureASC = m_sPictureAsc
End Property
Public Property Let PictureASC(NewFileName As String)

    Dim sFullPath As String

    If (NewFileName = vbNullString) Then Exit Property

    '// Img vide, rien a faire on sort.
    '// Evite les messages répétitif d'érreur.
    If (m_sPicturePath = vbNullString) Then Exit Property

    '// Le fichier existe ?
    sFullPath = m_sPicturePath & NewFileName
    If (CheckFileExist(sFullPath) = False) Then
        MsgBox "Fichier " & vbCrLf & NewFileName & vbCrLf & " non trouver.", vbExclamation, "PictureASC"
        m_bCheckOk = False
        Exit Property
    End If

    m_sPictureAsc = NewFileName

End Property

'// Picture for sord desc
Public Property Get PictureDESC() As String
    PictureDESC = m_sPictureDesc
End Property
Public Property Let PictureDESC(NewFileName As String)

    Dim sFullPath As String

    If (NewFileName = vbNullString) Then Exit Property

    '// Img vide, rien a faire on sort.
    '// Evite les messages répétitif d'érreur.
    If (m_sPicturePath = vbNullString) Then Exit Property

    '// Le fichier existe ?
    sFullPath = m_sPicturePath & NewFileName
    If (CheckFileExist(sFullPath) = False) Then
        MsgBox "Fichier " & vbCrLf & NewFileName & vbCrLf & " non trouver.", vbExclamation, "PictureDESC"
        m_bCheckOk = False
        Exit Property
    End If

    m_sPictureDesc = NewFileName

End Property

Private Property Get GetFSO() As Object

    Static pFSO               As Object
    
    If pFSO Is Nothing Then
        Set pFSO = CreateObject("Scripting.FileSystemObject")
    End If

    Set GetFSO = pFSO

End Property

'//=======================================================================================

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Function SordNow() As Boolean
    On Error GoTo ERR_SordNow

    Dim bRet As Boolean

    DoCmd.Echo False    'TODO: DoCmd.Echo

    'TODO: test le parent du bouton si un sf
    '// Initialisation du form, si pas déjà fait...
    If (CheckForm = False) Then GoTo SORTIE_SordNow     '// Erreur on sort

    '// Initialise le bouton, si pas deja fait...
    If (CheckButton(Screen.ActiveControl) = False) Then GoTo SORTIE_SordNow     '// Erreur on sort.

    '// Défini le champs, si changement de CommandButton...
    If (m_sButtonOldName <> m_objButton.Name) Then
        bRet = SetFieldName(m_objButton.Name)
        If (bRet = False) Then GoTo SORTIE_SordNow      '// Erreur on sort.
    End If

    '// On inverse le tri.
    Dim sCurSord As String

    '// Determine le texte et l'image as utilisés suivant le sens du tri.
    Select Case m_eCurrentSord
        Case eAsc
            m_eCurrentSord = eDESC
            sCurSord = SORD_ASC
            m_sPictureSord = m_sPicturePath & m_sPictureAsc
        Case eDESC
            m_eCurrentSord = eAsc
            sCurSord = SORD_DESC
            m_sPictureSord = m_sPicturePath & m_sPictureDesc
    End Select

    '// Initialise le control cliquer...
    SavePropButton      '// Sauvegarde le bouton en cours...
    UpdatePropButton    '// MàJ du bouton en cours...

    '// Applique le tri.
    m_objForm.OrderBy = m_sFieldName & sCurSord
    m_objForm.OrderByOn = True
    
    SordNow = True

SORTIE_SordNow:
    DoCmd.Echo True     'TODO: DoCmd.Echo
    Exit Function

ERR_SordNow:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.CSordFormColumn.SordNow, ligne " & Erl & "."
    Resume SORTIE_SordNow
End Function
'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'// ######################## PRIVATE SUB/FUNC ###########################################
' ----------------------------------------------------------------
' Procedure Nom:    SetFieldName
' Sujet:            Initialisation du champ de tri.
'                   Si même nom que précedement, on fait rien.
'                   On vérifie si le nom est valide dans la table.
'                   On remet les options aux valeurs par défaut.
' Procedure Kind:   Sub
' Procedure Access: Private
'
' Author:  Laurent
' Date:    01/05/2022 - 14:42
' DateMod: 26/05/2022 11:51
'
' !Use! : CheckPrefixSuffix, CheckFieldName, RestorePropButton.
' ----------------------------------------------------------------
Private Function SetFieldName(NewFieldName As String) As Boolean

    '// Meme champ, rien a faire on sort.
    If (m_sFieldName = NewFieldName) Then Exit Function

    Dim sTmpField As String

    sTmpField = NewFieldName

    '// Extraction du nom du champs, on utilise le nom du commandButton sans le préfixe.
    sTmpField = ExtraireFiledName(m_objButton.Name)
    If (sTmpField = vbNullString) Then Exit Function

    '// ...Vérifier la validité du champ...
    If (CheckFieldName(sTmpField) = False) Then Exit Function

    '// Nouveau champ, Restaure le commandButton précédent...
    RestorePropButton

    '// Place les options par défaut, remet le tri a DESC.
    m_eCurrentSord = eAsc

    '// On stock le nom du champ.
    m_sFieldName = sTmpField

    SetFieldName = True

End Function

Private Sub SavePropButton()

    '// Rien a faire, meme bouton, on sort.
    If (m_sButtonOldName = m_objButton.Name) Then Exit Sub

    '// Sauvegarde des propriétés, si nouveau bouton...
    m_SvPicType = m_objButton.PictureType
    m_eSvPicArr = m_objButton.PictureCaptionArrangement
    Select Case m_SvPicType
        Case 0
            m_vSvPicData = m_objButton.PictureData
        Case Else
            m_sSvPicFile = m_objButton.Picture
    End Select

    m_sButtonOldName = m_objButton.Name

End Sub

Private Sub RestorePropButton()
    
    '// Rien a faire, 1er bouton ou même bouton, on sort.
    If ((m_sButtonOldName = vbNullString) Or (m_sButtonOldName = m_objButton.Name)) Then Exit Sub

    '// Image.
    m_objForm(m_sButtonOldName).PictureType = m_SvPicType
    m_objForm(m_sButtonOldName).PictureCaptionArrangement = m_eSvPicArr

    Select Case m_SvPicType
        Case 0
            m_objForm(m_sButtonOldName).PictureData = m_vSvPicData
        Case Else
            m_objForm(m_sButtonOldName).Picture = m_sSvPicFile
    End Select

End Sub

Private Sub UpdatePropButton()
'// Definir l'image à afficher sur le CommandButton.
    m_objButton.PictureCaptionArrangement = acLeft
    m_objButton.PictureType = 1
    m_objButton.Picture = m_sPictureSord
    
End Sub

' ----------------------------------------------------------------
' Procedure Nom:    CheckForm
' Sujet:            Vérifier si le form as bien était initialiser
' Procedure Kind:   Sub
' Procedure Access: Private
' Return Type:      Boolean TRUE si pas de problème.
' Author:           Laurent
' Date:             20/04/2022 - 06:37
' DateModif:        22/04/2022 - 09:04
' ----------------------------------------------------------------
Private Function CheckForm() As Boolean
    On Error GoTo ERR_CheckForm

    '// Initialise le form sur lequel appliquer le tri,
    '// Sauvegarde des filtres en cours.
    If (m_objForm Is Nothing) Then
        'TODO: Test si le form est un SF.
        If (Screen.ActiveControl.Parent.FormName <> Screen.ActiveForm.Name) Then
            Set m_objForm = Screen.ActiveForm.ActiveControl.Form  ' Application.Forms.Item(Screen.ActiveControl.Parent.FormName)
        Else
            Set m_objForm = Screen.ActiveForm
        End If
        m_sSvFrmSord = m_objForm.OrderBy
        m_bSvFrmSortOn = m_objForm.OrderByOn
    End If

    CheckForm = True

SORTIE_CheckForm:
    Exit Function

ERR_CheckForm:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            "(" & Err.Description & ")" & vbCrLf & _
            "Dans  CSordFormColumn.CheckForm," & vbCrLf & _
            Switch(Erl = 0, vbNullString, Erl <> 0, vbCrLf & "Ligne No: " & Erl)
    Resume SORTIE_CheckForm
End Function

' ----------------------------------------------------------------
' Procedure Nom:    CheckButton
' Sujet:            Vérifier si bouton défini suivant les options choisi.
' Procedure Kind:   Sub
' Procedure Access: Private
' Parameter NewButton (CommandButton):
' Return Type:      Boolean TRUE si pas de problème.
' Author:           Laurent
' Date:             20/04/2022 - 06:45
' DateModif:        22/04/2022 - 09:06
' ----------------------------------------------------------------
Private Function CheckButton(NewButton As CommandButton) As Boolean
    On Error GoTo ERR_CheckButton

    '// Initialise le bouton.
    If (m_objButton Is Nothing) Then
        Set m_objButton = NewButton
    ElseIf (m_objButton.Name <> NewButton.Name) Then
        Set m_objButton = NewButton
    End If

    CheckButton = True

SORTIE_CheckButton:
    Exit Function

ERR_CheckButton:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  SordFormCc.CSordFormColumn.CheckButton, ligne " & Erl & "."
    Resume SORTIE_CheckButton
End Function

' ----------------------------------------------------------------
' Procedure Nom:    CheckFieldName
' Sujet:            Vérifier si la table contient des enrg et si nom du champ valide...
'                   Retourne TRUE Si tout vas bien.
' Procedure Kind:   Sub
' Procedure Access: Private
'
'=== Paramètres ===
' sField (String): Nom du champ à vérifier.
'==================
'
' Author:  Laurent
' Date:    19/04/2022 - 11:35
' DateMod: 01/05/2022 - 13:4
'
' ----------------------------------------------------------------
Private Function CheckFieldName(ByRef sField As String) As Boolean
    On Error GoTo ERR_CheckFieldName

    CheckFieldName = IIf((DCount(sField, m_objForm.RecordSource) = 0), False, True)

SORTIE_CheckFieldName:
    Exit Function

ERR_CheckFieldName:
    If ((Err.Number = 3265) Or (Err.Number = 2471)) Then
        MsgBox "Le champ '" & sField & "' n'existe pas dans la source de donnée, vérifier la valeur de .FieldName", vbExclamation, "Classe SordFormColumn"
    Else
        MsgBox "Erreur " & Err.Number & " (" & Err.Description & ") Dans CheckFieldName." & Switch(Erl = 0, vbNullString, Erl <> 0, vbCrLf & "Ligne No: " & Erl)
    End If
    Resume SORTIE_CheckFieldName
End Function

Private Function CheckFileExist(ByVal sFullPathFile As String) As Boolean

    Dim oFSO            As Object

    Set oFSO = GetFSO

    CheckFileExist = oFSO.FileExists(sFullPathFile)

    Set oFSO = Nothing

End Function


'Returns True if the folder exists (and is accessible)
' - trailing backslash is completely optional
' - returns False if the full path to an existing file is passed
'   to the function (and not just the folder part)
Private Function CheckFolderExists(ByVal PathToFolder As String) As Boolean

    Dim oFSO As Object
    Dim bRes As Boolean

    Set oFSO = GetFSO
    bRes = oFSO.FolderExists(PathToFolder)

    Set oFSO = Nothing
    CheckFolderExists = bRes
    Set oFSO = Nothing

End Function

' ----------------------------------------------------------------
' Procedure Nom:    ExtraireFiledName
' Sujet:            Supprime le préfixe du nom du commandButton pour extraire le nom du champs à utiliser.
' Procedure Kind:   Function
' Procedure Access: Private
'
'=== Paramètres ===
' sCmbName (String): Nom du commandButton
'==================
'
' Return Type: String Retourne le nom du champs à utiliser.
'
' Author:   Laurent
' Date:     26/05/2022 19:29
' DateMod:
' ----------------------------------------------------------------
Private Function ExtraireFiledName(sCmbName As String) As String

    Dim sResult As String
    Dim lTemp   As Long

    lTemp = InStr(1, sCmbName, CMB_PREFIXE)
    If (lTemp = 0) Then Exit Function

    '// Extraire le nom du champ.
    lTemp = Len(sCmbName)
    sResult = Right$(sCmbName, lTemp - 4)

    ExtraireFiledName = sResult

End Function

'// #####################################################################################
