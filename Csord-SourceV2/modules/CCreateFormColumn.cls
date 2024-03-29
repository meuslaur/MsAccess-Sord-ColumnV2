﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCreateFormColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Creation")
' ------------------------------------------------------
' Name:     CCreateFormColumn
' Kind:     Module de classe
' Purpose:  Création d'un formulaire pour la classe CsordFormColumn
' Author:   Laurent
' Date:     27/04/2022
' DateMod:  26/05/2022 11:40
' ------------------------------------------------------
Option Compare Database
Option Explicit

'//::::::::::::::::::::::::::::::::::    VARIABLES      ::::::::::::::::::::::::::::::::::
    Private Const MOD_CLASS     As String = "CSordFormColumn"       '// Nom du module de classe à importer.
    Private Const MOD_STAND     As String = "MD_SordFormColumn"     '// Nom du module de standard à importer.

'// Options defaut.
    Private Const CLS_VAR       As String = "CSordForm"             '// Name variable pour la classe 'CsordFormColumn'
    Private Const FUNC_NAME     As String = "=SordColumn()"         '// Nom de la function pour lancer le tri (OnClick CommandButton).
    Private Const FUNC_CLOSE    As String = "=CloseSordColumn()"    '// Nom de la function reset classe (OnClose form).
    
    Private Const PREF_CMB      As String = "cmb_"                  '// Préfixe pour le commandbutton.
    Private Const PREF_TXTBOX   As String = "txt_"                  '// Préfixe pour la textbox.

    Private m_oMsApp            As Access.Application

    '// Options définies dans le formulaire.
    Private m_oMsBase           As DAO.Database
    Private m_sBaseFullName     As String       '// Sauvegarde du chemin\nom de la base ouverte.
    Private m_bMsAppIsUp        As Boolean
    Private m_bMsBaseIsOpen     As Boolean

    Private m_sOptFormName      As String
    Private m_sOptFormSource    As String
    Private m_sFormFields()     As String       '// Stock les champs à utiliser dans le form.
    Private m_lFieldIndex       As Long         '// Index du tableau des champs.

    Private m_sOptPictureFolder As String
    Private m_sOptPictureAsc    As String
    Private m_sOptPictureDesc   As String
    
    Private m_cUtils            As New CUtilitaires     '// FSO
'//:::::::::::::::::::::::::::::::::: END VARIABLES ::::::::::::::::::::::::::::::::::::::
    
'//----------------------------------       INIT/TERM       ----------------------------------
Private Sub Class_Initialize()

    '// Défini les options par défaut.
    
    m_sOptPictureFolder = vbNullString
    m_sOptPictureAsc = vbNullString
    m_sOptPictureDesc = vbNullString

    m_lFieldIndex = -1      '// Indique tableau non initialiser.

End Sub

Private Sub Class_Terminate()

    On Error Resume Next

    If (MsBaseIsOpen And (Not m_oMsBase Is Nothing)) Then m_oMsBase.Close
    If (MsAppIsUp And (Not m_oMsApp Is Nothing)) Then m_oMsApp.Quit
    DoEvents

    Set m_oMsBase = Nothing
    Set m_oMsApp = Nothing
    Set m_cUtils = Nothing

End Sub
'//---------------------------------------------------------------------------------------

'//==================================       PROP        ==================================

Public Property Get objMsBase() As DAO.Database
    Set objMsBase = m_oMsApp.CurrentDb
End Property

Public Property Get MsAppIsUp() As Boolean
    MsAppIsUp = m_bMsAppIsUp
End Property
Public Property Let MsAppIsUp(bNewVal As Boolean)
    m_bMsAppIsUp = IIf(m_oMsApp Is Nothing, False, bNewVal)
End Property

Public Property Get MsBaseIsOpen() As Boolean
    MsBaseIsOpen = m_bMsBaseIsOpen
End Property
Public Property Let MsBaseIsOpen(bNewVal As Boolean)
    m_bMsBaseIsOpen = IIf((m_oMsApp.CurrentProject.AccessConnection Is Nothing), False, bNewVal)
End Property

Public Property Get GetBaseFullName() As String
    GetBaseFullName = m_sBaseFullName
End Property

Public Property Get OptFormName() As String
    OptFormName = m_sOptFormName
End Property
Public Property Let OptFormName(NewValue As String)
    m_sOptFormName = NewValue
End Property

Public Property Get OptFormSource() As String
    OptFormSource = m_sOptFormSource
End Property
Public Property Let OptFormSource(sNewValue As String)
    m_sOptFormSource = sNewValue
End Property

Public Property Get OptVarClasse() As String
    OptVarClasse = CLS_VAR
End Property

Public Property Get OptFunctionName() As String
    OptFunctionName = FUNC_NAME
End Property

Public Property Get GetFields() As String()
    '// Retourne le tableau des Fields sélectionnés.
    GetFields = m_sFormFields
End Property

'//     m_sFormFields()    '// Stock les champs à utiliser dans le from.
'//     m_iFieldIndex      '// Index(démarre à -1 donc 0) du tableau des champs.
Public Property Let AddField(sFieldName As String)
    If (Len(sFieldName) = 0) Then Exit Property

    If (m_lFieldIndex <= UBound(m_sFormFields)) Then
        m_lFieldIndex = m_lFieldIndex + 1
        ReDim Preserve m_sFormFields(0 To m_lFieldIndex) As String
    End If
    m_sFormFields(m_lFieldIndex) = sFieldName

End Property

Public Property Get OptPictureFolder() As String
    OptPictureFolder = m_sOptPictureFolder
End Property
Public Property Let OptPictureFolder(ByVal sNewValue As String)
    m_sOptPictureFolder = sNewValue
End Property

Public Property Get OptPictureAsc() As String
    OptPictureAsc = m_sOptPictureAsc
End Property
Public Property Let OptPictureAsc(sNewValue As String)
    m_sOptPictureAsc = sNewValue
End Property

Public Property Get OptPictureDesc() As String
    OptPictureDesc = m_sOptPictureDesc
End Property
Public Property Let OptPictureDesc(sNewValue As String)
    m_sOptPictureDesc = sNewValue
End Property

Public Property Get OptCmbPrefix() As String
    OptCmbPrefix = PREF_CMB
End Property

Public Property Get OptTextBoxPrefix() As String
    OptTextBoxPrefix = PREF_TXTBOX
End Property
'//====================================== END PROP =======================================

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function OpenMsApp() As Boolean
    On Error GoTo ERR_OpenMsApp

    If Me.MsAppIsUp Then
        MsgBox "Un objet Access.Application à déjà était créer.", vbInformation + vbOKOnly, "OpenMsApp"
        Exit Function
    End If

    '// Création de l'objet Access.Application.
    Set m_oMsApp = New Access.Application
    DoEvents
    m_oMsApp.Visible = False
    m_bMsAppIsUp = True
    OpenMsApp = True

SORTIE_OpenMsApp:
    Exit Function

ERR_OpenMsApp:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.CCreateFormContinu.OpenMsApp, ligne " & Erl & "."
    Resume SORTIE_OpenMsApp
End Function

' ----------------------------------------------------------------
' Procedure Nom:    OpenMsBase
' Sujet:            Ouvre la base.
' Procedure Kind:   Function
' Procedure Access: Public
' Références:       Ouvre la base.
'
'=== Paramètres ===
' sBaseFullName (String): Chemin et nom de la base à ouvrir.
'==================
'
' Return Type: Boolean TRUE si pas de problème.
'
' Author: Laurent
' Date:    30/04/2022 - 18:11
' DateMod: 07/05/2022 - 12:42
'
' !Use! : CheckFileExist, HasAutoexec, GetStartUpForm
' ----------------------------------------------------------------
Public Function OpenMsBase(sBaseFullName As String) As Boolean
    On Error GoTo ERR_OpenMsBase

    If Me.MsBaseIsOpen Then
        MsgBox "Un objet DAO.Database est déjà ouvert, veuillez fermer cette base svp.", vbInformation + vbOKOnly, "OpenMsBase"
        Exit Function
    End If

    Dim sObj    As String
    Dim bRep    As Boolean
    Dim bAutoEx As Boolean

    '// Voir si la base déjà ouverte (fichier .laccdb)...
    bRep = CheckFileExist(sBaseFullName, "laccdb")
    If bRep Then
        MsgBox "La base " & sBaseFullName & vbCrLf & "est déjà ouverte, opération annuler.", vbExclamation, "OpenMsBase"
        OpenMsBase = False
        Exit Function
    End If

    '// Ouverture de la base (sBaseName).
    m_oMsApp.OpenCurrentDatabase sBaseFullName, True
    m_oMsApp.Visible = False
    DoEvents

    Set m_oMsBase = m_oMsApp.CurrentDb
    '// Vérifie si bien initialisée.
    If (m_oMsBase Is Nothing) Then
        MsgBox "Impossible d'ouvrir la base " & sBaseFullName, vbCritical, "OpenMsBase"
        OpenMsBase = False
        Exit Function
    End If

    '// Vérifier si macro Autoexec...
    bAutoEx = HasAutoexec(m_oMsBase)
    If bAutoEx Then
        MsgBox "La base utilise une macro 'Autoexec', désactivez la avant d'utiliser cet outil." & vbCrLf & _
               "Opération annulée.", vbExclamation, "OpenMsBase"
        Me.MsBaseIsOpen = True
        Exit Function
    End If

    '// Vérifier si un form ouvert au démarrage...
    sObj = GetStartUpForm(m_oMsBase)
    If (sObj <> vbNullString) Then

        MsgBox "Le Formulaire '" & sObj & "' est ouvert automatiquement au lancement de la base," & vbCrLf & _
               "il doit être fermer pour continuer." & vbCrLf & vbCrLf & _
               "Note: il peux contenir du code qui vas fermer base ou l'application.", vbExclamation, "OpenMsBase"
        '// On ferme l'objet.
        m_oMsApp.DoCmd.Close acForm, sObj
    End If

    '// Vérification si la base est toujours ouverte.
    If (m_oMsApp.CurrentProject.AccessConnection Is Nothing) Then
        MsgBox "La base à était fermée par l'objet " & sObj, vbExclamation, "OpenMsBase"
        m_bMsBaseIsOpen = True      '// Evite message 'base déjà fermée".
        Exit Function
    End If

    Me.MsBaseIsOpen = True              '// Flag base ouverte.
    m_sBaseFullName = sBaseFullName     '// Stock le nom.
    ReDim m_sFormFields(0)              '// RaZ du tableau des champs.
    m_lFieldIndex = -1                  '// RaZ index tableau.

    OpenMsBase = True

SORTIE_OpenMsBase:
    Exit Function

ERR_OpenMsBase:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.CCreateFormContinu.OpenMsBase, ligne " & Erl & "."
    Resume SORTIE_OpenMsBase
End Function

Public Function CloseMsBase(Optional CloseMsApp As Boolean = False) As Boolean
    On Error GoTo ERR_CloseMsBase

    'TODO: m_oMsApp.Visible = True

'    If (Me.MsAppIsUp = False) Or (Me.MsBaseIsOpen = False) And (CloseMsApp = False) Then
'        MsgBox "La base est déjà fermée.", vbInformation + vbOKOnly, "OpenMsBase"
'        Exit Function
'    End If

    '// Ouverture de la base (sBaseName).
    If (Not m_oMsBase Is Nothing) Then
        If (Not m_oMsApp.CurrentProject.AccessConnection Is Nothing) Then m_oMsApp.CloseCurrentDatabase
    End If
    DoEvents

    m_bMsBaseIsOpen = False
    Set m_oMsBase = Nothing

    '// On ferme aussi l'application ?
    If CloseMsApp Then
        If (MsAppIsUp And (Not m_oMsApp Is Nothing)) Then m_oMsApp.Quit
        MsAppIsUp = False
        DoEvents
        Set m_oMsApp = Nothing
    End If

    CloseMsBase = True

SORTIE_CloseMsBase:
    Exit Function

ERR_CloseMsBase:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.CCreateFormContinu.CloseMsBase, ligne " & Erl & "."
    Resume SORTIE_CloseMsBase
End Function

Public Function LanceCreation() As Boolean
On Error GoTo ERR_LanceCreation

    Dim bRep        As Boolean  '// Retour Function.

    '// ------------------------------
    bRep = CreateModule(MOD_CLASS, vbext_ct_ClassModule)    '// Création du module de Classe...
    If (bRep = False) Then Exit Function
    bRep = CreateModule(MOD_STAND, vbext_ct_StdModule)      '// Création du module de Standard...
    If (bRep = False) Then Exit Function

    '// ------------------------------
    bRep = CreateFormColumn         '// Création du formulaire...

    If bRep Then
        m_oMsApp.DoCmd.Minimize
        LanceCreation = bRep
    End If

SORTIE_LanceCreation:
    Exit Function

ERR_LanceCreation:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  CSord.CCreateFormContinu.LanceCreation, ligne " & Erl & "."
    Resume SORTIE_LanceCreation
End Function
'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ END PUB. SUB/FUNC \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'// ######################## PRIVATE SUB/FUNC ###########################################

Private Function CreateModule(sModName As String, lModType As Long) As Boolean
    On Error GoTo ERR_CreateModule

    Dim oMd     As VBIDE.VBComponent
    Dim bRet    As Boolean
    Dim sMsg    As String   '// MsgBox.
    Dim lRep    As Long     '// MsgBox.

    '// Le module existe déjà ? ...
    If (ModuleExiste(sModName, m_oMsApp.VBE.ActiveVBProject)) Then
        sMsg = "Le module " & sModName & " existe déjà dans le projet source." & vbNewLine & vbNewLine & "Voulez-vous le remplacer ?"
        lRep = MsgBox(sMsg, vbDefaultButton2 + vbYesNo, "Remplacer le module")

        If (lRep = vbNo) Then                           '// Ne remplace pas le module existant dans la base.
            CreateModule = True
            Exit Function
        End If
        
        With m_oMsApp.VBE.ActiveVBProject.VBComponents  '// Supprime le module pour remplacement.
            .Remove .Item(sModName)
        End With
    End If

    '// Création du module.
    Set oMd = m_oMsApp.VBE.ActiveVBProject.VBComponents.Add(lModType)

    '// Renome.
    m_oMsApp.DoCmd.Close acModule, oMd.Name, acSaveYes
    m_oMsApp.DoCmd.Rename sModName, acModule, oMd.Name

    '// Copie le module de classe...
    If (sModName = MOD_CLASS) Then
        sMsg = CopyModule(sModName, Application.VBE.ActiveVBProject, m_oMsApp.VBE.ActiveVBProject, True)
    
        If (sMsg <> vbNullString) Then
            MsgBox sMsg, vbCritical, "Erreur CopyModule"
        Else
            '// Sauvegarde.
            m_oMsApp.DoCmd.Close acModule, sModName, acSaveYes
            DoEvents
    
            CreateModule = True
        End If
    End If

    '// Duplique le module standard en modifiant le code...
    If (sModName = MOD_STAND) Then
        bRet = ModifCodeModule(MOD_STAND, oMd)
        If bRet Then
            '// Sauvegarde.
            m_oMsApp.DoCmd.Close acModule, sModName, acSaveYes
            DoEvents
            CreateModule = bRet
        End If
    End If
    
SORTIE_CreateModules:
    Set oMd = Nothing
    Exit Function

ERR_CreateModule:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  CSord.CCreateFormContinu.CreateModule, ligne " & Erl & "."
    Resume SORTIE_CreateModules
End Function

' ----------------------------------------------------------------
' Procedure Nom:    CreateFormColumn
' Sujet:            Création du formulaire en continu pour la classe 'CsordFormColumn'
' Procedure Kind:   Function
' Procedure Access: Private
'
'=== Paramètres ===
'==================
'
' Return Type: Boolean True si pas d'erreur.
'
' Author:  Laurent
' Date:    03/05/2022 - 15:41
' DateMod: 23/04/2022 - 20:01
'
' !Use! : RemplaceCode
'
' NOTE: cmb 1 - Left 566 - Top 170 - Height 405 - Width 1125
'       twips 567 twips par centimètre / header 1700 twips
' ----------------------------------------------------------------
Private Function CreateFormColumn() As Boolean
    On Error GoTo ERR_CreateFormColumn

    Dim oObj        As Object
    Dim oForm       As Access.Form
    Dim oModule     As Access.Module
    Dim oCtrCmb     As Access.CommandButton
    Dim oCtrTxtB    As Access.TextBox
    Dim oDb         As DAO.Database
    Dim oRst        As DAO.Recordset
    Dim oRstField   As DAO.Field

    Const CTR_WIDTH     As Long = 2835  '// Width  TxtB/Cmb
    Const CMB_HEIGHT    As Long = 405   '// Height  CommandButton
    Const TXB_HEIGHT    As Long = 315   '// Height  TextBox
    Const CMB_TOP       As Long = 1303  '// Top CommandButton (en bas de la section détail)

    Dim lLeft       As Long     '// Position gauche TxtB/Cmb 566

    Dim lIndex      As Long     '// Index pour le tableau des champs.

    Dim lRep    As Long     '// MsgBox.
    Dim sMsg    As String   '// MsgBox.
    Dim sTmp    As String

    '// Vérification si le formulaire existe déjà.
    For Each oObj In m_oMsApp.CurrentProject.AllForms

        If oObj.Name = Me.OptFormName Then
            sMsg = "Le formulaire '" & Me.OptFormName & "' Existe déjà dans le projet source." & vbNewLine & vbNewLine & "Voulez-vous le remplacer ?"
            Exit For
        End If
    Next

    '// si le formulaire existe, le supprimer ?
    If (sMsg <> vbNullString) Then
        lRep = MsgBox(sMsg, vbDefaultButton2 + vbYesNo, "Remplacer le formulaire")
        If (lRep = vbYes) Then
            m_oMsApp.DoCmd.DeleteObject acForm, Me.OptFormName
        Else
            MsgBox "Opération annulée", vbInformation, "Le formulaire existe déjà"
            Exit Function
        End If
    End If

'// ------------ Création Form, TextBox, CommandButton ----------
    
    Set oForm = m_oMsApp.CreateForm()

    '// Paramètrages.
    With oForm
        .OnClose = FUNC_CLOSE                       '// Reset var classe onClose.
        .RecordSource = Me.OptFormSource
        .DefaultView = 1                            '// Mode continu.
        .Width = 17000
        .ScrollBars = 2                             '// Verticale seulement.
        .Caption = Me.OptFormName
        m_oMsApp.DoCmd.RunCommand acCmdFormHdrFtr   '// Création de la section header/footer.
        .Section(acHeader).Height = 1708
        .Section(acDetail).Height = 315
        sTmp = .Name                                '// Stock temporairement le nom du form.
    End With

    m_oMsApp.DoCmd.Close acForm, sTmp, acSaveYes
    m_oMsApp.DoCmd.Rename Me.OptFormName, acForm, sTmp      '// Renome le form.
    m_oMsApp.DoCmd.OpenForm Me.OptFormName, acDesign, , , , acHidden

    Set oForm = m_oMsApp.Forms.Item(Me.OptFormName)

    '// Ouvre la source du form, pour définir le nom des commandButton et TextBox.
    Set oDb = m_oMsApp.CurrentDb
    Set oRst = oDb.OpenRecordset(Me.OptFormSource, dbOpenDynaset, dbReadOnly + dbForwardOnly)

    lLeft = 0   '// point de depart position left pour des boutons/Textebox.

    For Each oRstField In oRst.Fields
        '// Le champs fait partie de ceux sélectionnés ?
        If (oRstField.Name = m_sFormFields(lIndex)) Then
        
            '// Création des CommandButton (in Header).
            Set oCtrCmb = m_oMsApp.CreateControl(Me.OptFormName, acCommandButton, acHeader, , , lLeft, CMB_TOP, CTR_WIDTH, CMB_HEIGHT)
            oCtrCmb.Name = Me.OptCmbPrefix & oRstField.Name
            oCtrCmb.Caption = oRstField.Name
            oCtrCmb.OnClick = FUNC_NAME             '// Ajoute la Procédure event pour les CommandButton.
            oCtrCmb.TabStop = False

            '// Création des TexteBox (in Detail).
            Set oCtrTxtB = m_oMsApp.CreateControl(oForm.Name, acTextBox, acDetail, , , lLeft, 0, CTR_WIDTH, TXB_HEIGHT)
            oCtrTxtB.Name = Me.OptTextBoxPrefix & oRstField.Name
            oCtrTxtB.ControlSource = oRstField.Name

            lLeft = lLeft + CTR_WIDTH               '// Incrémente position gauche des txtb/cmb.
            
            '// Tous les champs sélectionnés ont était créer, on sort.
            lIndex = lIndex + 1
            If (lIndex > UBound(m_sFormFields)) Then Exit For

        End If
    Next

    oRst.Close
    oDb.Close
'// --------- FIN Création Form, TextBox, CommandButton ---------

    '// Enregistrement du form.
    m_oMsApp.DoCmd.Close acForm, Me.OptFormName, acSaveYes
    CreateFormColumn = True

SORTIE_CreateFormColumn:
    On Error Resume Next
    Set oForm = Nothing
    Set oModule = Nothing
    Set oCtrCmb = Nothing
    Set oCtrTxtB = Nothing
    Set oRst = Nothing
    Set oRstField = Nothing
    Set oDb = Nothing

    Exit Function

ERR_CreateFormColumn:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  CCreateFormContinu.MD_CreateForm.CreateFormColumn, ligne " & Erl & "."
    Resume SORTIE_CreateFormColumn
End Function

' ----------------------------------------------------------------
' Procedure Nom:    ModifCodeModule
' Sujet:            Modifie le code du module MD_SordFormColumn
'                   Remplace le Code par le dossier et les images choisi.
' Procedure Kind:   Function
' Procedure Access: Private
' Références:
'
'=== Paramètres ===
' ModuleName (String):      '// Module a dupliquer.
' VBCompDest (VBComponent): '// VBComp de destination.
'==================
'
' Return Type: Boolean      '// TRUE si tous ce passe bien.
'
' Author:  Laurent
' Date:    27/05/2022 - 18:34
' DateMod:
'
' ----------------------------------------------------------------
Private Function ModifCodeModule(ModuleName As String, VBCompDest As VBComponent) As Boolean
On Error GoTo ERR_ModifCodeModule

    Dim oFSO        As Object
    Dim VBCompSrc   As VBIDE.VBComponent
    Dim sPath       As String
    Dim sTmp        As String
    Dim sLineSrc    As String
    Dim sLineDest   As String

    If Trim$(ModuleName) = vbNullString Then Exit Function
'        CopyModule = "Valeur de ModuleName est Null."
'        Exit Function
'    End If

    If m_oMsApp.VBE.ActiveVBProject Is Nothing Then Exit Function
'        CopyModule = "VBIDE.VBProject destination non initialisé."
'        Exit Function
'    End If

    If m_oMsApp.VBE.ActiveVBProject.Protection = vbext_pp_locked Then Exit Function
'        CopyModule = "Le projet destination est vérouillé pour l'affichage."
'        Exit Function
'    End If

    Set VBCompSrc = Application.VBE.ActiveVBProject.VBComponents(ModuleName)
    sLineSrc = VBCompSrc.CodeModule.Lines(1, VBCompSrc.CodeModule.CountOfLines)
    
    Set oFSO = m_cUtils.GetoFSO
    sTmp = oFSO.GetParentFolderName(Me.OptPictureFolder)
    If (InStr(1, Me.GetBaseFullName, sTmp) > 0) Then
        sTmp = """\" & oFSO.GetBaseName(Me.OptPictureFolder) & "\"""
        sPath = "CurrentProject.Path & " & sTmp
    Else
        sPath = """ &  Me.OptPictureFolder & """
    End If

    sLineDest = Replace(sLineSrc, """PICFOLDER""", sPath): sTmp = sLineDest
    sLineDest = Replace(sTmp, "PICIMGASC", Me.OptPictureAsc): sTmp = sLineDest
    sLineDest = Replace(sTmp, "PICIMGDESC", Me.OptPictureDesc)
    
    Set VBCompDest = m_oMsApp.VBE.ActiveVBProject.VBComponents(ModuleName)
    With VBCompDest.CodeModule
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, sLineDest
    End With

    ModifCodeModule = True
    
SORTIE_ModifCodeModule:
    Set VBCompSrc = Nothing
    Exit Function

ERR_ModifCodeModule:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  CSord.CCreateFormColumn.ModifCodeModule, ligne " & Erl & "."
    Resume SORTIE_ModifCodeModule
End Function
'// ################################# END PRIV. SUB/FUNC #################################
