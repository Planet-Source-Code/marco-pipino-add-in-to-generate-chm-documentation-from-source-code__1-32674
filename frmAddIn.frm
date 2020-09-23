VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Doc Generator"
   ClientHeight    =   4845
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   9960
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdViewDoc 
      Caption         =   "View Doc"
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   2220
      Width           =   1335
   End
   Begin VB.TextBox txtDOSOutputs 
      Height          =   1455
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3300
      Width           =   4935
   End
   Begin VB.CheckBox chkSourceCode 
      Caption         =   "Source Code"
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CheckBox chkVarAsProperty 
      Caption         =   "Variables like properties"
      Height          =   195
      Left            =   2880
      TabIndex        =   10
      Top             =   2820
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   9480
      TabIndex        =   8
      Top             =   2880
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9360
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleziona il Compilatore"
      Filter          =   "Executables(*.exe)|*.exe|"
   End
   Begin VB.TextBox txtCompiler 
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Top             =   2880
      Width           =   4455
   End
   Begin MSComctlLib.ListView lstTags 
      Height          =   1815
      Left            =   60
      TabIndex        =   4
      Top             =   300
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tag Type"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Comment"
         Object.Width           =   8819
      EndProperty
   End
   Begin MSComctlLib.TreeView trvComps 
      Height          =   2235
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3942
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.CheckBox chkPublicOnly 
      Caption         =   "Public member only"
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   2220
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Generate"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   2220
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Help compiler"
      Height          =   195
      Left            =   4920
      TabIndex        =   9
      Top             =   2640
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Project Components"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2220
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comments Tag"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   1065
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose: This Form is the Interface with the user. The most part of this
'   code is generated automatically by the Visual Basic Add-In Wizard.<BR>
'   Other Code is added for visualization of the CHM file and for the visualization
'   of result of compilation.<BR>
'Author:    Marco Pipino
Option Explicit

Private Const REG_SZ As Long = 1
Private Const HKEY_CLASSES_ROOT As Long = &H80000000

'Purpose: This declaration is used for read the path of application
'   for read the CHM File.<Br>
'   This information is stored in the registry in the key<BR>
'   <B>HKEY_CLASSES_ROOT\.chm\shell\open\command</B>
Private Declare Function RegOpenKey Lib "advapi32.dll" _
  Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As _
  String, phkResult As Long) As Long

'Purpose: This declaration is used for read the path of application
'   for read the CHM File.<Br>
'   This information is stored in the registry in the key<BR>
'   <B>HKEY_CLASSES_ROOT\.chm\shell\open\command</B>
Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

'Purpose: This declaration is used for read the path of application
'   for read the CHM File.<Br>
'   This information is stored in the registry in the key<BR>
'   <B>HKEY_CLASSES_ROOT\.chm\shell\open\command</B>
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
   Alias "RegQueryValueExA" (ByVal hKey As Long, _
   ByVal lpValueName As String, ByVal lpReserved As Long, _
   lpType As Long, lpData As Any, lpcbData As Long) As Long


Public VBInstance As VBIDE.VBE          'The instance of the VB IDE
Public Connect As Connect               'The Connect object
Private objProject As cProject          'The cProject object
Private ItemClicked As ListItem         'User for managing the ListBox

Private WithEvents objDos As DOSOutputs 'DOSOutput object for viewing
Attribute objDos.VB_VarHelpID = -1
                                        'the optputs of the hhc compiler
                                        
'Purpose: Hide the form
Private Sub CancelButton_Click()
    Connect.Hide
End Sub

'Purpose: Control of the check for Technical documentation.<BR>
'   Remove comments in the code if you want remove mixed documentations.
Private Sub chkPublicOnly_Click()
'    If chkPublicOnly.Value = 1 Then
'        chkVarAsProperty.Value = 1
'        chkVarAsProperty.Enabled = False
'        chkSourceCode.Value = 0
'        chkSourceCode.Enabled = False
'    Else
'        chkVarAsProperty.Enabled = True
'        chkSourceCode.Enabled = True
'    End If
End Sub

'Purpose: Select the compiler
Private Sub cmdBrowse_Click()
    CommonDialog1.ShowOpen
    If Len(CommonDialog1.FileName) > 0 Then
        txtCompiler.Text = CommonDialog1.FileName
    End If
End Sub

'Purpose: After copilation we can view the CHM file. <BR>
'   We must read the <B>HKEY_CLASSES_ROOT\.chm\shell\open\command</B> registry key
'   for the path of the HH.exe viewer.
Private Sub cmdViewDoc_Click()
    On Error GoTo cmvViewDoc_Error
    Dim lngRet As Long
    Dim lngKey As Long
    Dim lngKeyType As Long
    Dim strBuffer As String
    Dim lngBufferSize As Long
    
    strBuffer = Space(256)
    'Open the registry key
    lngRet = RegOpenKey(HKEY_CLASSES_ROOT, ".chm\shell\open\command", lngKey)
    If lngRet <> 0 Then
        GoTo cmvViewDoc_Error
    End If
        
    'Get the key type value
    lngRet = RegQueryValueEx(lngKey, "", 0&, lngKeyType, ByVal strBuffer, lngBufferSize)
    If lngKeyType <> REG_SZ Then
        GoTo cmvViewDoc_Error
    End If
    
    'Get the value of the key i.e. the path of the HH.exe application
    lngRet = RegQueryValueEx(lngKey, "", 0&, REG_SZ, ByVal strBuffer, lngBufferSize)
    If lngRet <> 0 Then
        GoTo cmvViewDoc_Error
    End If
    
    'Close the key
    lngRet = RegCloseKey(lngKey)
    strBuffer = Left(strBuffer, lngBufferSize)
    
    'Launch the help file
    Shell Replace(strBuffer, "%1", gProjectFolder & "\" & gProjectName & ".chm"), vbNormalFocus
    Exit Sub
cmvViewDoc_Error:
    MsgBox ("Can't open the file")
End Sub

'Purpose: Load the form and get the setting from the registry
Private Sub Form_Load()
    cmdViewDoc.Enabled = False
    GetSettings
End Sub

'Purpose: Unload the form
Private Sub Form_Unload(Cancel As Integer)
    Set objProject = Nothing
End Sub

'Purpose: With double click we can change the Tags for comments
Private Sub lstTags_DblClick()
    On Error Resume Next
    Dim strTemp As String
    strTemp = ItemClicked.ListSubItems(1).Text
    strTemp = InputBox(ItemClicked.ListSubItems(2).Text, , strTemp)
    If Len(strTemp) > 0 Then
        ItemClicked.ListSubItems(1).Text = strTemp
    End If
End Sub

'Purpose: Select the item clicked
Private Sub lstTags_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set ItemClicked = Item
End Sub

'Purpose: Event of the DOSOutputs object. Fill the TextBox txtOutputs with
'   the response of the compiler.
Private Sub objDos_ReceiveOutputs(CommandOutputs As String)
    txtDOSOutputs.Text = txtDOSOutputs.Text & CommandOutputs
End Sub

'Purpose: Launch the generation of the documentation.<BR>
'   Create a new cProject object, save the current settings to the registry,
'   save the check values and then create the Tree of the project, then the HTML files
'   and the compile.<B>
'Remarks: See the cProject for more information
Private Sub OKButton_Click()
    Dim strDosCommand As String
    Dim objMod As VBComponent
    
    
    Set objProject = New cProject
    Set objProject.VBI = VBInstance
    objProject.GetSettings
    
    
    MousePointer = vbHourglass
    SaveSettings
    
    gPublicOnly = chkPublicOnly.Value
    'gInsConsts = chkCostants.Value
    gVarAsProperty = chkVarAsProperty.Value
    gSourceCode = chkSourceCode.Value
    
    txtDOSOutputs.Text = ""
    
    objProject.BuildTree Me
    objProject.BuildTypesValue
    objProject.CreateHTMLFiles
    
    Set objDos = New DOSOutputs
    txtDOSOutputs.Text = ""
    strDosCommand = Chr(34) & gHHCCompiler & Chr(34) & " " & Chr(34) & _
        gProjectFolder & "\" & App.Title & "\" & gProjectName & ".hhp" & Chr(34)
    If (objDos.ExecuteCommand(strDosCommand)) Then
        cmdViewDoc.Enabled = True
    Else
        cmdViewDoc.Enabled = False
    End If
    MousePointer = vbDefault
    Set objDos = Nothing
    Set objProject = Nothing
End Sub

'Purpose: Get the setting from the registry and fill the Controls on the form.
Private Sub GetSettings()
    Dim comp As VBComponent
    
    gBLOCK_PURPOSE = GetSetting(App.Title, "Blocks", "Purpose", "Purpose:")
    gBLOCK_PROJECT = GetSetting(App.Title, "Blocks", "Project", "Project:")
    gBLOCK_AUTHOR = GetSetting(App.Title, "Blocks", "Author", "Author:")
    gBLOCK_DATE_CREATION = GetSetting(App.Title, "Blocks", "Date_Creation", "Creation:")
    gBLOCK_DATE_LAST_MOD = GetSetting(App.Title, "Blocks", "Date_Last_Mod", "Modified:")
    gBLOCK_VERSION = GetSetting(App.Title, "Blocks", "Version", "Version:")
    gBLOCK_EXAMPLE = GetSetting(App.Title, "Blocks", "Example", "Example:")
    gBLOCK_CODE = GetSetting(App.Title, "Blocks", "Code", "Code:")
    gBLOCK_REMARKS = GetSetting(App.Title, "Blocks", "Remarks", "Remarks:")
    gBLOCK_NO_COMMENT = GetSetting(App.Title, "Blocks", "NoComment", "''")
    gBLOCK_PARAMETER = GetSetting(App.Title, "Blocks", "Parameter", "Parameter:")
    gHHCCompiler = GetSetting(App.Title, "Compiler", "Path", "C:\Program Files\HTML Help Workshop\hhc.exe")
    
    lstTags.ListItems.Add 1, , "PROJECT"
    lstTags.ListItems(1).ListSubItems.Add 1, , gBLOCK_PROJECT
    lstTags.ListItems(1).ListSubItems.Add 2, , "Begin and End of project section"
    lstTags.ListItems.Add 2, , "AUTHOR"
    lstTags.ListItems(2).ListSubItems.Add 1, , gBLOCK_AUTHOR
    lstTags.ListItems(2).ListSubItems.Add 2, , "Author Block"
    lstTags.ListItems.Add 3, , "DATE_CREATION"
    lstTags.ListItems(3).ListSubItems.Add 1, , gBLOCK_DATE_CREATION
    lstTags.ListItems(3).ListSubItems.Add 2, , "Date Creation Block"
    lstTags.ListItems.Add 4, , "DATE_LAST_MOD"
    lstTags.ListItems(4).ListSubItems.Add 1, , gBLOCK_DATE_LAST_MOD
    lstTags.ListItems(4).ListSubItems.Add 2, , "Date of last modify Block"
    lstTags.ListItems.Add 5, , "VERSION"
    lstTags.ListItems(5).ListSubItems.Add 1, , gBLOCK_VERSION
    lstTags.ListItems(5).ListSubItems.Add 2, , "Version of the project"
    lstTags.ListItems.Add 6, , "PURPOSE"
    lstTags.ListItems(6).ListSubItems.Add 1, , gBLOCK_PURPOSE
    lstTags.ListItems(6).ListSubItems.Add 2, , "Purpose of the block"
    lstTags.ListItems.Add 7, , "EXAMPLE"
    lstTags.ListItems(7).ListSubItems.Add 1, , gBLOCK_EXAMPLE
    lstTags.ListItems(7).ListSubItems.Add 2, , "Example Block"
    lstTags.ListItems.Add 8, , "REMARKS"
    lstTags.ListItems(8).ListSubItems.Add 1, , gBLOCK_REMARKS
    lstTags.ListItems(8).ListSubItems.Add 2, , "Remarks for the Block"
    lstTags.ListItems.Add 9, , "PARAMETER"
    lstTags.ListItems(9).ListSubItems.Add 1, , gBLOCK_PARAMETER
    lstTags.ListItems(9).ListSubItems.Add 2, , "Parameter Block of a member"
    lstTags.ListItems.Add 10, , "CODE"
    lstTags.ListItems(10).ListSubItems.Add 1, , gBLOCK_CODE
    lstTags.ListItems(10).ListSubItems.Add 2, , "Code in a Example Block"
    lstTags.ListItems.Add 11, , "NO COMMENT"
    lstTags.ListItems(11).ListSubItems.Add 1, , gBLOCK_NO_COMMENT
    lstTags.ListItems(11).ListSubItems.Add 2, , "No Comments"
    
    txtCompiler.Text = gHHCCompiler
    
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If Len(comp.Name) > 0 Then
            trvComps.Nodes.Add Null, tvwChild, comp.Name, comp.Name
            If comp.Type = vbext_ct_ClassModule Then
                trvComps.Nodes(comp.Name).Checked = True
            End If
        End If
    Next
    
End Sub

'Purpose: Save the settings
Private Sub SaveSettings()
    gBLOCK_PROJECT = lstTags.ListItems(1).ListSubItems(1).Text
    gBLOCK_AUTHOR = lstTags.ListItems(2).ListSubItems(1).Text
    gBLOCK_DATE_CREATION = lstTags.ListItems(3).ListSubItems(1).Text
    gBLOCK_DATE_LAST_MOD = lstTags.ListItems(4).ListSubItems(1).Text
    gBLOCK_VERSION = lstTags.ListItems(5).ListSubItems(1).Text
    gBLOCK_PURPOSE = lstTags.ListItems(6).ListSubItems(1).Text
    gBLOCK_EXAMPLE = lstTags.ListItems(7).ListSubItems(1).Text
    gBLOCK_REMARKS = lstTags.ListItems(8).ListSubItems(1).Text
    gBLOCK_PARAMETER = lstTags.ListItems(9).ListSubItems(1).Text
    gBLOCK_CODE = lstTags.ListItems(10).ListSubItems(1).Text
    gBLOCK_NO_COMMENT = lstTags.ListItems(11).ListSubItems(1).Text
    gHHCCompiler = txtCompiler.Text
    
    SaveSetting App.Title, "Blocks", "Project", gBLOCK_PROJECT
    SaveSetting App.Title, "Blocks", "Author", gBLOCK_AUTHOR
    SaveSetting App.Title, "Blocks", "Date_Creation", gBLOCK_DATE_CREATION
    SaveSetting App.Title, "Blocks", "Date_Last_Mod", gBLOCK_DATE_LAST_MOD
    SaveSetting App.Title, "Blocks", "Version", gBLOCK_VERSION
    SaveSetting App.Title, "Blocks", "Purpose", gBLOCK_PURPOSE
    SaveSetting App.Title, "Blocks", "Example", gBLOCK_EXAMPLE
    SaveSetting App.Title, "Blocks", "Code", gBLOCK_CODE
    SaveSetting App.Title, "Blocks", "Remarks", gBLOCK_REMARKS
    SaveSetting App.Title, "Blocks", "Parameter", gBLOCK_PARAMETER
    SaveSetting App.Title, "Blocks", "NoComment", gBLOCK_NO_COMMENT
    
    SaveSetting App.Title, "Compiler", "Path", gHHCCompiler
End Sub

'Purpose: Put the cusor at the end of the textbox
Private Sub txtDOSOutputs_Change()
    txtDOSOutputs.SelStart = Len(txtDOSOutputs.Text)
End Sub
