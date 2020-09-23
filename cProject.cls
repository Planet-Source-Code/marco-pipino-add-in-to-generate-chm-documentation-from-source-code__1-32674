VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2008
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Purpose: This is the core object of this application. After the parsing it contains
'all object of the project and the it creates the documentation.
'Author:    Marco Pipino
Option Explicit

Public Classes As OrderedColl           'Ordered collection of Classes modules
Attribute Classes.VB_VarHelpID = 2010
Public Modules As OrderedColl           'Ordered collection of Standard modules
Public Forms As OrderedColl             'Ordered collection of Form
Public UserControls As OrderedColl      'Ordered collection of User Controls
Public ActiveXDesigners As OrderedColl  'Ordered collection of ActiveX Designers
Public MDIForms As OrderedColl          'Ordered collection of MDIForms

Public Properties As OrderedColl        'Ordered collection of Properties
Attribute Properties.VB_VarHelpID = 2011
Public Variables As OrderedColl         'Ordered collection of Variables
Public Methods As OrderedColl           'Ordered collection of Methods
Attribute Methods.VB_VarHelpID = 2012
Public Events As OrderedColl            'Ordered collection of Events
Attribute Events.VB_VarHelpID = 2013
Public Consts As OrderedColl            'Ordered collection of Consts
Attribute Consts.VB_VarHelpID = 2014
Public Enums As OrderedColl             'Ordered collection of Enumerations
Attribute Enums.VB_VarHelpID = 2016
Public Types As OrderedColl             'Ordered collection of UDT's
Attribute Types.VB_VarHelpID = 2017
Public References As OrderedColl        'Ordered collection of References
Public Declarations As OrderedColl      'Ordered collection of Declares
Public Implementations As OrderedColl   'Ordered collection of Implements

Public VBI As VBIDE.VBE                 'The IDE of Visual Basic
Attribute VBI.VB_VarHelpID = 2018
    
Public Name As String                   'The name of the project
Public Description As String            'The purpose of the project
Public Author As String                 'The Author of the project
Public DateBuild As String              'The Date of creation
Public DateLastMod As String            'The date of last modification
Public Version As String                'The version
Public Example As String                'Examples of project

'Templates
Private HTML_TEMPLATE_PROJECT_OVERVIEW As String
Private HTML_TEMPLATE_REFERENCE As String
Private HTML_TEMPLATE_REFERENCES As String
Private HTML_TEMPLATE_MODULE_OVERVIEW As String
Private HTML_TEMPLATE_METHOD As String
Private HTML_TEMPLATE_PROPERTY As String
Private HTML_TEMPLATE_VARIABLE As String
Private HTML_TEMPLATE_EVENT As String
Private HTML_TEMPLATE_ENUM As String
Private HTML_TEMPLATE_TYPE As String
Private HTML_TEMPLATE_PARAMETER As String
Private HTML_TEMPLATE_DECLARATION As String
Private HTML_TEMPLATE_IMPLEMENTS As String
Private HTML_TEMPLATE_SOURCECODE As String
Private HTML_TEMPLATE_CONSTS As String
Private HTML_TEMPLATE_ENUM_MEMBER As String
Private HTML_TEMPLATE_TYPE_MEMBER As String
Private HTML_TEMPLATE_PROPERTY_VALUE As String
Private HTML_TEMPLATE_VARIABLE_VALUE As String
Private HTML_TEMPLATE_CONST As String

'Purpose: Create all the ordered collection of the project
Private Sub Class_Initialize()
    Set Classes = New OrderedColl
    Set Modules = New OrderedColl
    Set Forms = New OrderedColl
    Set UserControls = New OrderedColl
    Set ActiveXDesigners = New OrderedColl
    Set MDIForms = New OrderedColl
    
    Set Properties = New OrderedColl
    Set Methods = New OrderedColl
    Set Events = New OrderedColl
    Set Consts = New OrderedColl
    Set Properties = New OrderedColl
    Set Variables = New OrderedColl
    Set Enums = New OrderedColl
    Set Types = New OrderedColl
    Set References = New OrderedColl
    Set Declarations = New OrderedColl
    Set Implementations = New OrderedColl
End Sub

'Purpose: The core of the application.<BR>
'Create a tree with each object (members, submemers) for each module
Public Sub BuildTree(myForm As Object)
Attribute BuildTree.VB_HelpID = 2019
    Dim comp As VBComponent
    Dim objClass As cModule
    Dim objModule As cModule
    Dim objForm As cModule
    Dim objUserControl As cModule
    Dim objActiveXDesigner As cModule
    Dim objMDIForm As cModule
    
    'Set the default name for project
    Name = gProjectName
    
    For Each comp In VBI.ActiveVBProject.VBComponents
        Select Case comp.Type
        Case vbext_ct_ClassModule
            If myForm.trvComps.Nodes(comp.Name).Checked = True Then
                Set objClass = New cModule
                objClass.ComponentType = comp.Type
                objClass.Parse comp, Me
                Classes.Add objClass, objClass.Name
            End If
        Case vbext_ct_StdModule
            If myForm.trvComps.Nodes(comp.Name).Checked = True _
                    And Not gPublicOnly Then
                Set objModule = New cModule
                objModule.ComponentType = comp.Type
                objModule.Parse comp, Me
                Modules.Add objModule, objModule.Name
            End If
        Case vbext_ct_VBForm
            If myForm.trvComps.Nodes(comp.Name).Checked = True Then
                Set objForm = New cModule
                objForm.ComponentType = comp.Type
                objForm.Parse comp, Me
                Forms.Add objForm, objForm.Name
            End If
        Case vbext_ct_UserControl
            If myForm.trvComps.Nodes(comp.Name).Checked = True Then
                Set objUserControl = New cModule
                objUserControl.ComponentType = comp.Type
                objUserControl.Parse comp, Me
                UserControls.Add objUserControl, objUserControl.Name
            End If
        Case vbext_ct_ActiveXDesigner
            If myForm.trvComps.Nodes(comp.Name).Checked = True Then
                Set objActiveXDesigner = New cModule
                objActiveXDesigner.ComponentType = comp.Type
                objActiveXDesigner.Parse comp, Me
                ActiveXDesigners.Add objActiveXDesigner, objActiveXDesigner.Name
            End If
        Case vbext_ct_VBMDIForm
            If myForm.trvComps.Nodes(comp.Name).Checked = True Then
                Set objMDIForm = New cModule
                objMDIForm.ComponentType = comp.Type
                objMDIForm.Parse comp, Me
                MDIForms.Add objMDIForm, objMDIForm.Name
            End If
        End Select
    Next
End Sub

'Purpose: Get Settings for application from registry. Read the template files
Public Sub GetSettings()
Attribute GetSettings.VB_HelpID = 2020
    On Error GoTo ErrorHandler
    Dim strTemp As String
    
    gProjectFolder = Left(VBI.ActiveVBProject.FileName, InStrRev(VBI.ActiveVBProject.FileName, "\") - 1)
    
    MkDir gProjectFolder & "\" & App.Title
    
    gBLOCK_PURPOSE = GetSetting(App.Title, "Blocks", "Purpose", "Purpose:")
    gBLOCK_PROJECT = GetSetting(App.Title, "Blocks", "Project", "Project:")
    gBLOCK_AUTHOR = GetSetting(App.Title, "Blocks", "Author", "Author:")
    gBLOCK_DATE_CREATION = GetSetting(App.Title, "Blocks", "Date_Creation", "Creation:")
    gBLOCK_DATE_LAST_MOD = GetSetting(App.Title, "Blocks", "Date_Last_Mod", "Modified:")
    gBLOCK_VERSION = GetSetting(App.Title, "Blocks", "Version", "Version:")
    gBLOCK_EXAMPLE = GetSetting(App.Title, "Blocks", "Example", "Example:")
    gBLOCK_CODE = GetSetting(App.Title, "Blocks", "Code", "Code:")
    gBLOCK_REMARKS = GetSetting(App.Title, "Blocks", "Remarks", "Remarks:")
    gBLOCK_PARAMETER = GetSetting(App.Title, "Blocks", "Parameter", "Parameter:")
    gBLOCK_NO_COMMENT = GetSetting(App.Title, "BLocks", "NoComment", "''")
    
    gHHCCompiler = GetSetting(App.Title, "Compiler", "Path", "C:\Program Files\HTML Help Workshop\hhc.exe")
    
    
    HTML_TEMPLATE_MODULE_OVERVIEW = ReadTextFile(App.Path & "\TemplateModule.htm")
    HTML_TEMPLATE_METHOD = ReadTextFile(App.Path & "\TemplateMethod.htm")
    HTML_TEMPLATE_PROPERTY = ReadTextFile(App.Path & "\TemplateProperty.htm")
    HTML_TEMPLATE_VARIABLE = ReadTextFile(App.Path & "\TemplateVariable.htm")
    HTML_TEMPLATE_EVENT = ReadTextFile(App.Path & "\TemplateEvent.htm")
    HTML_TEMPLATE_ENUM = ReadTextFile(App.Path & "\TemplateEnum.htm")
    HTML_TEMPLATE_TYPE = ReadTextFile(App.Path & "\TemplateType.htm")
    HTML_TEMPLATE_CONSTS = ReadTextFile(App.Path & "\TemplateConsts.htm")
    HTML_TEMPLATE_SOURCECODE = ReadTextFile(App.Path & "\TemplateSourceCode.htm")
    HTML_TEMPLATE_PROJECT_OVERVIEW = ReadTextFile(App.Path & "\TemplateProject.htm")
    HTML_TEMPLATE_REFERENCES = ReadTextFile(App.Path & "\TemplateReferences.htm")
    HTML_TEMPLATE_DECLARATION = ReadTextFile(App.Path & "\TemplateDeclaration.htm")
    HTML_TEMPLATE_IMPLEMENTS = ReadTextFile(App.Path & "\TemplateImplements.htm")

    gHHPTemplate = ReadTextFile(App.Path & "\TemplateHHP.hhp")

    HTML_TEMPLATE_PARAMETER = "<TR VALIGN=top><TD width=27%><I>###Name###</I></TD><TD width=73%>###IsOptional### ###Type### ###Description###</TD></TR>"
    HTML_TEMPLATE_ENUM_MEMBER = "<TR vAlign=top><TD>###Name###</TD><TD align=middle>###Value###</TD><TD width=73%>###Description###</TD></TR>"
    HTML_TEMPLATE_TYPE_MEMBER = "<TR vAlign=top><TD>###Name###</TD><TD align=middle>###Type###</TD><TD width=73%>###Description###</TD></TR>"
    HTML_TEMPLATE_PROPERTY_VALUE = "<TR VALIGN=""top""><TD width=27%><I>###TypeValue###</I></TD><TD width=73%>###ReturnType###.</TD></TR>"
    HTML_TEMPLATE_VARIABLE_VALUE = "<TR VALIGN=""top""><TD width=27%><I>###TypeValue###</I></TD><TD width=73%>###ReturnType###.</TD></TR>"
    HTML_TEMPLATE_CONST = "<TR><TD>###Scope###</TD><TD><b>###Name###</b></TD><TD>###Type###</TD><TD><b>###Value###</b></TD><TD>###Description###</TD></TR>"

    HTML_TEMPLATE_REFERENCE = "<TR vAlign=top><TD width=30%><B>###FileName###</B></TD><TD width=20%><B>###Version###</B></TD><TD width=*><B>###Description###</B></TD></TR>"

    gProjectName = Right(VBI.ActiveVBProject.FileName, Len(VBI.ActiveVBProject.FileName) - InStrRev(VBI.ActiveVBProject.FileName, "\"))
    gProjectName = Left(gProjectName, Len(gProjectName) - 4)
    gProjectFolder = Left(VBI.ActiveVBProject.FileName, InStrRev(VBI.ActiveVBProject.FileName, "\") - 1)
    Exit Sub
ErrorHandler:
    'if the error is Path/File access error i.e. The directory exist
    If Err.Number = 75 Then
        Resume Next
    End If
End Sub

'Puprpose: Create the HTML References for the project
Private Sub CreateHTMLReferences()
    Dim temp As String
    Dim templibref As String
    Dim refFile As String
    
    Dim objref As Reference
    Dim ref As Reference
    For Each ref In VBI.ActiveVBProject.References
        temp = temp & HTML_TEMPLATE_REFERENCE
        temp = Replace(temp, "###Version###", CStr(ref.Major) & "." & CStr(ref.Minor))
        templibref = Right(ref.FullPath, Len(ref.FullPath) - InStrRev(ref.FullPath, "\"))
        templibref = Right(ref.FullPath, Len(ref.FullPath) - InStrRev(ref.FullPath, "\"))
        Do While InStr(1, templibref, ".") = 0
            templibref = Right(ref.FullPath, Len(ref.FullPath) - InStrRev(ref.FullPath, "\", Len(ref.FullPath) - Len(templibref) - 1))
        Loop
        temp = Replace(temp, "###FileName###", templibref)
        temp = Replace(temp, "###Description###", ref.Description)
    Next
    
    refFile = HTML_TEMPLATE_REFERENCES
    refFile = Replace(refFile, "###References###", temp)
    WriteTextFile gProjectFolder & "\" & App.Title & "\References.htm", refFile
End Sub

'Purpose: Create the project overview
Private Sub CreateHTMLProjectOverview()
    Dim strHTMLProject As String
    'Name = gProjectName
    
    strHTMLProject = HTML_TEMPLATE_PROJECT_OVERVIEW
    strHTMLProject = Replace(strHTMLProject, "###Name###", Name)
    strHTMLProject = Replace(strHTMLProject, "###Description###", Description)
    strHTMLProject = Replace(strHTMLProject, "###Version###", IIf(Len(Version) > 0, "Version: " & Version, ""))
    strHTMLProject = Replace(strHTMLProject, "###DateRelease###", IIf(Len(DateLastMod) > 0, "Release: " & DateLastMod, ""))
    strHTMLProject = Replace(strHTMLProject, "###Author###", IIf(Len(Author) > 0, "Author: " & Author, ""))
    If Len(Example) > 0 Then Example = "<B>Example</B><BR>" & Example
    strHTMLProject = Replace(strHTMLProject, "###Example###", Example)
    
    WriteTextFile gProjectFolder & "\" & App.Title & "\" & "Intro.htm", strHTMLProject
End Sub


'Purpose: Create all HTML files
Public Sub CreateHTMLFiles()

    Dim objCurr As Object
    Dim objTOC As cToc
    Dim strPath As String
    Dim objMemb As cMember
    
    strPath = gProjectFolder & "\" & App.Title & "\"
    'Create the TOC object
    Set objTOC = New cToc
    Set objTOC.Project = Me
    objTOC.CreateTOC
    Set objTOC = Nothing
    
    CreateHTMLProjectOverview
    CreateHTMLReferences
    CreateHHP
    
    For Each objCurr In Classes
        objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
            strPath & objCurr.Name & "_Overview.htm", gPublicOnly
        'If gInsConsts Then
            objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
                strPath & objCurr.Name & "_Consts.htm"
        'End If
        If gSourceCode Then
            Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
        End If
    Next
    
    For Each objCurr In Modules
        objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
            strPath & objCurr.Name & "_Overview.htm", gPublicOnly
        'If gInsConsts Then
            objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
                strPath & objCurr.Name & "_Consts.htm"
        'End If
        If gSourceCode Then
            Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
        End If
    Next
    
    For Each objCurr In Forms
        objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
            strPath & objCurr.Name & "_Overview.htm", gPublicOnly
        'If gInsConsts Then
            objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
                strPath & objCurr.Name & "_Consts.htm"
        'End If
        If gSourceCode Then
            Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
        End If
    Next
   
    For Each objCurr In UserControls
        objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
            strPath & objCurr.Name & "_Overview.htm", gPublicOnly
        'If gInsConsts Then
            objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
                strPath & objCurr.Name & "_Consts.htm"
        'End If
        If gSourceCode Then
            Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
        End If
    Next
  
    For Each objCurr In ActiveXDesigners
        objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
            strPath & objCurr.Name & "_Overview.htm", gPublicOnly
        'If gInsConsts Then
            objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
                strPath & objCurr.Name & "_Consts.htm"
        'End If
        If gSourceCode Then
            Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
        End If
    Next
    
    For Each objCurr In MDIForms
        objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
            strPath & objCurr.Name & "_Overview.htm", gPublicOnly
        'If gInsConsts Then
            objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
                strPath & objCurr.Name & "_Consts.htm"
        'End If
        If gSourceCode Then
            Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
        End If
    Next
    
    For Each objMemb In Methods
        objMemb.CreateHTMLMethod HTML_TEMPLATE_METHOD, HTML_TEMPLATE_PARAMETER, _
            strPath & objMemb.RefModule & "_" & objMemb.Name & "_Method.htm", _
            gPublicOnly
    Next
    
    For Each objMemb In Properties
        objMemb.CreateHTMLProperty HTML_TEMPLATE_PROPERTY, HTML_TEMPLATE_PROPERTY_VALUE, _
            strPath & objMemb.RefModule & "_" & objMemb.Name & "_Property.htm", _
            gPublicOnly
    Next
    
    For Each objMemb In Events
        objMemb.CreateHTMLEvent HTML_TEMPLATE_EVENT, HTML_TEMPLATE_PARAMETER, _
        strPath & objMemb.RefModule & "_" & objMemb.Name & "_Event.htm"
    Next
    
    For Each objMemb In Variables
        objMemb.CreateHTMLVariable HTML_TEMPLATE_VARIABLE, HTML_TEMPLATE_VARIABLE_VALUE, _
            strPath & objMemb.RefModule & "_" & objMemb.Name & "_Property.htm", _
            gPublicOnly
    Next
    
    For Each objMemb In Enums
        objMemb.CreateHTMLEnum HTML_TEMPLATE_ENUM, HTML_TEMPLATE_ENUM_MEMBER, _
             strPath & objMemb.RefModule & "_" & objMemb.Name & "_Enum.htm", _
             gPublicOnly
    Next
    
    For Each objMemb In Types
        objMemb.CreateHTMLType HTML_TEMPLATE_TYPE, HTML_TEMPLATE_TYPE_MEMBER, _
            strPath & objMemb.RefModule & "_" & objMemb.Name & "_Type.htm", _
            gPublicOnly
    Next
    
    For Each objMemb In Declarations
        objMemb.CreateHTMLDeclaration HTML_TEMPLATE_DECLARATION, HTML_TEMPLATE_PARAMETER, _
            strPath & objMemb.RefModule & "_" & objMemb.Name & "_Declaration.htm", _
            gPublicOnly
    Next

    For Each objMemb In Implementations
        objMemb.CreateHTMLImplements HTML_TEMPLATE_IMPLEMENTS, _
            strPath & objMemb.RefModule & "_" & objMemb.Name & "_Implements.htm"
    Next

End Sub

'Purpose: Create the hhp file from the template. It insert all files HTML created
Private Sub CreateHHP()

    Dim strHHPTemplate As String
    Dim FileList As String
    Dim objCurr As Object
    Dim strPath As String
    
    FileList = ""
    strPath = gProjectFolder & "\" & App.Title & "\"
    strHHPTemplate = gHHPTemplate
    
    strHHPTemplate = Replace(strHHPTemplate, "###Title###", gProjectName)
    strHHPTemplate = Replace(strHHPTemplate, "###ProjectTitle###", strPath & gProjectName)
    strHHPTemplate = Replace(strHHPTemplate, "###CompiledFile###", gProjectFolder & "\" & gProjectName)
    strHHPTemplate = Replace(strHHPTemplate, "###IntroFile###", gProjectFolder & "\" & App.Title & "\Intro.htm")
    
    For Each objCurr In Methods
        FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Method.htm" & vbCrLf
    Next

    For Each objCurr In Properties
        FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Property.htm" & vbCrLf
    Next

    For Each objCurr In Events
        FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Event.htm" & vbCrLf
    Next
    
    For Each objCurr In Variables
        FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Property.htm" & vbCrLf
    Next
    
    For Each objCurr In Enums
        FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Enum.htm" & vbCrLf
    Next
    
    For Each objCurr In Types
        FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Type.htm" & vbCrLf
    Next
    
    For Each objCurr In Classes
        FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
        If Consts.Count(objCurr.Name) > 0 Then
            FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
        End If
        If gSourceCode Then
            FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
        End If
    Next
    
    For Each objCurr In Modules
        FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
        If Consts.Count(objCurr.Name) > 0 Then
            FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
        End If
        If Not gSourceCode Then
            FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
        End If
    Next
    
    For Each objCurr In Forms
        FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
        If Consts.Count(objCurr.Name) > 0 Then
            FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
        End If
        If Not gSourceCode Then
            FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
        End If
    Next
    
    For Each objCurr In UserControls
        FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
        If Consts.Count(objCurr.Name) > 0 Then
            FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
        End If
        If Not gSourceCode Then
            FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
        End If
    Next
    
    For Each objCurr In ActiveXDesigners
        FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
        If Consts.Count(objCurr.Name) > 0 Then
            FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
        End If
        If Not gSourceCode Then
            FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
        End If
    Next
    
    For Each objCurr In MDIForms
        FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
        If Consts.Count(objCurr.Name) > 0 Then
            FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
        End If
        If Not gSourceCode Then
            FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
        End If
    Next
    
    strHHPTemplate = Replace(strHHPTemplate, "###FileList###", FileList)
    WriteTextFile strPath & gProjectName & ".hhp", strHHPTemplate
End Sub


'Purpose: Build a collection of Type Files user for recognize if the value of a
'member is as object or a not.
Public Function BuildTypesValue() As Boolean
Attribute BuildTypesValue.VB_HelpID = 2128
    Dim objMember As cMember
    On Error Resume Next
    Set gTypeValues = New Collection
    
    gTypeValues.Add "BOOLEAN", "BOOLEAN"
    gTypeValues.Add "BYTE", "BYTE"
    gTypeValues.Add "INTEGER", "INTEGER"
    gTypeValues.Add "LONG", "LONG"
    gTypeValues.Add "CURRENCY", "CURRENCY"
    gTypeValues.Add "SINGLE", "SINGLE"
    gTypeValues.Add "DOUBLE", "DOUBLE"
    gTypeValues.Add "DECIMAL", "DECIMAL"
    gTypeValues.Add "DATE", "DATE"
    gTypeValues.Add "STRING", "STRING"
    gTypeValues.Add "VARIANT", "VARIANT"
    gTypeValues.Add "ANY", "ANY"
    For Each objMember In Enums
        gTypeValues.Add UCase(objMember.Name), UCase(objMember.Name)
    Next
    For Each objMember In Types
        gTypeValues.Add UCase(objMember.Name), UCase(objMember.Name)
    Next
End Function

'Purpose: create the HTML for modules source code
Public Sub CreateHTMLCode(VBcode As CodeModule, ModuleName As String)
    
    Dim i As Integer
    Dim strCode As String
    Dim strCodeModule As String
    
    strCode = HTML_TEMPLATE_SOURCECODE
    strCodeModule = ""
    For i = 1 To VBcode.CountOfLines
        strCodeModule = strCodeModule & HTMLCodeLine(VBcode.Lines(i, 1)) & vbCrLf
    Next
    strCode = Replace(strCode, "###Name###", ModuleName)
    strCode = Replace(strCode, "###Code###", strCodeModule)
    WriteTextFile gProjectFolder & "\" & App.Title & "\" & ModuleName & "_Code.htm", strCode
End Sub

