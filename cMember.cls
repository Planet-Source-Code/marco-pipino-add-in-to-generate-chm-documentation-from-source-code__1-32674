VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Purpose: This class contain information about the each members of the module
'Author:    Marco Pipino

Option Explicit

Public RefModule As String              'The module of the member
Public MemberType As EnumMemberType     'Indicate if the member is a method, a property,
                                        'an event, etc..
Public Name As String                   'The name of the meber

Public Description As String            'Purpose of the member
Public Remarks As String                'Remorks for the member
Public Example As String                'One or more examples for the member

Public Scope As vbext_Scope             'The scope of the member (i.e. public,
                                        'private or friend)
Public IsStatic As Boolean              'Indicate if the member is static
Public TypeValue As String              'The type of member

Public Value As String                  'The value of member valid only for Consts
Public Library As String                'The library of the declare
Public Alias As String                  'The Alias of the declare
Public SubMembers As Collection         'This is a <B>NOT</B> ordered collection
                                        'beacause the order it's is important for
                                        'parameters, enum members and type members
Public PropertyType As EnumPropertyType 'Only for property, indicate if it's Read
                                        'Only, Write Only or Read and Write

'Purpose: Enumerate the type of member
Public Enum EnumMemberType
    memb_type_Method
    memb_type_Property
    memb_type_Const
    memb_type_Declare
    memb_type_Enum
    memb_type_Event
    memb_type_Implements
    memb_type_UDTs
    memb_type_Variable
End Enum

'Purpose: Enumerate the type of property
Public Enum EnumPropertyType
    Prop_Type_ReadOnly
    Prop_Type_ReadWrite
    Prop_Type_WriteOnly
End Enum

'Purpose: Class initialization, create only the SubMembers Colletion <B>
'not</B> ordered.
Private Sub Class_Initialize()
    Set SubMembers = New Collection
End Sub

'Purpose: Destroy the SubMembers collection.
Private Sub Class_Terminate()
    Set SubMembers = Nothing
End Sub

'Purpose: Parse the method and declare definitions and the its comments i.e.<BR><BR>
'for Declarations Syntax 1:<BR>
'[Public | Private] Declare Sub name Lib "libname" [Alias "aliasname"] [([arglist])]<BR>
'Syntax 2:<BR>
'[Public | Private] Declare Function name Lib "libname" [Alias "aliasname"] [([arglist])] [As type]<BR><BR>
'for Functions:<BR>
'[Public | Private | Friend] [Static] Function name [(arglist)] [As type]<BR><BR>
'for Subs:<BR>
'[Private | Public | Friend] [Static] Sub name [(arglist)]
Public Sub ParseMethod(VBcode As CodeModule, VBMemb As Member)
    On Error Resume Next
    Dim MethodDef As String

    Dim intParentesi As Integer
    Dim i As Integer
    Dim strTemp As String
    Dim strComment As String
    Dim intCodeLocation As Integer
    
    MethodDef = GetDef(VBcode, VBMemb, intCodeLocation, strComment)
    
    ScopeOfMemb MethodDef
    Scope = VBMemb.Scope
    
    IsStatic = IsStaticMemb(MethodDef)
    If FirstLeftPart(MethodDef, "Declare", False, True) Then
        MemberType = memb_type_Declare
    Else
        MemberType = memb_type_Method
    End If
    
    If FirstLeftPart(MethodDef, "Sub", False, True) Then
        TypeValue = "None"
    ElseIf FirstLeftPart(MethodDef, "Function", False, True) Then
        TypeValue = "Variant"
    End If
    
    If MemberType = memb_type_Declare Then
        'if it's a declaration
        strTemp = LeftPart(MethodDef, "(", True, True)
        Name = LeftPart(strTemp, "Lib", False, True, True)
        Library = LeftPart(strTemp, "Alias", False, True, True)
        
        Library = Left(Library, Len(Library) - 1)
        Library = Right(Library, Len(Library) - 1)
        Alias = strTemp
        Alias = Left(Alias, Len(Alias) - 1)
        Alias = Right(Alias, Len(Alias) - 1)
        
    Else
        'It's a method
        Name = LeftPart(MethodDef, "(", True, True)
    End If
    
    intParentesi = 0
    strTemp = ""
    For i = 1 To Len(MethodDef)
        If intParentesi = 0 And Mid(MethodDef, i, 1) = ")" Then
            Exit For
        ElseIf Mid(MethodDef, i, 1) = ")" Then
            intParentesi = intParentesi - 1
        ElseIf Mid(MethodDef, i, 1) = "(" Then
            intParentesi = intParentesi + 1
        Else
            strTemp = strTemp & Mid(MethodDef, i, 1)
        End If
    Next
    'Generate Parameter
    SetParameter (Trim(strTemp))
    
    MethodDef = Trim(Right(MethodDef, Len(MethodDef) - Len(strTemp) - 1))
    If UCase(TypeValue) = "VARIANT" And _
        FirstLeftPart(MethodDef, "AS", False, True) Then
            TypeValue = Trim(MethodDef)
    End If
    
    ParseDescriptions VBcode, GetFirstComment(VBcode, intCodeLocation)
    
    Description = Description & " " & strComment
End Sub

'Purpose: Parse descriptions for all members
Private Sub ParseDescriptions(VBcode As CodeModule, CodeLocation As Integer)
    On Error Resume Next
    Dim strTemp As String
    Dim tempParamDesc As String
    Dim tempParamName As String
    strTemp = VBcode.Lines(CodeLocation, 1)
    Do While FirstLeftPart(strTemp, "'", True, True)
        If FirstLeftPart(strTemp, gBLOCK_REMARKS, False, True) Then
            Remarks = Remarks & " " & strTemp & " " & vbCrLf & NextComments(VBcode, CodeLocation)
        ElseIf FirstLeftPart(strTemp, gBLOCK_NO_COMMENT, False, True) Then
            'There are no commands for the no comments tag
        ElseIf FirstLeftPart(strTemp, gBLOCK_EXAMPLE, False, True) Then
            Example = Example & " " & strTemp & vbCrLf & NextComments(VBcode, CodeLocation, True)
        ElseIf FirstLeftPart(strTemp, gBLOCK_PARAMETER, False, True) Then
            tempParamDesc = strTemp & " " & vbCrLf & NextComments(VBcode, CodeLocation)
            tempParamName = LeftPart(tempParamDesc, " ", True, True, True)
            SubMembers(tempParamName).Description = tempParamDesc
        ElseIf FirstLeftPart(strTemp, gBLOCK_PURPOSE, False, True) Then
            Description = Description & " " & strTemp & " " & vbCrLf & NextComments(VBcode, CodeLocation)
        End If
        CodeLocation = CodeLocation + 1
        strTemp = VBcode.Lines(CodeLocation, 1)
    Loop
End Sub

'Purpose: Parse a string of type <BR>
'[Optional] [ByVal | ByRef] [ParamArray] varname[( )] [As type]
'Remarks: The ParamArrya keyword generate error!!!!!
Private Sub SetParameter(argList As String)
    Dim temp As String
    Dim cpar As cSubMember
    Do While Len(argList) > 0
        temp = LeftPart(argList, ",", True, True, True)
        Set cpar = New cSubMember
        cpar.ParseParameter (temp)
        SubMembers.Add cpar, cpar.Name
    Loop
End Sub


'Purpose: This function used by the Events,Enum,UDTs and Declaration Classes
'   retrieve the first line of comment before the definition
'Remarks: It's important a blank line between a declaration af a member
'   and the next initial comment for another member
'Parameter: VBcode      The Code Module for the member
'Parameter: intLine     The Line number of the definition
Private Function GetFirstComment(VBcode As CodeModule, _
        intLine As Integer) As Integer
    On Error GoTo GetFirstComment_Error
    GetFirstComment = intLine - 1
    Do While FirstLeftPart(VBcode.Lines(GetFirstComment - 1, 1), "'", True, False)
        GetFirstComment = GetFirstComment - 1
    Loop
    Exit Function
GetFirstComment_Error:

End Function

'Purpose: this complex function count the numbeer of line that terminate with
'the Underscore char.<BR>
'The CodeLocation property of the Member object of VBIDE fails
'when encounters this type of lines
Private Function PrevUnderScore(VBModule As CodeModule, _
        CodeLocation As Integer) As Integer
    Dim i As Integer
    Dim intTemp As Integer
    Dim strTemp As String
    Dim intRowToScroll As Integer
    Dim intTempUnderScore As Integer
    
    PrevUnderScore = 0
    intTemp = 0
    intRowToScroll = CodeLocation
    Do
        intTempUnderScore = 0
        For i = intTemp + 1 To intTemp + intRowToScroll
            strTemp = Trim(VBModule.Lines(i, 1))
            If Right(strTemp, 1) = "_" Then
                intTempUnderScore = intTempUnderScore + 1
            End If
        Next
        PrevUnderScore = PrevUnderScore + intTempUnderScore
        intTemp = intTemp + intRowToScroll
        intRowToScroll = intTempUnderScore
    Loop While intTempUnderScore > 0
End Function

'Purpose: Parse the events declarations and its descriptions<BR>
'[<B>Public</B>] <B>Event</B> procedurename [<B>(</B>arglist<B>)</B>]
Public Sub ParseEvent(VBcode As CodeModule, VBMemb As Member)
    Dim EventDef As String

    Dim intParentesi As Integer
    Dim i As Integer
    Dim strTemp As String
    Dim strComment As String
    Dim intLastLine As Integer
    Dim intCodeLocation As Integer
    
    EventDef = GetDef(VBcode, VBMemb, intCodeLocation, strComment)
    ScopeOfMemb EventDef
    Call LeftPart(EventDef, "(", True, True, True)
    
    intParentesi = 0
    strTemp = ""
    For i = 1 To Len(EventDef)
        If intParentesi = 0 And Mid(EventDef, i, 1) = ")" Then
            Exit For
        ElseIf Mid(EventDef, i, 1) = ")" Then
            intParentesi = intParentesi - 1
        ElseIf Mid(EventDef, i, 1) = "(" Then
            intParentesi = intParentesi + 1
        End If
        strTemp = strTemp & Mid(EventDef, i, 1)
    Next
    SetParameter (Trim(strTemp))
    
    LeftPart EventDef, "'", True, True, True
    
    ParseDescriptions VBcode, GetFirstComment(VBcode, intCodeLocation)
    Description = Description & " " & EventDef
End Sub

'Purpose: Get the definition of a member even if on more code lines.
'Remarks: Returs the correct CodeLocation for the declaration of member in the
'byref parameter <I>CodeLocation</I>.<BR>
'It use the PrevUnderScore function for obtain the correct CodeLocation.<BR>
'The <I>strComment</I> byref parameter returns the comment after the definition.
Private Function GetDef(VBcode As CodeModule, _
    VBMemb As Member, CodeLocation As Integer, _
    Optional strCommentAfterDef As String) As String
    Dim temp As String
    Dim i As Integer
    Dim CodLoc As Integer
    
    CodLoc = VBMemb.CodeLocation + PrevUnderScore(VBcode, VBMemb.CodeLocation - 1)
    i = 0

    temp = RemoveDoubleBlank(VBcode.Lines(CodLoc, 1))
    Do While Right(temp, 1) = "_"
        i = i + 1
        temp = RemoveDoubleBlank(Left(temp, Len(temp) - 1)) _
            & " " & RemoveDoubleBlank(VBcode.Lines(CodLoc + i, 1))
    Loop
    
    GetDef = LeftPart(temp, "'", True, True, True)
    strCommentAfterDef = temp
    CodeLocation = CodLoc '+ i
End Function

'Purpose: Parse the variable code and its comments.<BR>

Public Sub ParseVariable(VBcode As CodeModule, VBMemb As Member)
    Dim VariableDef As String
    Dim i As Integer
    Dim strComment As String
    Dim intCodeLocation As Integer
    
    VariableDef = GetDef(VBcode, VBMemb, intCodeLocation, strComment)

    Call LeftPart(VariableDef, " AS ", False, True, True)
    If Len(VariableDef) > 0 Then
        TypeValue = VariableDef
    Else
        TypeValue = "Variant"
    End If
    ParseDescriptions VBcode, GetFirstComment(VBcode, intCodeLocation)
    Description = Description & " " & strComment & " " & vbCrLf & NextComments(VBcode, intCodeLocation)
End Sub

'Purpose: Parsing of the Const. When we define two or more consts on the same line this method
'FAILS.
Public Sub ParseConst(VBcode As CodeModule, VBMemb As Member)

    Dim ConstDef As String
    Dim intParentesi As Integer
    Dim i As Integer
    Dim strTemp As String
    Dim strComment As String
    Dim intCodeLocation As Integer
    
    ConstDef = GetDef(VBcode, VBMemb, intCodeLocation, strComment)
    
    ScopeOfMemb ConstDef
    
    FirstLeftPart ConstDef, "CONST", False, True
    FirstLeftPart ConstDef, VBMemb.Name, False, True
    
    If FirstLeftPart(ConstDef, "AS", False, True) Then
        TypeValue = LeftPart(ConstDef, "=", False, True, True)
    Else
        TypeValue = "Variant"
        Call LeftPart(ConstDef, "=", False, True, True)
    End If
    
    Value = LeftPart(ConstDef, "'", True, True)
    
    If TypeValue = "Variant" Then
        If IsNumeric(Value) Then
            TypeValue = "Long"
        ElseIf UCase(Value) = "TRUE" Or UCase(Value) = "FALSE" Then
            TypeValue = "Boolean"
        ElseIf IsDate(Value) Then
            TypeValue = "Date"
        Else
            TypeValue = "String"
        End If
    End If
    
    'Description = ConstDef
    ParseDescriptions VBcode, GetFirstComment(VBcode, intCodeLocation)
    Description = Description & " " & strComment
End Sub

'Purpose: Parse the property block. <B> The code location is on the Get
'block if exist, Let o Set else.</B>
Public Sub ParseProperty(VBcode As CodeModule, VBMemb As Member)
    'Viene definito se la proprietà è di sola lettura, sola scrittura
    'o lettura e scrittura.
    
    Dim strTemp As String
    Dim strDefProp As String
    Dim intCodeLocation As Integer
    Dim strComment As String
    
    'Recognize if the property is ReadOnly, WriteOnly o Read-Write
    If GetPropertyFirstLine(VBcode, VBMemb, vbext_pk_Get) > 0 Then
        If GetPropertyFirstLine(VBcode, VBMemb, vbext_pk_Let) > 0 Or _
            GetPropertyFirstLine(VBcode, VBMemb, vbext_pk_Set) > 0 Then
            PropertyType = Prop_Type_ReadWrite
        Else
            PropertyType = Prop_Type_ReadOnly
        End If
    Else
        PropertyType = Prop_Type_WriteOnly
    End If
        
    strDefProp = GetDef(VBcode, VBMemb, intCodeLocation, strComment)
    
    strTemp = LeftPart(strDefProp, " AS ", False, True, True)
    
    TypeValue = LeftPart(strDefProp, ")", True, True, True)
    If TypeValue = "" Then TypeValue = "Variant"
    
    ParseDescriptions VBcode, GetFirstComment(VBcode, intCodeLocation)
    Description = Description & " " & strComment
    
End Sub

'Purpose: Return the first code line of the block of property
'0 if the block does not exist
'Parameter: VBCode The code module
'Parameter: VBMemb The member property object
'Parameter: PropType The type of property
Private Function GetPropertyFirstLine(VBcode As CodeModule, _
        VBMemb As Member, PropType As vbext_ProcKind) As Long
    On Error Resume Next
    GetPropertyFirstLine = 0
    GetPropertyFirstLine = VBcode.ProcStartLine(VBMemb.Name, PropType)
End Function

'Purpose: Parse Implements sintax
'Parameter: VBCode The code module
Public Sub ParseImplements(VBcode As CodeModule, BeginLine As Integer)
    Dim i As Integer
    Dim strTemp As String
    
    
    strTemp = VBcode.Lines(BeginLine, 1)
    
    ScopeOfMemb strTemp
    
    FirstLeftPart strTemp, "Implements", False, True
    
    Name = LeftPart(strTemp, "'")
    ParseDescriptions VBcode, GetFirstComment(VBcode, BeginLine)
End Sub

'Purpose: Parse an enumeration
'Parameter: VBCode The code module
Public Sub ParseEnum(VBcode As CodeModule, _
    ByVal BeginLine As Integer, _
    EndLine As Integer)
    
    On Error Resume Next
    Dim i As Integer
    Dim strTemp As String
    Dim lngCurrMemberValue As Long
    Dim intInitComment As Integer
    Dim intEndComment As Integer
    
    strTemp = VBcode.Lines(BeginLine, 1)
    
    ScopeOfMemb strTemp
    FirstLeftPart strTemp, "Enum", False, True
    
    Name = LeftPart(strTemp, "'", True, True, True)
    ParseDescriptions VBcode, GetFirstComment(VBcode, BeginLine)
    
    lngCurrMemberValue = -1
    For i = BeginLine + 1 To EndLine
        strTemp = VBcode.Lines(i, 1)
        If Not FirstLeftPart(strTemp, "'", True, False) Then
            ParseEnumMember VBcode, strTemp, lngCurrMemberValue, i
        End If
    Next
End Sub

'Purpose: Parse Members of an enumeration
'Parameter: VBCode The code module
Private Function ParseEnumMember(VBcode As CodeModule, strLine As String, _
    currValue As Long, CurrLine As Integer) As Long
    
    Dim objEnumMember As cSubMember
    Dim defLine As String
    
    defLine = LeftPart(strLine, "'", True, True, True)
    Set objEnumMember = New cSubMember
    objEnumMember.Name = LeftPart(defLine, "=", False, True, True)
    If Len(defLine) > 0 Then
        objEnumMember.Value = defLine
        currValue = cLngP(defLine)
    Else
        currValue = currValue + 1
        objEnumMember.Value = currValue
    End If
    objEnumMember.Description = strLine & " " & vbCrLf & NextComments(VBcode, CurrLine)
    SubMembers.Add objEnumMember, objEnumMember.Name
End Function

'Purpose: Parse the UDT's block in a module
'Parameter: VBCode The code module
'Parameter: BeginLine The first line of the block
'Parameter: EndLine The last line of the block
Public Sub ParseType(VBcode As CodeModule, _
        BeginLine As Integer, _
        EndLine As Integer)
    
    Dim i As Integer
    Dim strTemp As String

    strTemp = VBcode.Lines(BeginLine, 1)
    
    ScopeOfMemb strTemp
    FirstLeftPart strTemp, "Type", False, True

    Name = LeftPart(strTemp, "'", True, True, True)
    ParseDescriptions VBcode, GetFirstComment(VBcode, BeginLine)
    
    For i = BeginLine + 1 To EndLine
        strTemp = VBcode.Lines(i, 1)
        If Not FirstLeftPart(strTemp, "'", True, False) Then
            ParseTypeMember VBcode, strTemp, i
        End If
    Next
End Sub

'Purpose: Parse the type submember
Public Sub ParseTypeMember(VBcode As CodeModule, _
        strLine As String, _
        CurrLine As Integer)
    Dim objTypeMember As cSubMember
    Dim strComment As String
    
    strComment = strLine
    strLine = LeftPart(strComment, "'", False, True)
    Set objTypeMember = New cSubMember
    objTypeMember.TypeValue = "Variant"
    objTypeMember.Name = LeftPart(strLine, " As ", False, True, True)
    If Len(strLine) > 0 Then objTypeMember.TypeValue = strLine
    objTypeMember.Description = strComment & " " & vbCrLf & NextComments(VBcode, CurrLine)
    SubMembers.Add objTypeMember, objTypeMember.Name
End Sub

'Purpose: Create HTML for Declarations
Public Sub CreateHTMLDeclaration(ByVal Template As String, _
    TemplateParamenter As String, _
    FileName As String, _
    PublicOnly As Boolean)
    
    Dim CurrParam As cSubMember
    Dim sintax As String
    Dim params As String
    
    If PublicOnly Then
        Exit Sub
    End If
    Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
    Template = Replace(Template, "###Name###", Name)
    Template = Replace(Template, "###Description###", Description)
    Template = Replace(Template, "###Library###", Library)
    
    If Len(Alias) > 0 Then Alias = "<p><B>Alias</B></p><p>" & Alias & "</P>"
    Template = Replace(Template, "###Alias###", Alias)
    
    Template = Replace(Template, "###ReturnType###", IIf(isAnObject(TypeValue), "A <b>" & TypeValue & "</b> " & "Object", IIf(TypeValue = "None", "<b>None</b>", "A <b>" & TypeValue & "</b> " & "Value")))
    Template = Replace(Template, "###Remarks###", IIf(Len(Remarks) > 0, "<p><b> Remarks</b></p>" & Remarks, ""))
    Template = Replace(Template, "###ObjectName###", RefModule)
    'Sintax
    sintax = "<b>" & Name & "</b> "
    If TypeValue <> "None" Then
        sintax = sintax & "<b>(</b>"
    End If
    params = ""
    sintax = sintax & "<i>"
    For Each CurrParam In SubMembers
        sintax = sintax & " " & IIf(CurrParam.IsOptional, "[", "") & CurrParam.Name & IIf(CurrParam.IsOptional, "]", "") & ","
        params = params & TemplateParamenter
        params = Replace(params, "###Name###", CurrParam.Name)
        params = Replace(params, "###Type###", IIf(isAnObject(CurrParam.TypeValue), "A <b>" & CurrParam.TypeValue & "</b> Object.", "A <b>" & CurrParam.TypeValue & "</b> Value."))
        params = Replace(params, "###Description###", CurrParam.Description)
        params = Replace(params, "###IsOptional###", IIf(CurrParam.IsOptional, "Optional.", "Required."))
    Next
    If Right(sintax, 1) = "," Then sintax = Left(sintax, Len(sintax) - 1)
    sintax = sintax & "</i>"
    If TypeValue <> "None" Then
        sintax = sintax & "<b>)</b>"
    End If
    Template = Replace(Template, "###Sintax###", sintax)
    Template = Replace(Template, "###Parameters###", params)
    
    If Len(Example) > 0 Then Example = "<B>Example</B><BR>" & Example
    Template = Replace(Template, "###Example###", Example)
    
    WriteTextFile FileName, Template
End Sub

'Purpose: create HTML For Enumerations
Public Sub CreateHTMLEnum(ByVal Template As String, TemplateMember As String, _
    FileName As String, PublicOnly As Boolean)
    Dim CurrMember As cSubMember
    Dim members As String
    
    If PublicOnly Then
        Template = Replace(Template, "###Scope###", "")
    Else
        Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
    End If
    Template = Replace(Template, "###Name###", Name)
    Template = Replace(Template, "###Description###", Description)
    Template = Replace(Template, "###Remarks###", IIf(Len(Remarks) > 0, "<p><b> Remarks</b></p>" & Remarks, ""))
    members = ""
    For Each CurrMember In SubMembers
        members = members & TemplateMember
        members = Replace(members, "###Name###", CurrMember.Name)
        members = Replace(members, "###Value###", CurrMember.Value)
        members = Replace(members, "###Description###", CurrMember.Description)
    Next
    Template = Replace(Template, "###Members###", members)
    If Len(Example) > 0 Then Example = "<B>Example</B><BR>" & Example
    Template = Replace(Template, "###Example###", Example)
    
    WriteTextFile FileName, Template
End Sub

'Purpose: Create HTML for Events
Public Sub CreateHTMLEvent(ByVal Template As String, _
        TemplateParameter As String, _
        FileName As String)
        
    Dim CurrParam As cSubMember
    Dim sintax As String
    Dim params As String
    If gPublicOnly Then
        Template = Replace(Template, "###Scope###", "")
    Else
        Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
    End If
    Template = Replace(Template, "###Name###", Name)
    Template = Replace(Template, "###Description###", Description)
    Template = Replace(Template, "###Remarks###", IIf(Len(Remarks) > 0, "<p><b> Remarks</b></p>" & Remarks, ""))
    Template = Replace(Template, "###ObjectName###", RefModule)
    'Sintax
    sintax = "<b>Private Sub </b><i>Object_</i><b>" & Name & " (</b> "
    params = ""
    sintax = sintax & "<i>"
    For Each CurrParam In SubMembers
        sintax = sintax & " " & IIf(CurrParam.IsOptional, "[", "") & CurrParam.Name & IIf(CurrParam.IsOptional, "]", "") & ","
        params = params & TemplateParameter
        params = Replace(params, "###Name###", CurrParam.Name)
        params = Replace(params, "###Type###", IIf(isAnObject(CurrParam.TypeValue), "A <b>" & CurrParam.TypeValue & "</b> Object.", "A <b>" & CurrParam.TypeValue & "</b> Value."))
        params = Replace(params, "###Description###", CurrParam.Description)
        params = Replace(params, "###IsOptional###", IIf(CurrParam.IsOptional, "Optional.", "Required."))
    Next
    If Right(sintax, 1) = "," Then sintax = Left(sintax, Len(sintax) - 1)
    sintax = sintax & "</i>"
    sintax = sintax & "<b> )</b>"
    Template = Replace(Template, "###Sintax###", sintax)
    Template = Replace(Template, "###Parameters###", params)
    If Len(Example) > 0 Then Example = "<B>Example</B><BR>" & Example
    Template = Replace(Template, "###Example###", Example)
    
    WriteTextFile FileName, Template
End Sub

'Purpose: Create HTML for Implements
Public Sub CreateHTMLImplements(ByVal Template As String, FileName As String)
    
    Dim CurrParam As cSubMember
    Dim sintax As String
    Dim params As String
    
    Template = Replace(Template, "###Name###", Name)
    Template = Replace(Template, "###Description###", Description)
    Template = Replace(Template, "###Remarks###", IIf(Len(Remarks) > 0, "<p><b> Remarks</b></p>" & Remarks, ""))
    Template = Replace(Template, "###ObjectName###", RefModule)
    If Len(Example) > 0 Then Example = "<B>Example</B><BR>" & Example
    Template = Replace(Template, "###Example###", Example)
    
    WriteTextFile FileName, Template
End Sub

'Purpose: Create HTML for Methods
Public Sub CreateHTMLMethod(ByVal Template As String, TemplateParamenter As String, FileName As String, PublicOnly As Boolean)
    Dim CurrParam As cSubMember
    Dim sintax As String
    Dim params As String
    
    If (PublicOnly And Scope = vbext_Public) Or Not PublicOnly Then
        If PublicOnly Then
            Template = Replace(Template, "###Scope###", "")
        Else
            Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
        End If
        Template = Replace(Template, "###Name###", Name)
        Template = Replace(Template, "###Description###", Description)
        Template = Replace(Template, "###ReturnType###", IIf(isAnObject(TypeValue), _
                        IIf(TypeValue = "None", "<b>None</b>", "A <b>" & TypeValue & "</b> " & "Object"), _
                        IIf(TypeValue = "None", "<b>None</b>", "A <b>" & TypeValue & "</b> " & "Value")))
        Template = Replace(Template, "###Remarks###", IIf(Len(Remarks) > 0, "<p><b> Remarks</b></p>" & Remarks, ""))
        Template = Replace(Template, "###ObjectName###", RefModule)
        'Sintax
        sintax = "<i>Object.</i><b>" & Name & "</b> "
        If TypeValue <> "None" Then
            sintax = sintax & "<b>(</b>"
        End If
        params = ""
        sintax = sintax & "<i>"
        For Each CurrParam In SubMembers
            sintax = sintax & " " & IIf(CurrParam.IsOptional, "[", "") & CurrParam.Name & IIf(CurrParam.IsOptional, "]", "") & ","
            params = params & TemplateParamenter
            params = Replace(params, "###Name###", CurrParam.Name)
            params = Replace(params, "###Type###", IIf(isAnObject(CurrParam.TypeValue), "A <b>" & CurrParam.TypeValue & "</b> Object.", "A <b>" & CurrParam.TypeValue & "</b> Value."))
            params = Replace(params, "###Description###", CurrParam.Description)
            params = Replace(params, "###IsOptional###", IIf(CurrParam.IsOptional, "Optional.", "Required."))
        Next
        If Right(sintax, 1) = "," Then sintax = Left(sintax, Len(sintax) - 1)
        sintax = sintax & "</i>"
        If TypeValue <> "None" Then
            sintax = sintax & "<b>)</b>"
        End If
        Template = Replace(Template, "###Sintax###", sintax)
        Template = Replace(Template, "###Parameters###", params)
        If Len(Example) > 0 Then Example = "<B>Example</B><BR>" & Example
        Template = Replace(Template, "###Example###", Example)
        
        WriteTextFile FileName, Template
    End If
End Sub

'Purpose: Create HTML for propery
Public Sub CreateHTMLProperty(ByVal Template As String, _
        TemplateValue As String, _
        FileName As String, _
        PublicOnly As Boolean)
        
    Dim sintax As String
    
    If (gPublicOnly And Scope = vbext_Public) Or Not gPublicOnly Then
        If PublicOnly Then
            Template = Replace(Template, "###Scope###", "")
        Else
            Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
        End If
        Template = Replace(Template, "###PropValue###", TemplateValue)
        Template = Replace(Template, "###Name###", Name)
        Template = Replace(Template, "###Description###", Description)
        Template = Replace(Template, "###ReturnType###", IIf(isAnObject(TypeValue), "A <b>" & TypeValue & "</b> " & "Object", "A <b>" & TypeValue & "</b> " & "Value") & ". " & ReadWrite)
        Template = Replace(Template, "###Remarks###", IIf(Len(Remarks) > 0, "<p><b> Remarks</b></p>" & Remarks, ""))
        Template = Replace(Template, "###ObjectName###", RefModule)
        Template = Replace(Template, "###TypeValue###", TypeValue)

        'Sintax
        sintax = ""
        If PropertyType = Prop_Type_ReadWrite And isAnObject(TypeValue) Then
            sintax = "<b>[Set] </b>"
        ElseIf PropertyType = Prop_Type_WriteOnly And isAnObject(TypeValue) Then
            sintax = "<b>Set </b>"
        End If
        
        sintax = sintax & "<i>Object.</i><b>" & Name & "</b> "
        
        If PropertyType = Prop_Type_WriteOnly Then
            sintax = sintax & " = <i>" & TypeValue & "</i>"
        ElseIf PropertyType = Prop_Type_ReadWrite Then
            sintax = sintax & "[ = <i>" & TypeValue & "</i> ]"
        End If
        Template = Replace(Template, "###Sintax###", sintax)
        
        If Len(Example) > 0 Then Example = "<B>Example</B><BR>" & Example
        Template = Replace(Template, "###Example###", Example)
        
        WriteTextFile FileName, Template
    End If
End Sub

'Purpose: Create HTML for UDT's
Public Sub CreateHTMLType(ByVal Template As String, _
        TemplateMember As String, _
        FileName As String, _
        PublicOnly As Boolean)
    
    Dim CurrMember As cSubMember
    Dim members As String
    
    If (PublicOnly And Scope = vbext_Public) Or Not PublicOnly Then
        If PublicOnly Then
            Template = Replace(Template, "###Scope###", "")
        Else
            Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
        End If
        Template = Replace(Template, "###Name###", Name)
        Template = Replace(Template, "###Description###", Description)
        Template = Replace(Template, "###Remarks###", IIf(Len(Remarks) > 0, "<p><b> Remarks</b></p>" & Remarks, ""))
        members = ""
        For Each CurrMember In SubMembers
            members = members & TemplateMember
            members = Replace(members, "###Name###", CurrMember.Name)
            members = Replace(members, "###Type###", CurrMember.TypeValue)
            members = Replace(members, "###Description###", CurrMember.Description)
        Next
        Template = Replace(Template, "###Members###", members)
        If Len(Example) > 0 Then Example = "<B>Example</B><BR>" & Example
        Template = Replace(Template, "###Example###", Example)
        
        WriteTextFile FileName, Template
    End If
End Sub

'Purpose: Create HTML for Variables<BR>
'IMPORTANT: Each public variabile is a read and write property
Public Sub CreateHTMLVariable(ByVal Template As String, _
        TemplateValue As String, _
        FileName As String, _
        PublicOnly As Boolean)
'###VarProp###
    Dim sintax As String
        
    If (PublicOnly And Scope = vbext_Public) Or Not gPublicOnly Then
        If PublicOnly Then
            Template = Replace(Template, "###Scope###", "")
        Else
            Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
        End If
        'If is a Developer documentation create a Variable page
        Template = Replace(Template, "###VarProp###", IIf(gVarAsProperty Or gPublicOnly, "Property", "Variable"))
            
        Template = Replace(Template, "###PropValue###", TemplateValue)
        Template = Replace(Template, "###Name###", Name)
        Template = Replace(Template, "###Description###", Description)
        Template = Replace(Template, "###ReturnType###", _
            IIf(isAnObject(TypeValue), "A <b>" & TypeValue & "</b> " & "Object", "A <b>" & TypeValue & "</b> " & "Value") & _
            IIf(gVarAsProperty Or gPublicOnly, ". The property is Read-Write", ""))
        Template = Replace(Template, "###Remarks###", IIf(Len(Remarks) > 0, "<p><b> Remarks</b></p>" & Remarks, ""))
        Template = Replace(Template, "###ObjectName###", RefModule)
        Template = Replace(Template, "###TypeValue###", TypeValue)
        
        'Sintax
        sintax = ""
        If isAnObject(TypeValue) Then
            sintax = "<b>[Set] </b>"
        End If
        
        sintax = sintax & "<i>Object.</i><b>" & Name & "</b> "
        sintax = sintax & "[ = <i>" & TypeValue & "</i> ]"
        
        Template = Replace(Template, "###Sintax###", sintax)
        Template = Replace(Template, "###Example###", "")
        
        WriteTextFile FileName, Template
    End If
End Sub

'Purpose: returns a string value that describes the type of property. <BR>
'Used for creation of the HTML file of a property.
Private Function ReadWrite() As String
    If PropertyType = Prop_Type_ReadOnly Then
        ReadWrite = "The Property is Read Only"
    ElseIf PropertyType = Prop_Type_ReadWrite Then
        ReadWrite = "The Property is Read-Write"
    Else
        ReadWrite = "The Property is Write Only"
    End If
End Function

