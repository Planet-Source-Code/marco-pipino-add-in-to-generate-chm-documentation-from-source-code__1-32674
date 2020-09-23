Attribute VB_Name = "Module1"
Option Explicit

Public VBI As vbext_MemberType

Public Const HTML_ENUM_MEMBER_TEMPLATE = "<TR VALIGN=""top""><TD width=27%><I>###MembName###</I></TD><TD width=20%>###MembVal###</TD><TD width=57%>###MembDescr###</TD></TR>"
Public Const HTML_PARAMETER_TEMPLATE = "<TR VALIGN=""top""><TD width=27%><I>###Parameter###</I></TD><TD width=73%>###ParamDescr###</TD></TR>"

Public Enum vbext_myMemberTypeProperty
    vbext_myNone
    vbext_myValue
    vbext_myObject
End Enum

Sub Main()
'    Dim collParam As New cllParameters
'    Dim collEnumMember As New cllEnumMembers
'
'    Dim s As String, temp As String
'    Dim cpar As cParameter
'    Dim ceM As cEnumMember
'    Dim vb2 As New VBide.References
'
'
'    s = " byval pippo() as long, byref pluto as integer, optional antonio as ADODB.Connection , gianni as any  "
'    '[Optional] [ByVal | ByRef] [ParamArray] varname[( )] [As type]
'    Debug.Print s
'    collParam.ParseParam "Classe1", "Metodo1", s
'    For Each cpar In collParam
'        'cpar.Description = "Andiamo a vedere se <b>funziona</b> bene come abbiamo deciso di fare."
'
'        Debug.Print cpar.HTMLDescr
''        Debug.Print IIf(cpar.IsByRef, "ByRef", "ByVal")
''        Debug.Print IIf(cpar.IsOptional, "Optional", "Required")
''        Debug.Print IIf(cpar.IsArray, "IsArray", "Not is Array")
''        Debug.Print cpar.ParamName
''        Debug.Print cpar.ParamType
'    Next
'    Set ceM = New cEnumMember
'    ceM.ClassName = "Classe1"
'    ceM.EnumName = "Mio Tipo"
'    ceM.EnumMemberName = "mvt_Edit"
'    ceM.EnumMemberValue = 0
'    ceM.EnumMemberDescrption = ""
'    collEnumMember.Add ceM
'
'    Set ceM = New cEnumMember
'    ceM.ClassName = "Classe1"
'    ceM.EnumName = "Mio Tipo"
'    ceM.EnumMemberName = "mvt_Insert"
'    ceM.EnumMemberValue = 1
'    ceM.EnumMemberDescrption = "Inserimento"
'    collEnumMember.Add ceM
'
'    Set ceM = New cEnumMember
'    ceM.ClassName = "Classe1"
'    ceM.EnumName = "Mio Tipo"
'    ceM.EnumMemberName = "mvt_Delete"
'    ceM.EnumMemberValue = 2
'    ceM.EnumMemberDescrption = "Cancellazione"
'
'    collEnumMember.Add ceM
'
'    For Each ceM In collEnumMember
'        Debug.Print ceM.HTMLDescr
'    Next
End Sub


