VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Purpose: This class stores the parameters for method, properties, events and
'declarations, the Enum Member for enumerations and the TypeMember for
'the UDT's.
'Remarks: There are some propertites valid only for a member type like
'the IsOptional, IsArray and IsByRef are valid only on Parameters.
'Value is valid only for Enumerations Member
'Author:    Marco Pipino
Option Explicit

Public Name As String           'The name of the submember
Public Value As String          'The value of the SubMember (Valid for consts)
Public Description As String    'Purpose of the SubMember
Public IsOptional As Boolean    'If the parameter is optional
Public IsArray As Boolean       'If the parameter is an array
Public IsByRef As Boolean       'If the parameter is passed byRef or byVal
Public TypeValue As String      'Type of the SubMember

Private Sub Class_Initialize()
    IsOptional = False
    IsArray = False
    IsByRef = True
    TypeValue = "Variant"
End Sub

'Purpose: Parse a single parameter
Public Sub ParseParameter(temp As String)
    If FirstLeftPart(temp, "Optional", False, True) Then
        IsOptional = True
    End If
    
    If FirstLeftPart(temp, "ByRef", False, True) Then
        IsByRef = True
    ElseIf FirstLeftPart(temp, "ByVal", False, True) Then
        IsByRef = False
    End If
    
    Name = LeftPart(temp, " As ", False, True)
    If Len(temp) > 0 Then
        TypeValue = temp
    Else
        TypeValue = "Variant"
    End If
    
    If Right(Name, 1) = ")" Then IsArray = True
End Sub
