Attribute VB_Name = "GlobalFunctions"
'Purpose: This module contains all public functions and sub used by all class<BR>
'   in this project.
'Example: This is a test for
'   an example.<BR><BR>
'Code:      dim a as string
'Code:      a = 0
'Code:      b =1
'
'Another example
'Code:'In the Example Code you can write comments!!!
'Code:      Public Function Prova() as String
'Code:          Dim h as long 'comment
'Code:      End Function
'
'
'End of examples.
Option Explicit

'Purpose: Determine if the parameter strTemp is an Object or a Value during
'   the creation of the html file.<BR>
'Remarks: It check the gTypeValues Collection that contains the standard value type of
'   Visual Basic and the Enums and the UDTs defined in the project.
'Paramter: strType the type of the Object or value
Public Function isAnObject(ByVal strType As String) As Boolean
    Dim temp As Variant
    strType = UCase(strType)
    isAnObject = True
    For Each temp In gTypeValues
        If strType = CStr(temp) Then isAnObject = False
    Next
End Function

'Purpose: Return tru if the member is static
'Parameter: Definition the definition elaborated that have or not the static keyword
Public Function IsStaticMemb(Definition As String) As Boolean
    IsStaticMemb = False
    If FirstLeftPart(Definition, "Static", False, True) Then IsStaticMemb = True
End Function

'Purpose: Read the file and storing the return the result text
Public Function ReadTextFile(FileName As String) As String
    Dim intFile As Integer
    Dim strTemp As String
    
    intFile = FreeFile
    ReadTextFile = ""
    Open FileName For Input As #intFile
    Do While Not EOF(intFile)
        Input #intFile, strTemp
        ReadTextFile = ReadTextFile & strTemp & vbCrLf
    Loop
    Close #intFile
End Function

'Purpose: Return the scope of the memb in string format
Public Function ScopeOfMemb(Definition As String, _
        Optional DefaultValue As vbext_Scope = vbext_Public) As vbext_Scope
    On Error GoTo err_ScopeOfMemb
    ScopeOfMemb = DefaultValue
    If FirstLeftPart(Definition, "Public", False, True) Then
        ScopeOfMemb = vbext_Public
    ElseIf FirstLeftPart(Definition, "Private", False, True) Then
        ScopeOfMemb = vbext_Private
    ElseIf FirstLeftPart(Definition, "Friend", False, True) Then
        ScopeOfMemb = vbext_Friend
    End If
    Exit Function
err_ScopeOfMemb:
    ScopeOfMemb = DefaultValue
End Function

'Purpose: Write the HTML File using the Scripting.FileSystemObject
'Parameter FileName     The name of the file to be create
'Parameter strFile      The HTML text
Public Sub WriteTextFile(FileName As String, strFile As String)
    Dim intFile As Integer
    
    intFile = FreeFile
    Open FileName For Output As #intFile
    Print #intFile, strFile
    Close #intFile
End Sub

'Purpose: Get the current line of code and check the nexts lines of comment without
'   tags.<BR> It Returns the last line of comment.
'Parameter: IsExample if Is an Example test for Code Comments.
Public Function NextComments(VBcode As CodeModule, _
            ByRef CurrLine As Integer, _
            Optional IsExample As Boolean = False) As String
            
    Dim strComment As String
    NextComments = ""
    strComment = VBcode.Lines(CurrLine + 1, 1)
    
    Do While FirstLeftPart(strComment, "'", True, True)
        If Not IsKeyTag(strComment) Then
            If IsExample And FirstLeftPart(CStr(strComment), gBLOCK_CODE, False, True) Then
                strComment = Replace(strComment, " ", "&nbsp;")
                LeftPart strComment, gBLOCK_CODE, False, True
                NextComments = NextComments & "<FONT face=monospace size=2>" & HTMLCodeLine(strComment) & "</FONT>" & vbCrLf
            Else
                NextComments = NextComments & " " & strComment & vbCrLf
            End If
            CurrLine = CurrLine + 1
            strComment = VBcode.Lines(CurrLine + 1, 1)
        Else
            Exit Do
        End If
    Loop
End Function

'Purpose: Return true if the first left part of the parameter
'   strComment is a recognized tag
Private Function IsKeyTag(ByVal strComment As String) As Boolean
    IsKeyTag = False
    If FirstLeftPart(strComment, gBLOCK_AUTHOR) Then
        IsKeyTag = True
    ElseIf FirstLeftPart(strComment, gBLOCK_DATE_CREATION) Then
        IsKeyTag = True
    ElseIf FirstLeftPart(strComment, gBLOCK_DATE_LAST_MOD) Then
        IsKeyTag = True
    ElseIf FirstLeftPart(strComment, gBLOCK_EXAMPLE) Then
        IsKeyTag = True
    ElseIf FirstLeftPart(strComment, gBLOCK_PARAMETER) Then
        IsKeyTag = True
    ElseIf FirstLeftPart(strComment, gBLOCK_PROJECT) Then
        IsKeyTag = True
    ElseIf FirstLeftPart(strComment, gBLOCK_REMARKS) Then
        IsKeyTag = True
    ElseIf FirstLeftPart(strComment, gBLOCK_VERSION) Then
        IsKeyTag = True
    ElseIf FirstLeftPart(strComment, gBLOCK_NO_COMMENT) Then
        IsKeyTag = True
    End If
End Function

'Purpose:Remove double blank from a string
Public Function RemoveDoubleBlank(myStr As String) As String
    myStr = Trim(myStr)
    Do While InStr(1, myStr, "  ")
        myStr = Replace(myStr, "  ", " ")
    Loop
    RemoveDoubleBlank = myStr
End Function

'Purpose: Returns true if the left part of myStr is equal to myChars
'Remarks: This function is used for parsing the declaration and the comment<BR>
'   The myStr parameter is passed byRef and then it is truncated if the left part
'   of myStr is myChars and DeleteMyChars is True.
'Parameter: DeleteMyChars Remove the myChars string from myStr
Public Function FirstLeftPart(ByRef myStr As String, _
    myChars As String, _
    Optional MatchCase As Boolean = False, _
    Optional DeleteMyChars As Boolean = False) As Boolean
    Dim tempMystr As String
    tempMystr = myStr
    
    tempMystr = Trim(tempMystr)
    If Not MatchCase Then
        tempMystr = UCase(tempMystr)
        myChars = UCase(myChars)
    End If
    FirstLeftPart = (InStr(1, tempMystr, myChars) = 1)
    If FirstLeftPart And DeleteMyChars Then LeftPart myStr, myChars, MatchCase, True
End Function

'Purpose: This function returns the first part of myStr at the left of
'   the myChar character.<BR> The myStr variable is truncated of the left part.
'Remarks: The parameter DeleteChar indicate if the myChar character will be deleted.<BR>
'   If myChar is not encountered return all the myStr parameter.
'Example:
'Code:
'Code:       myVar = "lngTemp = 6"
'Code:       myResult = LeftPart(myvar,"=",TRUE,FALSE,FALSE)
'Code:
'       myReult is "lngTemp"
Public Function LeftPart(ByRef myStr As String, _
    myChar As String, _
    Optional MatchCase As Boolean = False, _
    Optional DeleteChar As Boolean = True, _
    Optional NotInQuotes As Boolean = True) As String
    
    Dim i As Integer
    Dim inext As Integer
    Dim InQuotes As Boolean
    Dim tempMystr As String
    
    tempMystr = UCase(myStr)
    If Not MatchCase Then myChar = UCase(myChar)
    If Not NotInQuotes Then
        i = InStr(1, tempMystr, myChar)
        If i > 0 Then
            LeftPart = Trim(Left(myStr, i - 1))
            myStr = Trim(Right(myStr, Len(myStr) - i + 1 - IIf(DeleteChar, Len(myChar), 0)))
        Else
            LeftPart = myStr
            myStr = ""
        End If
    Else
        inext = 0
        InQuotes = True
        Do While InQuotes
            InQuotes = False
            inext = InStr(inext + 1, tempMystr, myChar)
            If inext = 0 Then Exit Do
            For i = 1 To inext
                If Mid(tempMystr, i, 1) = """" Then InQuotes = Not InQuotes
            Next
        Loop
        If inext > 0 Then
            LeftPart = Trim(Left(myStr, inext - 1))
            myStr = Trim(Right(myStr, Len(myStr) - inext + 1 - IIf(DeleteChar, Len(myChar), 0)))
        Else
            LeftPart = Trim(myStr)
            myStr = ""
        End If
    End If
End Function

'Purpose: Return the String 'version' of the scope of a member<BR>
'   It's used during the creaton of HTML files.
Public Function ScopeToString(vScope As vbext_Scope) As String
    Select Case vScope
    Case vbext_Public
        ScopeToString = "Public"
    Case vbext_Private
        ScopeToString = "Private"
    Case vbext_Friend
        ScopeToString = "Friend"
    End Select
End Function

'Purpose: This function used by the Events,Enum,UDTs and Declaration Classes
'   retrieve the first line of comment before the definition
'Remarks: It's important a blank line between a declaration af a member
'   and the next initial comment for another member
'Parameter: VBcode      The Code Module for the member
'Parameter: intLine     The Line number of the definition
Public Function GetFirstComment(VBcode As CodeModule, intLine As Integer) As Integer
    On Error GoTo GetFirstComment_Error
    GetFirstComment = intLine - 1
    Do While FirstLeftPart(VBcode.Lines(GetFirstComment - 1, 1), "'", True, False)
        GetFirstComment = GetFirstComment - 1
    Loop
    Exit Function
GetFirstComment_Error:

End Function

'Purpose: This complex function fix a bug of the VB IDE.<BR>
'   When there are some lines with the underscore at final of string
'   the next CodeLocation property are incorrect.<BR>
'   The if I want know the correct codeline i must count the underline at the
'   end of lines before ma declaration.
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

'Purpose: It's used only by cEnum, then maybe it will move in this class
Public Function cLngP(var As Variant) As Long
    On Error Resume Next
    cLngP = 0
    cLngP = CLng(var)
End Function

'Purpose: This function Parse a signle line of code and returns it colored
Public Function HTMLCodeLine(Line) As String
    
    Dim strTemp As String
    Dim strComment As String
    Dim strDelimiter As String
    Dim strLine As String
    
    strTemp = ""
    HTMLCodeLine = ""
    
    'read all string and replace spaces with &nbsp;
    strComment = Line
    strComment = Replace(strComment, " ", "&nbsp;")
    'remove comments
    strLine = LeftPart(strComment, "'", True, False, True)
    
    strTemp = ""
    Do While Len(strLine) > 0
    
        strDelimiter = GetDelimiter(strLine)
        
        strTemp = LeftPart(strLine, LCase(strDelimiter), False, True, True)
        strTemp = Replace(strTemp, "<", "&lt;")
        strTemp = Replace(strTemp, ">", "&gt;")
        If isKeyWord(strTemp) Then
            HTMLCodeLine = HTMLCodeLine & "<FONT color=Blue>" & strTemp & "</FONT>" & strDelimiter
        Else
            HTMLCodeLine = HTMLCodeLine & strTemp & strDelimiter
        End If
    Loop
    
    HTMLCodeLine = HTMLCodeLine & IIf(Len(strComment) > 0, "<font color=Green>" & strComment & "</font>", "") & "<br>" & vbCrLf
End Function

'Purpose: Get the next valid delimiter in the source code line
Private Function GetDelimiter(Line As String) As String
    Dim intCloseDel As Integer
    Dim intTemp As Integer
    
    intCloseDel = 256
    GetDelimiter = "&nbsp;"
    intTemp = InStr(1, Line, "&nbsp;")
    If intTemp > 0 And intTemp < intCloseDel Then
        intCloseDel = intTemp
        GetDelimiter = "&nbsp;"
    End If
    intTemp = InStr(1, Line, ",")
    If intTemp > 0 And intTemp < intCloseDel Then
        intCloseDel = intTemp
        GetDelimiter = ","
    End If
    intTemp = InStr(1, Line, ")")
    If intTemp > 0 And intTemp < intCloseDel Then
        intCloseDel = intTemp
        GetDelimiter = ")"
    End If
    intTemp = InStr(1, Line, "(")
    If intTemp > 0 And intTemp < intCloseDel Then
        intCloseDel = intTemp
        GetDelimiter = "("
    End If
    
End Function

'Purpose: If I have forget some keyword please add it in this list
Private Function isKeyWord(strWord As String) As Boolean
    isKeyWord = False
    Dim a As Date
    Select Case UCase(strWord)
    Case "AND", "ANY", "AS", "BOOLEAN", "BYREF", "BYTE", "BYVAL", _
            "CASE", "CONST", "CURRENCY", "DATE", "DECLARE", "DIM", "DO", _
            "DOUBLE", "EACH", "ELSE", "ELESEIF", "END", "ENUM", _
            "ERROR", "EVENT", "EXIT", "EXPLICIT", "FALSE", "FOR", "FRIEND", _
            "FUNCTION", "GET", "GOSUB", "GOTO", "IF", "IMPLEMENTS", _
            "IN", "INTEGER", "IS", "LET", "LIB", "LONG", "LOOP", "NEXT", _
            "NEW", "NOT", "OBJECT", "ON", "OPTION", "OPTIONAL", "OR", "PARAMARRAY", _
            "PRIVATE", "PROPERTY", "PUBLIC", "REDIM", "RESUME", "RETURN", "SELECT", _
            "SET", "STEP", "STOP", "STRING", "SUB", "THEN", "TIME", "TO", "TRUE", _
            "TYPE", "UNTIL", "VARIANT", "WEND", "WHILE", "WITH"
        '...... add other
        isKeyWord = True
    End Select
End Function


