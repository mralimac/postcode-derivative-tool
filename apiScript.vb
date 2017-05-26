'Public Sub XmlLookup(ByVal outputColumn As String, ByVal addressColumn As String, ByVal startRow As Integer, ByVal endRow As Integer) 'This is a line for a future GUI
Public Sub XmlLookup()

'~~~Declaring Variables~~~'
Dim outputColumn As String
Dim Row As Integer
Dim addressColumn As String
Dim areaColumn As String
Dim failureCount As Integer
Dim goodCount As Integer
Dim startRow As Integer
Dim endRow As Integer
Dim apiReply As String
Dim apiStat As Integer

'Setting Global Variables
'~~End Declaring Variables~~'
'~~~Assigning Values~~~'
    startRow = 1
    endRow = startRow + 1
    failureCount = 0
    goodCount = 0    
    addressColumn = ""
    areaColumn = ""
	outputColumn = ""
'~~End Assigning Values~~'
Application.ScreenUpdating = False 'Having this to false make it run faster, but it will unresponsive during operation
'~~Start Program~~'
    For Row = startRow To endRow    
	
		'Resetting Values        
        apiReply = ""
        apiStat = 0        
       
        
        scroller (Row) 'This scrolls the screen down to keep up with row
		
        If getExistingValue(outputColumn, Row) <> True Then 'Checks if the output column is occupied
		
			
			apiReply = sendXmlHttp(getRawAddress(areaColumn, Row) & "+" & getRawAddress(addressColumn, Row)) 'This sends a request to google and returns XML file        
			apiStat = getAPIStatus(apiReply) 'This checks the status of the API
			Select Case apiStat
			Case 1  'If everything is good				 
				Range(outputColumn & Row).Value = getStrippedXml(apiReply) 'Retrieves only the postcode from XML and edits the cell to the Postcode
				goodCount = goodCount + 1 'Add one to Success counter
			Case 2  'If over API Query Limit
				MsgBox ("Over API Query Limit") 'Outputs error message
				End 'End Program
			Case 3 'If no address was found
				
				failureCount = failureCount + 1 'Add one to Failure counter
			Case Else 'If Error is unknown
				MsgBox ("Unknown Error") 'Outputs error message
				MsgBox (apiReply)
				failureCount = failureCount + 1 'Add one to Failure counter
			End Select			
        End If
    Next Row
    MsgBox ("Complete" & vbNewLine & failureCount & " Failed" & vbNewLine & goodCount & " Passed") 'Outputs summary box and ends program
    ActiveWorkbook.Save 'Saves the workbook that is being edited
End Sub
'~~End Program~~'

Function getExistingValue(ByVal Column As String, ByVal Row As Integer) As Boolean
    Dim cell As String
    cell = Range(Column & Row).Value
    If Len(cell) > 5 Then
        getExistingValue = True: Exit Function
        Else
        getExistingValue = False: Exit Function
    End If
End Function


Function getRawAddress(ByVal Column As String, ByVal Row As Integer) As String
    Dim addressCellLocation As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    addressCellLocation = Column & Row
    getRawAddress = Range(addressCellLocation).Value: Exit Function
End Function


Function sendXmlHttp(ByVal addressInput As String) As String
    Dim xmlhttp As New MSXML2.XMLHTTP60
    Dim formattedAddress As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    formattedAddress = Replace(addressInput, " ", "+")
    googleUrl = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & formattedAddress & ",&sensor=false&key=AIzaSyDzFgeh60tJuw2AibtlWRXUtw28p9Mv7e8"
    xmlhttp.Open "Get", googleUrl, False
    xmlhttp.send
    
    
    
    sendXmlHttp = xmlhttp.responseText: Exit Function
End Function


Function getAPIStatus(ByVal InpStr As String) As String
    Dim openPos As Integer
    Dim closePos As Integer
    Dim midbit As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    openPos = InStr(InpStr, "<status>")
    closePos = InStr(InpStr, "</status>")
    midbit = Mid(InpStr, openPos + 8, closePos - openPos - 8)
    If midbit = "OK" Then
        getAPIStatus = 1: Exit Function
    ElseIf midbit = "OVER_QUERY_LIMIT" Then
        getAPIStatus = 2: Exit Function
    ElseIf midbit = "ZERO_RESULTS" Then
        getAPIStatus = 3: Exit Function
    Else
        getAPIStatus = 4: Exit Function
    End If
End Function


Function getStrippedXml(ByVal InpStr As String) As String
    Dim openPos As Integer
    Dim closePos As Integer
    Dim midbit As String
    Dim finalString As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    openPos = InStr(InpStr, "<formatted_address>")
    closePos = InStr(InpStr, "</formatted_address>")
    midbit = Mid(InpStr, openPos + 19, closePos - openPos - 19)
    finalString = postCode(midbit)
    getStrippedXml = finalString: Exit Function
End Function
Function addressCleaner(ByVal InStr As String) As String
        
End Function

Function postCode(ByVal InpStr As String) As String
    Dim w       As String
    Dim j       As Long
    Dim Ptrn1
    Dim Ptrn2   As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    x = Split(Replace(InpStr, ",", " "), " ")
     
    Ptrn1 = Array("[A-Z][0-9]", "[A-Z][0-9][0-9]", "[A-Z][A-Z][0-9]", "[A-Z][A-Z][0-9][0-9]", _
    "[A-Z][0-9][A-Z]", "[A-Z][A-Z][0-9][A-Z]")
     
    Ptrn2 = "[0-9]*" '"[0-9][A-Z][A-Z]"
     
    On Error Resume Next
    For i = 0 To UBound(x)
        w = x(i)
        For j = LBound(Ptrn1) To UBound(Ptrn1)
            If Len(w) Then
                If w Like Ptrn1(j) And x(i + 1) Like Ptrn2 Then
                    If Err.Number <> 0 Then
                        Err.Clear
                        If w Like Ptrn1(j) & Ptrn2 Then
                            postCode = w: Exit Function
                        End If
                    Else
                        postCode = w & Space(1) & x(i + 1)
                        Exit Function
                    End If
                ElseIf w Like Ptrn1(j) Then
                    postCode = w: Exit Function
                End If
            End If
        Next
    Next
End Function
