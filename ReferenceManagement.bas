Attribute VB_Name = "ReferenceManagement"
Option Explicit
Option Compare Text
Option Base 0

Private Declare PtrSafe Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameW" _
  (ByVal lpszShortPath As LongPtr, ByVal lpszLongPath As LongPtr, ByVal cchBuffer As LongPtr) As LongPtr
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       24th March 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Sub EntryPointGetReferences()
Const strNoDescription As String = "#REF: No description"
Dim strDescription As String
Dim oRef As vbide.Reference
Dim strReferenceType As String
Dim strLongFilePath As String
Dim oTargetWorkbook As Excel.Workbook


Set oTargetWorkbook = ThisWorkbook

For Each oRef In oTargetWorkbook.VBProject.References
    Select Case oRef.Type
        Case vbide.vbext_RefKind.vbext_rk_Project
            strReferenceType = "Project"
        Case vbide.vbext_RefKind.vbext_rk_TypeLib
            strReferenceType = "TypeLib"
    End Select
     
    strLongFilePath = String(1024, vbNullChar)
    GetLongPathName StrPtr(oRef.FullPath), StrPtr(strLongFilePath), Len(strLongFilePath)
    strLongFilePath = Application.WorksheetFunction.Trim(strLongFilePath)
'   Debug.Print ("Original=" & oRef.FullPath & " New=" & strLongFilePath)
        
    If oRef.IsBroken Then
        strDescription = strNoDescription
    Else
        strDescription = oRef.Description
    End If
    
    Debug.Print ("IsBroken=" & oRef.IsBroken & " BuiltIn=" & oRef.BuiltIn & " Description=""" & strDescription & """ LongPath=" & strLongFilePath & " GUID=" & oRef.GUID & " MajorMinor=" & oRef.Major & "." & oRef.Minor & " Name=" & oRef.Name & " Type=" & strReferenceType)
    strLongFilePath = ""
    
Next
End Sub

