VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWordCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 32-Bit Windows functions needed
Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)

Private SafeArrayHeader(5) As Long      ' Header for the SafeArray Map
Private Map() As Integer                ' Maps onto the Text1's string
Public i As Long

Private Sub Class_Initialize()
    ' Set up the SafeArrayHeader
    SafeArrayHeader(0) = 1              ' Number of dimensions
    SafeArrayHeader(1) = 2              ' Bytes per element (integer = 2)
    
    SafeArrayHeader(4) = &H7FFFFFFF     ' Array size
  
    ' Force Map to use SafeArrayHeader as its own header
    RtlMoveMemory ByVal ArrPtr(Map), VarPtr(SafeArrayHeader(0)), 4
    
End Sub

' Experimental version developed due to a comment by Chaz.  This
' version only works properly if the words are space delimited,
' tabs, linefeeds and other legitemit "white spaces" aren't counted
Friend Function InStrWordCount(strSample As String) As Long
    Dim pos_curr As Long        ' Traveling pointer
    Dim pos_next As Long        ' Traveling pointer
    Dim strSpace As String      ' Cache a space
    
    ' Caching a space is faster than constructing one with each InStr call
    strSpace = " "
    
    ' Find the first space (character 32) in strSample
    pos_curr = InStr(1, strSample, strSpace, vbBinaryCompare)
    
    ' Check for a singleton word
    If pos_curr = 0 And Len(strSample) > 1 Then
        InStrWordCount = 1
        Exit Function
    End If
    
    ' If pos_curr is 2 or greater, position 1 is a word
    If pos_curr > 1 Then InStrWordCount = InStrWordCount + 1
    
    ' Enter the loop, the only way to break is for InStr to return
    ' 0 indicating that there are no more spaces in strSample
    Do
    
        ' Find the next space in strSample
        pos_next = InStr(pos_curr + 1, strSample, strSpace, vbBinaryCompare)
        
        ' If pos_next = 0 we've hit the end of strSample we need to bail
        If pos_next = 0 Then
            ' We need to make one more check before we're done
            ' If the difference between the length of strSample
            ' and pos_curr is greater than 1 then strSample ends
            ' in a word
            If Len(strSample) - pos_curr > 1 Then InStrWordCount = InStrWordCount + 1
            Exit Function
        End If
        
        ' If the difference between the two positions is 1 we didn't step
        ' over a word, otherwise we did
        If pos_next - pos_curr > 1 Then InStrWordCount = InStrWordCount + 1

        ' Move the pointer
        pos_curr = pos_next

    Loop
End Function

Friend Function FastWordCount(strSample As String) As Long
    Dim i As Long
    Dim lngLength As Long
    
    lngLength = Len(strSample)
    SafeArrayHeader(3) = StrPtr(strSample) ' Pointer to the 1st character
    
    For i = 0 To lngLength - 1
        If Map(i) > 32 Then
            FastWordCount = FastWordCount + 1
            Do
                i = i + 1
            Loop While Map(i) > 32
        End If
    Next i
End Function

Friend Function MidWordCount(strSample As String) As Long
    Dim i As Long
    Dim lngLength As Long
    
    lngLength = Len(strSample)
    
    For i = 1 To lngLength - 1
        If Asc(Mid$(strSample, i, 1)) > 32 Then
            MidWordCount = MidWordCount + 1
            Do
                i = i + 1
                If i >= lngLength Then Exit Function
            Loop While Asc(Mid$(strSample, i, 1)) > 32
        End If
    Next i
End Function

Friend Function SplitWordCount(strSample As String) As Long
    Dim i As Long
    Dim tmp() As String
    
    tmp = Split(strSample, " ")
    
    ' You cannot just use the ubound extra spaced are counted as words!
    For i = LBound(tmp) To UBound(tmp)
        If Len(tmp(i)) > 0 Then SplitWordCount = SplitWordCount + 1
    Next i
End Function

Private Sub Class_Terminate()
    ' Clean up our mess, do not prevent this from executing!
    RtlMoveMemory ByVal ArrPtr(Map), 0&, 4
End Sub
