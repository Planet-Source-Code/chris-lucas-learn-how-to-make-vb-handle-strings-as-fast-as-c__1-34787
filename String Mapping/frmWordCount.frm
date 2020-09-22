VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWordCount 
   Caption         =   "String Mapping - Word Count Demo"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7725
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompare 
      Caption         =   "&Compare Methods"
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtIterations 
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Text            =   "1000"
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtWords 
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Text            =   "100"
      Top             =   3720
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvwResults 
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2778
      SortKey         =   2
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Method"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Counting Result"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Execution Time (ms)"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Relative Call Length"
         Object.Width           =   4939
      EndProperty
   End
   Begin VB.Frame fraFunctions 
      Caption         =   "Functions To Test:"
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   2895
      Begin VB.CheckBox chkFunctions 
         Caption         =   "String Mapping Method"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkFunctions 
         Caption         =   "InStr Method"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkFunctions 
         Caption         =   "Asc(Mid$) Method"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkFunctions 
         Caption         =   "Split Method"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Value           =   1  'Checked
         Width           =   2415
      End
   End
   Begin VB.TextBox txtSample 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Itertions:"
      Height          =   240
      Left            =   6270
      TabIndex        =   2
      Top             =   4260
      Width           =   750
   End
   Begin VB.Label lblWords 
      AutoSize        =   -1  'True
      Caption         =   "&Words per string:"
      Height          =   240
      Left            =   5520
      TabIndex        =   0
      Top             =   3780
      Width           =   1500
   End
End
Attribute VB_Name = "frmWordCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Classes we'll be needing
Private WordCounter As clsWordCount     ' Contains the word count rountines
Private Timer As PrecisionTimer         ' Implements QueryPerformanceCounter for timing

Private Sub cmdCompare_Click()
    Dim i As Long
    Dim ret As Long
    Dim MAX_ITERATIONS As Long
    Dim dblMax As Double
    Dim strText As String
    Dim lvwItem As ListItem
    
    lvwResults.ListItems.Clear
    lvwResults.GridLines = True
    MAX_ITERATIONS = CLng(txtIterations)
    
    ' Cache the string
    strText = txtSample.Text
    
    Screen.MousePointer = vbHourglass
    
    ' Call the appropriate functions and display the timings
    If chkFunctions(0) = vbChecked Then
        ' Time the function
        Timer.ResetTimer
        For i = 1 To MAX_ITERATIONS
            ret = WordCounter.SplitWordCount(strText)
        Next i
        Timer.StopTimer
        
        ' Display the results
        Set lvwItem = lvwResults.ListItems.Add
        lvwItem.Text = "Split"
        lvwItem.SubItems(1) = ret
        lvwItem.SubItems(2) = Format$(Timer.Elapsed / 1000, "###,###.00")
        
        
    End If
    
    DoEvents
    
    If chkFunctions(1) = vbChecked Then
        ' Time the function
        Timer.ResetTimer
        For i = 1 To MAX_ITERATIONS
            ret = WordCounter.MidWordCount(strText)
        Next i
        Timer.StopTimer
        
        ' Display the results
        Set lvwItem = lvwResults.ListItems.Add
        lvwItem.Text = "Asc(Mid$)"
        lvwItem.SubItems(1) = ret
        lvwItem.SubItems(2) = Format$(Timer.Elapsed / 1000, "###,###.00")
    End If
    
    DoEvents
    
    If chkFunctions(2) = vbChecked Then
        ' Time the function
        Timer.ResetTimer
        For i = 1 To MAX_ITERATIONS
            ret = WordCounter.InStrWordCount(strText)
        Next i
        Timer.StopTimer
        
        ' Display the results
        Set lvwItem = lvwResults.ListItems.Add
        lvwItem.Text = "InStr"
        lvwItem.SubItems(1) = ret
        lvwItem.SubItems(2) = Format$(Timer.Elapsed / 1000, "###,###.00")
    End If
    
    DoEvents
    
    If chkFunctions(3) = vbChecked Then
        ' Time the function
        Timer.ResetTimer
        For i = 1 To MAX_ITERATIONS
            ret = WordCounter.FastWordCount(strText)
        Next i
        Timer.StopTimer
        
        ' Display the results
        Set lvwItem = lvwResults.ListItems.Add
        lvwItem.Text = "String Mapping"
        lvwItem.SubItems(1) = ret
        lvwItem.SubItems(2) = Format$(Timer.Elapsed / 1000, "###,###.00")
    End If
    
    ' Display the relative timing results
    With lvwResults
        If .ListItems.Count > 0 Then
            dblMax = .ListItems(1).SubItems(2)
            ' Find the minimal time
            For i = 1 To .ListItems.Count
                If .ListItems(i).SubItems(2) > dblMax Then
                    dblMax = .ListItems(i).SubItems(2)
                End If
            Next i
            ' Display the relative result
            For i = 1 To .ListItems.Count
                .ListItems(i).SubItems(3) = Format$(.ListItems(i).SubItems(2) / dblMax * 100, "#.0") & " %"
            Next i
        End If
    End With
    
    Screen.MousePointer = vbDefault
    Set lvwItem = Nothing
End Sub

Private Sub Form_Load()
    Set WordCounter = New clsWordCount
    Set Timer = New PrecisionTimer
    txtWords_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Set WordCounter = Nothing
     Set Timer = Nothing
End Sub

Private Sub txtWords_Change()
    On Error GoTo Bail
    Dim i As Long
    
    If Len(txtWords) Then
        If IsNumeric(txtWords) Then
            With txtSample
                .Visible = False
                .Text = vbNullString
                .SelStart = 0
        
                For i = 1 To CLng(txtWords) - 1
                    .SelText = "Words  "
                Next i
                .SelText = "END"
                .Visible = True
            End With
        End If
    End If
    Exit Sub
    
Bail:
    If Err.Number = 7 Then
        MsgBox "VB ran out of memory, use a smaller number", vbCritical
        txtWords.SelStart = 0
        txtWords.SelLength = Len(txtWords)
    End If
End Sub
