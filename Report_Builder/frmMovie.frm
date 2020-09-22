VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMovie 
   BorderStyle     =   0  'None
   ClientHeight    =   7845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12810
   ControlBox      =   0   'False
   Icon            =   "frmMovie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMovie.frx":0442
   ScaleHeight     =   7845
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1110
      Left            =   7665
      Picture         =   "frmMovie.frx":16F1AC
      ScaleHeight     =   1050
      ScaleWidth      =   1425
      TabIndex        =   31
      Top             =   1530
      Width           =   1485
   End
   Begin VB.CommandButton cmdMovieNameDelete 
      Caption         =   "&Delete"
      Height          =   525
      Left            =   11640
      MouseIcon       =   "frmMovie.frx":16FB35
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   7020
      Width           =   1050
   End
   Begin MSComctlLib.ListView lvwMovieName 
      Height          =   6210
      Left            =   9270
      TabIndex        =   27
      Top             =   780
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   10954
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMovie.frx":16FE3F
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Movie Name"
         Object.Width           =   5821
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMovieReport 
      Height          =   4230
      Left            =   45
      TabIndex        =   26
      Top             =   2700
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   7461
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMovie.frx":170159
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Movie Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Writer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Co-writer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Producer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Director"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Actor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Actress"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Supporting Actor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Comments"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   795
      Left            =   60
      TabIndex        =   20
      Top             =   6945
      Width           =   9195
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   525
         Left            =   7410
         MouseIcon       =   "frmMovie.frx":170473
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   165
         Width           =   1125
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   525
         Left            =   5955
         MouseIcon       =   "frmMovie.frx":17077D
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   525
         Left            =   4515
         MouseIcon       =   "frmMovie.frx":170A87
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   525
         Left            =   3120
         MouseIcon       =   "frmMovie.frx":170D91
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   525
         Left            =   1770
         MouseIcon       =   "frmMovie.frx":17109B
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   525
         Left            =   465
         MouseIcon       =   "frmMovie.frx":1713A5
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4620
      TabIndex        =   9
      Top             =   2085
      Width           =   1815
   End
   Begin VB.TextBox txtSupportingActor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1620
      TabIndex        =   8
      Top             =   2100
      Width           =   1815
   End
   Begin VB.TextBox txtVillen 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4620
      TabIndex        =   7
      Top             =   1650
      Width           =   1815
   End
   Begin VB.TextBox txtActor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1605
      TabIndex        =   6
      Top             =   1650
      Width           =   1815
   End
   Begin VB.TextBox txtDirector 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7350
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtProducer 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4635
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtCoWriter 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1650
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtWriter 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7335
      TabIndex        =   2
      Top             =   780
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1635
      TabIndex        =   1
      Top             =   765
      Width           =   1815
   End
   Begin VB.TextBox txtMovieType 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4650
      TabIndex        =   0
      Top             =   765
      Width           =   1815
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6555
      TabIndex        =   32
      Top             =   2340
      Width           =   1005
   End
   Begin VB.Image imgClose 
      Height          =   330
      Left            =   12420
      MouseIcon       =   "frmMovie.frx":1716AF
      MousePointer    =   99  'Custom
      Top             =   45
      Width           =   315
   End
   Begin VB.Image imgMinimize 
      Height          =   315
      Left            =   12090
      MouseIcon       =   "frmMovie.frx":1719B9
      MousePointer    =   99  'Custom
      Top             =   45
      Width           =   315
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hollywood/Bollywood Movies Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   28
      Top             =   -15
      Width           =   4965
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3555
      TabIndex        =   19
      Top             =   2100
      Width           =   960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supporting Actor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   18
      Top             =   2070
      Width           =   1485
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3645
      TabIndex        =   17
      Top             =   1650
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1065
      TabIndex        =   16
      Top             =   1650
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Director"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6570
      TabIndex        =   15
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3585
      TabIndex        =   14
      Top             =   1230
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Co-Writer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   735
      TabIndex        =   13
      Top             =   1230
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Writer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6540
      TabIndex        =   12
      Top             =   795
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1155
      TabIndex        =   11
      Top             =   780
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MovieType"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3555
      TabIndex        =   10
      Top             =   780
      Width           =   1035
   End
End
Attribute VB_Name = "frmMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------
'       Developed by            Anuj sharma
'       Product Name            Report Builder v1.1
'       Email                   anujsharrma@yahoo.com
'-----------------------------------------------------


Private Sub cmdAdd_Click()
    txtTitle.Text = ""
    txtActor.Text = ""
    txtComments.Text = ""
    txtCoWriter.Text = ""
    txtDirector.Text = ""
    txtMovieType.Text = ""
    txtProducer.Text = ""
    txtSupportingActor.Text = ""
    txtVillen.Text = ""
    txtWriter.Text = ""
    txtMovieType.Enabled = True
    txtTitle.Enabled = True
    txtWriter.Enabled = True
    txtCoWriter.Enabled = True
    txtProducer.Enabled = True
    txtDirector.Enabled = True
    txtActor.Enabled = True
    txtVillen.Enabled = True
    txtSupportingActor.Enabled = True
    txtComments.Enabled = True
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdUpdate.Enabled = False
    lvwMovieReport.Enabled = False
    txtTitle.SetFocus
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
Dim iCount As Integer
Dim lvwItem As ListItem
Dim sFileName As String
Dim iFileNo  As Integer
Dim sArr() As String
Dim sTypeSpyware As String
Dim sData As String

    sFileName = App.Path & "\MovieReport.txt"
     If StrComp(UCase(Trim$(Dir(sFileName))), UCase("MovieReport.txt"), vbTextCompare) = 0 Then
        iFileNo = FreeFile
        Open sFileName For Input As #iFileNo
            lvwMovieReport.ListItems.Clear
            Do While Not EOF(iFileNo)
                Line Input #iFileNo, sData
                sArr = Split(sData, ",", , vbTextCompare)
                Set lvwItem = lvwMovieReport.ListItems.Add(, , sArr(0))
                    With lvwItem
                        .SubItems(1) = Trim$(sArr(1))
                        .SubItems(2) = Trim$(sArr(2))
                        .SubItems(3) = Trim$(sArr(3))
                        .SubItems(4) = Trim$(sArr(4))
                        .SubItems(5) = Trim$(sArr(5))
                        .SubItems(2) = Trim$(sArr(6))
                        .SubItems(3) = Trim$(sArr(7))
                        .SubItems(4) = Trim$(sArr(8))
                        .SubItems(5) = Trim$(sArr(9))
                    End With
                DoEvents
            Loop
            lvwMovieReport.Enabled = True
        Close #iFileNo
        cmdUpdate.Enabled = True
        cmdSave.Enabled = False
        cmdAdd.Enabled = True
        cmdCancel.Enabled = False
        cmdDelete.Enabled = True
        
        txtTitle.Text = ""
        txtActor.Text = ""
        txtComments.Text = ""
        txtCoWriter.Text = ""
        txtDirector.Text = ""
        txtMovieType.Text = ""
        txtProducer.Text = ""
        txtSupportingActor.Text = ""
        txtVillen.Text = ""
        txtWriter.Text = ""
        
        txtTitle.Enabled = False
        txtActor.Enabled = False
        txtComments.Enabled = False
        txtCoWriter.Enabled = False
        txtDirector.Enabled = False
        txtMovieType.Enabled = False
        txtProducer.Enabled = False
        txtSupportingActor.Enabled = False
        txtVillen.Enabled = False
        txtWriter.Enabled = False
        
        lvwMovieReport.ListItems.Item(1).Selected = True
        lvwMovieReport.FullRowSelect = True
    End If
    iCount = lvwMovieReport.ListItems.Count
    If iCount > 0 Then
        cmdUpdate.Enabled = True
        lvwMovieReport.ListItems.Item(1).Selected = True
        lvwMovieReport.FullRowSelect = True
    Else
        cmdSave.Enabled = False
        cmdAdd.Enabled = True
        cmdCancel.Enabled = False
        cmdDelete.Enabled = False
        lvwMovieReport.Enabled = True
        txtTitle.Text = ""
        txtActor.Text = ""
        txtComments.Text = ""
        txtCoWriter.Text = ""
        txtDirector.Text = ""
        txtMovieType.Text = ""
        txtProducer.Text = ""
        txtSupportingActor.Text = ""
        txtVillen.Text = ""
        txtWriter.Text = ""
        
        txtTitle.Enabled = False
        txtActor.Enabled = False
        txtComments.Enabled = False
        txtCoWriter.Enabled = False
        txtDirector.Enabled = False
        txtMovieType.Enabled = False
        txtProducer.Enabled = False
        txtSupportingActor.Enabled = False
        txtVillen.Enabled = False
        txtWriter.Enabled = False
    End If
Exit Sub
Err_Handler:
    Call MakeLogFile("cmdCancel_Click")
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err_Handler
Dim sFileName As String
Dim iFileNo As Integer
Dim iCount As Integer
Dim lvwItem As ListItem
Dim sWriter As String
Dim sCoWriter As String
Dim sProducer  As String
Dim sDirector As String
Dim sActor As String
Dim sSupportingActor As String
Dim sComments As String

    iCount = lvwMovieReport.ListItems.Count
    If iCount > 0 Then
        cmdDelete.Enabled = False
        lvwMovieReport.ListItems.Remove (lvwMovieReport.SelectedItem.Index)
        txtTitle.Text = ""
        txtActor.Text = ""
        txtComments.Text = ""
        txtCoWriter.Text = ""
        txtDirector.Text = ""
        txtMovieType.Text = ""
        txtProducer.Text = ""
        txtSupportingActor.Text = ""
        txtVillen.Text = ""
        txtWriter.Text = ""
        Kill (App.Path & "\MovieReport.txt")
        sFileName = App.Path & "\MovieReport.txt"
        iFileNo = FreeFile
        iCount = lvwMovieReport.ListItems.Count
        If iCount > 0 Then
            Open sFileName For Output As #iFileNo
                For iCount = 1 To lvwMovieReport.ListItems.Count
                    Set lvwItem = lvwMovieReport.ListItems.Item(iCount)
                        sTitle = Trim(lvwItem.SubItems(1))
                        sWriter = Trim(lvwItem.SubItems(2))
                        sCoWriter = Trim(lvwItem.SubItems(3))
                        sProducer = Trim(lvwItem.SubItems(4))
                        sDirector = Trim(lvwItem.SubItems(5))
                        sActor = Trim(lvwItem.SubItems(6))
                        sActress = Trim(lvwItem.SubItems(7))
                        sSupportingActor = Trim(lvwItem.SubItems(8))
                        sComments = Trim(lvwItem.SubItems(9))
                        Print #iFileNo, lvwItem, ","; sTitle, ","; sWriter, ","; sCoWriter, ","; sProducer, ","; sDirector, ","; sActor; ","; sActress, ","; sSupportingActor, ","; sComments, ""
                        DoEvents
                Next
            Close #iFileNo
        End If
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
        cmdUpdate.Enabled = False
    End If
    
Exit Sub
Err_Handler:
    Call MakeLogFile("cmdDelete_Click")
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdMovieNameDelete_Click()
On Error GoTo Err_Handler
Dim sFileName As String
Dim iFileNo As Integer
Dim iCount As Integer
Dim lvwItem As ListItem
Dim sSpywareName  As String

    cmdMovieNameDelete.Enabled = False
    iCount = lvwMovieName.ListItems.Count
    If iCount > 0 Then
        lvwMovieName.ListItems.Remove (lvwMovieName.SelectedItem.Index)
        Kill (App.Path & "\MovieName.txt")
        sFileName = App.Path & "\MovieName.txt"
            iFileNo = FreeFile
            iCount = lvwMovieName.ListItems.Count
            If iCount > 0 Then
                Open sFileName For Output As #iFileNo
                    For iCount = 1 To lvwMovieName.ListItems.Count
                        Set lvwItem = lvwMovieName.ListItems.Item(iCount)
                        Print #iFileNo, lvwItem
                        DoEvents
                    Next
                Close #iFileNo
            End If
    End If
    cmdMovieNameDelete.Enabled = True
Exit Sub
Err_Handler:
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_Handler
Dim lvwItem As ListItem
Dim iCount As Integer
Dim sFileName As String
Dim iFileNo As Integer
Dim sWriter As String
Dim sCoWriter As String
Dim sProducer  As String
Dim sDirector As String
Dim sActor As String
Dim sSupportingActor As String
Dim sComments As String

If Trim$(txtMovieType.Text) = "" Then
        MsgBox "Title can't be left empty.", vbInformation + vbOKOnly, App.Title
        txtTitle.SetFocus
        Exit Sub
    Else
        Set lvwItem = lvwMovieReport.FindItem(Trim$(txtTitle.Text), lvwSubItem)
        If lvwItem Is Nothing Then
            Set lvwItem = lvwMovieReport.ListItems.Add(, , Trim$(txtMovieType.Text))
            With lvwItem
                .SubItems(1) = Trim$(txtTitle.Text)
                .SubItems(2) = Trim$(txtWriter.Text)
                .SubItems(3) = Trim$(txtCoWriter.Text)
                .SubItems(4) = Trim$(txtProducer.Text)
                .SubItems(5) = Trim$(txtDirector.Text)
                .SubItems(6) = Trim$(txtActor.Text)
                .SubItems(7) = Trim$(txtVillen.Text)
                .SubItems(8) = Trim$(txtSupportingActor.Text)
                .SubItems(9) = Trim$(txtComments.Text)
            End With
            Set lvwItem = lvwMovieName.FindItem(Trim$(txtTitle.Text))
            If lvwItem Is Nothing Then
                lvwMovieName.ListItems.Add , , Trim$(txtTitle.Text)
            End If
            sFileName = App.Path & "\MovieReport.txt"
            iFileNo = FreeFile
            iCount = lvwMovieReport.ListItems.Count
                Open sFileName For Output As #iFileNo
                    For iCount = 1 To lvwMovieReport.ListItems.Count
                        Set lvwItem = lvwMovieReport.ListItems.Item(iCount)
                        sTitle = Trim(lvwItem.SubItems(1))
                        sWriter = Trim(lvwItem.SubItems(2))
                        sCoWriter = Trim(lvwItem.SubItems(3))
                        sProducer = Trim(lvwItem.SubItems(4))
                        sDirector = Trim(lvwItem.SubItems(5))
                        sActor = Trim(lvwItem.SubItems(6))
                        sActress = Trim(lvwItem.SubItems(7))
                        sSupportingActor = Trim(lvwItem.SubItems(8))
                        sComments = Trim(lvwItem.SubItems(9))
                        Print #iFileNo, lvwItem, ","; sTitle, ","; sWriter, ","; sCoWriter, ","; sProducer, ","; sDirector, ","; sActor; ","; sActress, ","; sSupportingActor, ","; sComments, ""
                        DoEvents
                    Next
                Close #iFileNo
                cmdAdd.Enabled = True
                cmdSave.Enabled = False
                cmdUpdate.Enabled = True
                cmdCancel.Enabled = False
                cmdDelete.Enabled = True
                
                txtTitle.Enabled = False
                txtActor.Enabled = False
                txtComments.Enabled = False
                txtCoWriter.Enabled = False
                txtDirector.Enabled = False
                txtMovieType.Enabled = False
                txtProducer.Enabled = False
                txtSupportingActor.Enabled = False
                txtVillen.Enabled = False
                txtWriter.Enabled = False
                lvwMovieReport.Enabled = True
        Else
            MsgBox "Field already exists.", vbInformation + vbOKOnly, App.Title
            txtTitle.Text = ""
            txtTitle.SetFocus
        End If
    End If
Exit Sub
Err_Handler:
    Call MakeLogFile("cmdSave_Click")
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Err_Handler
Dim lvwItem As ListItem
Dim sWriter As String
Dim sCoWriter As String
Dim sProducer  As String
Dim sDirector As String
Dim sActor As String
Dim sSupportingActor As String
Dim sComments As String
Dim iCount As Integer
Dim sTitle As String
Dim sActress As String

    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    lvwMovieReport.Enabled = False
    iCount = lvwMovieReport.ListItems.Count
    If iCount > 0 Then
        Set lvwItem = lvwMovieReport.ListItems.Item(lvwMovieReport.SelectedItem.Index)
            sTitle = Trim(lvwItem.SubItems(1))
            sWriter = Trim(lvwItem.SubItems(2))
            sCoWriter = Trim(lvwItem.SubItems(3))
            sProducer = Trim(lvwItem.SubItems(4))
            sDirector = Trim(lvwItem.SubItems(5))
            sActor = Trim(lvwItem.SubItems(6))
            sActress = Trim(lvwItem.SubItems(7))
            sSupportingActor = Trim(lvwItem.SubItems(8))
            sComments = Trim(lvwItem.SubItems(9))
            
            txtMovieType.Text = lvwItem
            txtTitle.Text = sTitle
            txtWriter.Text = sWriter
            txtCoWriter.Text = sCoWriter
            txtProducer.Text = sProducer
            txtDirector.Text = sDirector
            txtActor.Text = sActor
            txtVillen.Text = sActress
            txtSupportingActor.Text = sSupportingActor
            txtComments.Text = sComments
            
            txtMovieType.Enabled = True
            txtTitle.Enabled = True
            txtWriter.Enabled = True
            txtCoWriter.Enabled = True
            txtProducer.Enabled = True
            txtDirector.Enabled = True
            txtActor.Enabled = True
            txtVillen.Enabled = True
            txtSupportingActor.Enabled = True
            txtComments.Enabled = True
            lvwMovieReport.ListItems.Remove (lvwMovieReport.SelectedItem.Index)
    End If
Exit Sub
Err_Handler:
    Call MakeLogFile("cmdUpdate_Click")
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim iCount As Integer
Dim sFileName As String
Dim sData As String
Dim sFilesArr() As String
Dim iMovieFileNo As Integer
Dim sMovieFileName  As String
    
'    frmMovie.Caption = App.ProductName
    sFileName = App.Path & "\MovieReport.txt"
    If StrComp(UCase(Trim$(Dir(sFileName))), UCase("MovieReport.txt"), vbTextCompare) = 0 Then
        iFileNo = FreeFile
        Open sFileName For Input As #iFileNo
            Do While Not EOF(iFileNo)
                Line Input #iFileNo, sData
                sFilesArr = Split(sData, ",", , vbTextCompare)
                Set lvwItem = lvwMovieReport.ListItems.Add(, , Trim$(sFilesArr(0)))
                    With lvwItem
                        .SubItems(1) = Trim(sFilesArr(1))
                        .SubItems(2) = Trim$(sFilesArr(2))
                        .SubItems(3) = Trim$(sFilesArr(3))
                        .SubItems(4) = Trim$(sFilesArr(4))
                        .SubItems(5) = Trim$(sFilesArr(5))
                        .SubItems(6) = Trim$(sFilesArr(6))
                        .SubItems(7) = Trim$(sFilesArr(7))
                        .SubItems(8) = Trim$(sFilesArr(8))
                        .SubItems(9) = Trim$(sFilesArr(9))
                    End With
                DoEvents
            Loop
        Close #iFileNo
    End If
    '-----------------------------------------------------------------------------------------------
    '               Load MovieName in listview
    '-----------------------------------------------------------------------------------------------
    sMovieFileName = App.Path & "\MovieName.txt"
    If StrComp(UCase(Trim$(Dir(sMovieFileName))), UCase("MovieName.txt"), vbTextCompare) = 0 Then
        iMovieFileNo = FreeFile
            Open sMovieFileName For Input As #iMovieFileNo
                Do While Not EOF(iMovieFileNo)
                    Line Input #iMovieFileNo, sData
                    lvwMovieName.ListItems.Add , , sData
                    DoEvents
                Loop
            Close #iMovieFileNo
    End If
    iCount = lvwMovieReport.ListItems.Count
    If iCount > 0 Then
        cmdDelete.Enabled = True
        cmdUpdate.Enabled = True
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        txtTitle.Enabled = False
        txtActor.Enabled = False
        txtComments.Enabled = False
        txtCoWriter.Enabled = False
        txtDirector.Enabled = False
        txtMovieType.Enabled = False
        txtProducer.Enabled = False
        txtSupportingActor.Enabled = False
        txtVillen.Enabled = False
        txtWriter.Enabled = False
        lvwMovieReport.ListItems.Item(1).Selected = True
    Else
        cmdDelete.Enabled = False
        cmdUpdate.Enabled = False
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        txtTitle.Enabled = False
        txtActor.Enabled = False
        txtComments.Enabled = False
        txtCoWriter.Enabled = False
        txtDirector.Enabled = False
        txtMovieType.Enabled = False
        txtProducer.Enabled = False
        txtSupportingActor.Enabled = False
        txtVillen.Enabled = False
        txtWriter.Enabled = False
        lvwMovieReport.Enabled = True
    End If
Exit Sub
Err_Handler:
    Call MakeLogFile("Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Err_Handler
Dim sMovieFileName  As String
Dim iFileNo As Integer
Dim iSpyListCount As Integer
Dim lvwItem As ListItem
Dim iCount As Integer
    iCount = lvwMovieName.ListItems.Count
    If iCount > 0 Then
        sMovieFileName = App.Path & "\MovieName.txt"
        If StrComp(UCase(Trim$(Dir(sMovieFileName))), UCase("MovieName.txt"), vbTextCompare) = 0 Then
            Kill (App.Path & "\MovieName.txt")
            iFileNo = FreeFile
            iCount = lvwMovieName.ListItems.Count
            Open sMovieFileName For Append As #iFileNo
                For iCount = 1 To lvwMovieName.ListItems.Count
                    Set lvwItem = lvwMovieName.ListItems.Item(iCount)
                    Print #iFileNo, lvwItem
                    DoEvents
                Next
            Close #iFileNo
        Else
            iFileNo = FreeFile
            iCount = lvwMovieName.ListItems.Count
            Open sMovieFileName For Output As #iFileNo
                For iCount = 1 To lvwMovieName.ListItems.Count
                    Set lvwItem = lvwMovieName.ListItems.Item(iCount)
                    Print #iFileNo, lvwItem
                    DoEvents
                Next
            Close #iFileNo
        End If
    End If
Exit Sub
Err_Handler:
    Call MakeLogFile("Form_QueryUnload")
End Sub

Private Sub imgClose_Click()
    Unload Me
    End
End Sub

Private Sub imgMinimize_Click()
    WindowState = vbMinimized
End Sub

Private Sub lvwMovieReport_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err_Handler
Dim lvwItem As ListItem
Dim sWriter As String
Dim sCoWriter As String
Dim sProducer  As String
Dim sDirector As String
Dim sActor As String
Dim sSupportingActor As String
Dim sComments As String
Dim iCount As Integer
Dim sTitle As String
Dim sActress As String

    iCount = lvwMovieReport.ListItems.Count
    If iCount > 0 Then
        Set lvwItem = lvwMovieReport.ListItems.Item(lvwMovieReport.SelectedItem.Index)
            sTitle = Trim(lvwItem.SubItems(1))
            sWriter = Trim(lvwItem.SubItems(2))
            sCoWriter = Trim(lvwItem.SubItems(3))
            sProducer = Trim(lvwItem.SubItems(4))
            sDirector = Trim(lvwItem.SubItems(5))
            sActor = Trim(lvwItem.SubItems(6))
            sActress = Trim(lvwItem.SubItems(7))
            sSupportingActor = Trim(lvwItem.SubItems(8))
            sComments = Trim(lvwItem.SubItems(9))
            
            txtMovieType.Text = lvwItem
            txtTitle.Text = sTitle
            txtWriter.Text = sWriter
            txtCoWriter.Text = sCoWriter
            txtProducer.Text = sProducer
            txtDirector.Text = sDirector
            txtActor.Text = sActor
            txtVillen.Text = sActress
            txtSupportingActor.Text = sSupportingActor
            txtComments.Text = sComments
    End If
Exit Sub
Err_Handler:
    Call MakeLogFile("lvwMovieReport_ItemClick")
End Sub
