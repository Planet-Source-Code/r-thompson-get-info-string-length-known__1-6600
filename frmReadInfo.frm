VERSION 5.00
Begin VB.Form frmReadInfo 
   Caption         =   "Customer Information"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnRead 
      Caption         =   "Read Information"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   5115
   End
   Begin VB.Label lblField 
      Height          =   195
      Index           =   2
      Left            =   2700
      TabIndex        =   10
      Top             =   720
      Width           =   2505
   End
   Begin VB.Label lblField 
      Height          =   195
      Index           =   4
      Left            =   2700
      TabIndex        =   8
      Top             =   1320
      Width           =   2505
   End
   Begin VB.Label lblField 
      Height          =   195
      Index           =   3
      Left            =   2700
      TabIndex        =   7
      Top             =   1020
      Width           =   2505
   End
   Begin VB.Label lblField 
      Height          =   195
      Index           =   1
      Left            =   2700
      TabIndex        =   6
      Top             =   420
      Width           =   2505
   End
   Begin VB.Label lblField 
      Height          =   195
      Index           =   0
      Left            =   2700
      TabIndex        =   5
      Top             =   120
      Width           =   2505
   End
   Begin VB.Label lblPostCode 
      Caption         =   "Postcode"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2505
   End
   Begin VB.Label lblSuburb 
      Caption         =   "Suburb"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   2505
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2505
   End
   Begin VB.Label lblSecondName 
      Caption         =   "Second Name"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   2505
   End
   Begin VB.Label lblFirstName 
      Caption         =   "First Name"
      Height          =   200
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2500
   End
End
Attribute VB_Name = "frmReadInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' This easy code is design to open a text file, read a line at a time
' and take a set number of characters from the right hand side of each line

' This is a look at the basics of string and file manipulation
' I will also post a a program that will find a character and
' get the data left or right of it.

' My URL : http://hammer.prohosting.com/~ekans  (up soon)

' Hope this helps


Private Sub btnRead_Click()

Dim lngLength As Long

Open "customers.txt" For Input As #1 'Open File

Line Input #1, strLine 'Read Line 1 into strLine

lngLength = Len(strLine) 'Determine Length of string

intWantedChars = lngLength - 10 'Finds first location of wanted Chars

strField = Right(strLine, intWantedChars) 'Uses Right Function with string and wanted chars position to find letters

lblField(0).Caption = strField 'Puts the chars returned into lblField(0).CaptionText1.Text = strField


Line Input #1, strLine

lngLength = Len(strLine)

intWantedChars = lngLength - 8

strField = Right(strLine, intWantedChars)

lblField(1).Caption = strField


Line Input #1, strLine

lngLength = Len(strLine)

intWantedChars = lngLength - 8

strField = Right(strLine, intWantedChars)

lblField(2).Caption = strField


Line Input #1, strLine

lngLength = Len(strLine)

intWantedChars = lngLength - 7

strField = Right(strLine, intWantedChars)

lblField(3).Caption = strField



Line Input #1, strLine

lngLength = Len(strLine)

intWantedChars = lngLength - 9

strField = Right(strLine, intWantedChars)

lblField(4).Caption = strField





Close #1 'Close The File

End Sub




Private Sub Form_Load()

End Sub
