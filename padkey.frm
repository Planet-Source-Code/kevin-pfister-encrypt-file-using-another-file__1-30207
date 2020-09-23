VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmenc 
   BackColor       =   &H00915320&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text File Encrypter (ONE TIME PAD KEY)"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   Icon            =   "padkey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   8730
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtout 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Text            =   "Cipher.txt"
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txtoutf 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Output File:"
      Top             =   7920
      Width           =   8535
   End
   Begin VB.DriveListBox DRVOUT 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   5880
      TabIndex        =   12
      Top             =   480
      Width           =   2775
   End
   Begin VB.DirListBox DIROUT 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   2565
      Left            =   5880
      TabIndex        =   11
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmddecode 
      Caption         =   "Decode"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   8400
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgdisplay 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   8400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdencode 
      Caption         =   "Encode"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox txtkey 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Key:"
      Top             =   7440
      Width           =   8535
   End
   Begin VB.TextBox txtfile 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "File:"
      Top             =   6960
      Width           =   8535
   End
   Begin VB.FileListBox filkey 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   3405
      Left            =   3000
      TabIndex        =   5
      Top             =   3480
      Width           =   2775
   End
   Begin VB.DirListBox Dirkey 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   2565
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.DriveListBox drvkey 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.FileListBox Filtext 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   3405
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   2775
   End
   Begin VB.DirListBox Dirtext 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.DriveListBox drvtext 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "4. Press Encode or Decode depending on which mode you want."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   855
      Left            =   5880
      TabIndex        =   21
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Choose the output file for the encryption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   5880
      TabIndex        =   20
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Choose a file as a key (must be a Text File, and bigger than the first)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   855
      Left            =   5880
      TabIndex        =   19
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Choose a file to encrypt (must be a Text File)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   5880
      TabIndex        =   18
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Output File :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   6000
      TabIndex        =   17
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Key File :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   3120
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text File :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmenc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AL$(1 To 255), outtext$(1 To 10000)
Dim KEY$, TXT$

Private Sub cmddecode_Click()
    Open TXT$ For Input As #1   'OPENS THE FILE FOR THE TEXT TO BE DECODED
        stringa$ = Input$(LOF(1), 1)
    Close #1    'CLOSES THE FILE
    Open KEY$ For Input As #1   'OPENS THE FILE FOR USE AS THE KEY
        KEYR$ = Input$(LOF(1), 1)
    Close #1    'CLOSES THE FILE
    For a = 1 To 255    'CREATES THE ORGINAL SET OF THE 255 CHARACTERS
        alpha$ = alpha$ + Chr$(a)
    Next
    SL1 = Len(stringa$)
    prgdisplay.Max = Int(SL1 / 1000)    'SETS THE PROGRESS BAR
    For sl = 1 To SL1   'LOOPS UNTIL EVERY CHARACTER HAS BEEN DECODED
        prgdisplay = Int(sl / 1000) 'UPDATES THE PROGRESS BAR
        KEYB$ = Mid$(KEYR$, sl, 1)
        SB$ = Mid$(stringa$, sl, 1)
        KNO = Asc(KEYB$)
        For D = 1 To 255    'THIS LOOP FINDS THE MATCHING CHARACTER
            If Mid$(AL$(KNO), D, 1) = SB$ Then
                Exit For    'EXITS THE LOOP AND RETURNS D
            End If
        Next D
        outtext$(Int(sl / 1000) + 1) = outtext$(Int(sl / 1000) + 1) + Mid$(alpha$, D, 1)
    Next sl
    If Right(DIROUT.Path, 2) = ":\" Then    'THIS CORRECTS THE PATH
        Output$ = DIROUT.Path + txtout.Text 'THIS IS IF IT IS IN THE ROOT DIR EG. C:\
    Else
        Output$ = DIROUT.Path + "\" + txtout.Text
    End If
    Open Output$ For Output As #1   'OPENS THE OUTPUT FILE
        For a = 1 To 10000 'THIS WRITES THE TEXT ARRAY TO A FILE
            If outtext$(a) <> "" Then   'ONLY WRITES THE ARRAYS WHICH HAVE CHARACTERS IN IT
                Print #1, outtext$(a)
            End If
        Next
    Close #1 'CLOSES THE OUTPUT FILE
    For a = 1 To 10000  'THIS EMPTIES THE ARRAY
        outtext$(a) = ""
    Next
End Sub

Private Sub cmdencode_Click()
    Open TXT$ For Input As #1   'OPENS THE FILE TO BE ENCODED
        stringa$ = Input$(LOF(1), 1)
    Close #1    'CLOSES THE FILE
    Open KEY$ For Input As #1   'OPENS THE FILE FOR USE AS THE KEY
        KEYR$ = Input$(LOF(1), 1)
    Close #1    'CLOSES THE FILE
    If Right(DIROUT.Path, 2) = ":\" Then    'THIS CORRECTS THE PATH
        Output$ = DIROUT.Path + txtout.Text 'THIS IS IF IT IS IN THE ROOT DIR EG. C:\
    Else
        Output$ = DIROUT.Path + "\" + txtout.Text
    End If
    Open Output$ For Output As #1   'OPENS THE OUTPUT FILE
        sl = Len(stringa$)
        For b = 1 To sl 'LOOPS UNTIL EVERY CHARACTER HAS BEEN ENCODED
            KEYLET$ = Mid$(KEYR$, b, 1)
            SLET$ = Mid$(stringa$, b, 1)
            KNO = Asc(KEYLET$)
            SNO = Asc(SLET$)
            Print #1, Mid$(AL$(KNO), SNO, 1);   'THIS WRITES THE CHARACTER TO THE FILE
        Next b
        MsgBox "ENCODING COMPLETE"
    Close #1    'CLOSES THE FILE
End Sub

Private Sub Dirkey_Change()
    filkey = Dirkey 'UPDATES THE FILE BOX
    txtkey = "Key:" 'CLEARS THE KEY TXT BOX
    cmdencode.Enabled = False
    cmddecode.Enabled = False
End Sub

Private Sub DIROUT_Change()
    If Right(DIROUT.Path, 2) = ":\" Then    'THIS CORRECTS THE PATH
        txtoutf = "Output File : " + DIROUT.Path + txtout.Text  'THIS IS IF IT IS THE ROOT DIR EG. C:\
    Else
        txtoutf = "Output File : " + DIROUT.Path + "\" + txtout.Text
    End If
End Sub

Private Sub Dirtext_Change()
    Filtext = Dirtext   'UPDATES THE FILE BOX
    txtfile = "File:"   'CLEARS THE FILE TXT BOX
    cmdencode.Enabled = False
    cmddecode.Enabled = False
End Sub

Private Sub drvkey_Change()
    Dirkey = drvkey 'UPDATES THE DIRECTORY BOX
    filkey = Dirkey 'UPDATES THE FILE BOX
    txtkey = "Key:" 'CLEARS THE FILE TXT BOX
    cmdencode.Enabled = False
    cmddecode.Enabled = False
End Sub

Private Sub DRVOUT_Change()
    DIROUT = DRVOUT 'UPDATES THE DIRECTORY BOX
    If Right(DIROUT.Path, 2) = ":\" Then    'CORRECTS THE PATH
        txtoutf = "Output File : " + DIROUT.Path + txtout.Text  'THIS IS IF IT IS IN THE ROOT DIRECTORY EG. C:\
    Else
        txtoutf = "Output File : " + DIROUT.Path + "\" + txtout.Text
    End If
End Sub

Private Sub drvtext_Change()
    Dirtext = drvtext   'UPDATES THE DIRECTORY BOX
    Filtext = Dirtext   'UPDATES THE FILE BOX
    txtfile = "File:"   'CLEARS THE FILE TEXT BOX
    cmdencode.Enabled = False
    cmddecode.Enabled = False
End Sub

Private Sub filkey_Click()
    If Right(filkey.Path, 2) = ":\" Then    'UPDATES THE KEY TEXT BOX
        txtkey = "Key : " + filkey.Path + filkey
        KEY$ = filkey.Path + filkey
    Else
        txtkey = "Key : " + filkey.Path + "\" + filkey
        KEY$ = filkey.Path + "\" + filkey
    End If
    If txtfile <> "File:" Then  'ONLY ENABLES THE BUTTONS WHEN BOTH FILES ARE SELECTED
        cmdencode.Enabled = True
        cmddecode.Enabled = True
    Else
        cmdencode.Enabled = False
        cmddecode.Enabled = False
    End If
End Sub

Private Sub Filtext_Click()
    If Right(Filtext.Path, 2) = ":\" Then   'UPDATES THE FILE TEXT BOX
        txtfile = "File : " + Filtext.Path + Filtext
        TXT$ = Filtext.Path + Filtext
    Else
        txtfile = "File : " + Filtext.Path + "\" + Filtext
        TXT$ = Filtext.Path + "\" + Filtext
    End If
    If txtkey <> "Key:" Then    'ONLY ENABLES THE BUTTONS WHEN BOTH FIELS ARE SELECTED
        cmdencode.Enabled = True
        cmddecode.Enabled = True
    Else
        cmdencode.Enabled = False
        cmddecode.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    For a = 1 To 255    'CREATES THE ORGINAL 255 CHARACTERS
        alpha$ = alpha$ + Chr$(a)
    Next
    For a = 1 To 255    'THIS LOOPS CREATES THE UPDATED VIGENERE SQUARE
        AL$(a) = Mid$(alpha$, a)
        If a <> 1 Then
            AL$(a) = AL$(a) + Mid$(alpha$, 1, a - 1)
        End If
    Next a
End Sub

Private Sub txtout_Change()
    If Right(DIROUT.Path, 2) = ":\" Then    'UPDATES THE OUTPUT FILE BOX
        txtoutf = "Output File : " + DIROUT.Path + txtout.Text
    Else
        txtoutf = "Output File : " + DIROUT.Path + "\" + txtout.Text
    End If
End Sub
