VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Crypt sample from BlackBeltVB.com"
   ClientHeight    =   6990
   ClientLeft      =   915
   ClientTop       =   1695
   ClientWidth     =   5220
   Height          =   7680
   Left            =   855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   5220
   Top             =   1065
   Width           =   5340
   Begin VB.Frame Frame2 
      Caption         =   "Encryption 2 - Encrypt a data file"
      Height          =   1995
      Left            =   0
      TabIndex        =   9
      Top             =   4920
      Width           =   5175
      Begin VB.CommandButton Command2 
         Caption         =   "Encrypt the file."
         Height          =   315
         Left            =   1020
         TabIndex        =   14
         Top             =   1500
         Width           =   3555
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   900
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Enter the output filename:"
         Height          =   495
         Index           =   1
         Left            =   420
         TabIndex        =   12
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Encrypt a file.  Enter the filename:"
         Height          =   495
         Index           =   0
         Left            =   420
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encryption 1 - Method for INI type files"
      Height          =   4755
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   5175
      Begin VB.TextBox txtPassword 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   1260
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   1815
         Left            =   2820
         TabIndex        =   7
         Top             =   2880
         Width           =   2235
      End
      Begin VB.TextBox Text2 
         Height          =   1815
         Left            =   300
         TabIndex        =   5
         Top             =   2880
         Width           =   2235
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Do The Encryption"
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   1740
         Width           =   3435
      End
      Begin VB.TextBox Text1 
         Height          =   1335
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   15
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Decrypted Data (taken from the Encryption Data box)"
         Height          =   615
         Index           =   2
         Left            =   3060
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Encrypted Data"
         Height          =   315
         Index           =   1
         Left            =   540
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Enter text to encrypt:"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' Crypt sample from BlackBeltVB.com
' http://blackbeltvb.com
'
' Written by Matt Hart
' Copyright 1999 by Matt Hart
'
' This software is FREEWARE. You may use it as you see fit for 
' your own projects but you may not re-sell the original or the 
' source code. Do not copy this sample to a collection, such as
' a CD-ROM archive. You may link directly to the original sample
' using "http://blackbeltvb.com/crypt.htm"
' 
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.
'
' I've updated this sample with better encryption.
' Before, my simple XOR encryption could actually
' show the password somewhere in the encrypted
' file (using the file encryption).  This method
' is more secure.  It builds a "key" from the password
' to encrypt and decrypt, rather than using the actual
' password itself for the algorithm.
'
' Note that if you use Crypt$ to decrypt a file's data
' with a single pass, you must uncomment the line
' in the loop, since the initial encryption would have
' been broken into 4096 byte chunks.

Function EncryptINI$(Strg$, Password$)
   Dim b$, S$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, P$
   j = 1
   For i = 1 To Len(Password$)
     P$ = P$ & Asc(Mid$(Password$, i, 1))
   Next
    
   For i = 1 To Len(Strg$)
     A1 = Asc(Mid$(P$, j, 1))
     j = j + 1: If j > Len(P$) Then j = 1
     A2 = Asc(Mid$(Strg$, i, 1))
     A3 = A1 Xor A2
     b$ = Hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b$
     S$ = S$ + b$
   Next
   EncryptINI$ = S$
End Function

Function DecryptINI$(Strg$, Password$)
   Dim b$, S$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, P$
   j = 1
   For i = 1 To Len(Password$)
     P$ = P$ & Asc(Mid$(Password$, i, 1))
   Next
   
   For i = 1 To Len(Strg$) Step 2
     A1 = Asc(Mid$(P$, j, 1))
     j = j + 1: If j > Len(P$) Then j = 1
     b$ = Mid$(Strg$, i, 2)
     A3 = Val("&H" + b$)
     A2 = A1 Xor A3
     S$ = S$ + Chr$(A2)
   Next
   DecryptINI$ = S$
End Function

Function Crypt$(Strg$, Password$)
   Dim S$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, P$
   j = 1
   For i = 1 To Len(Password$)
     P$ = P$ & Asc(Mid$(Password$, i, 1))
   Next
   For i = 1 To Len(Strg$)
     A1 = Asc(Mid$(P$, j, 1))
     j = j + 1: If j > Len(P$) Then j = 1
     A2 = Asc(Mid$(Strg$, i, 1))
     A3 = A1 Xor A2
     S$ = S$ + Chr$(A3)
     'if i mod 4096 = 0 then j=1
   Next
   Crypt$ = S$
End Function

Private Sub Command1_Click()
    Text2.Text = EncryptINI$(Text1.Text, txtPassword.Text)
    Text3.Text = DecryptINI$(Text2.Text, txtPassword.Text)
End Sub

Private Sub Command2_Click()
    Dim KeyPress
    If Len(Dir$(Text4.Text)) = 0 Then
        MsgBox "File: " & Text4.Text & " not found!": Exit Sub
    End If
    If Len(Dir$(Text5.Text)) Then
        KeyPress = MsgBox("File: " & Text5.Text & " already exists.  Overwrite?", vbYesNoCancel)
        If KeyPress <> vbYes Then Exit Sub
    End If
    On Local Error Resume Next
    Open Text5.Text For Output As 1
    If Err Then
        MsgBox "Error accessing file " & Text5.Text
        Exit Sub
    End If
    Close 1
    Open Text5.Text For Binary As 1
    Open Text4.Text For Binary As 2
    If Err Then
        MsgBox "Error accessing file " & Text4.Text
        Close 1
        Exit Sub
    End If
    Dim l As Long, a$, k As Long, b$
    k = LOF(2) \ 4096
    If k Then
        a$ = Space$(4096)
        For l = 1 To k
            Get 2, , a$
            b$ = Crypt$(a$, txtPassword.Text)
            Put 1, , b$
        Next
    End If
    k = LOF(2) Mod 4096
    If k Then
        a$ = Space$(k)
        Get 2, , a$
        b$ = Crypt$(a$, txtPassword.Text)
        Put 1, , b$
    End If
    Close 1, 2
End Sub

Private Sub mnuAbout_Click()
    MsgBox "There are two Encryption/Decryption routines here.  The first is meant to excrypt data places in an INI file, such as a Password or User ID.  The data returned from the EncryptINI procedure contains only Hex characters - 0-9, A-F.  This makes it easy to put them in an INI file." & vbCrLf & vbCrLf & "The second routine does a straight binary encryption, which can be used to encrypt data files.  The first method requires two functions - one to encrypt the data and another to decrypt it.  The second method will Encrypt the data the first time through, and Decrypt the data the second time through, so only one function is needed."
End Sub
