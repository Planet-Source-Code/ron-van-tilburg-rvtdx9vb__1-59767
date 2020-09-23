VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "rvtDX9.dll Maker"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMakeDLL 
      Caption         =   "Make DLL"
      Height          =   435
      Left            =   1770
      TabIndex        =   1
      Top             =   2520
      Width           =   1275
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'MakeDLL - s short program which decodes the DLL from the DAT file in order to not put an executable onto PSC

'This reverses the XOR process shown below
Private Sub cmdMakeDLL_Click()
  
 Dim DLL() As Long, DAT() As Long, XORVal As Long, CheckSum As Long, i As Long
 
 Open App.Path & "\rvtDX9vb.dat" For Binary Access Read As #1
 List1.AddItem "Got DAT - Length=" & LOF(1) & " bytes"
 Open App.Path & "\rvtDX9vb.dll" For Binary Access Write As #2
 
 ReDim DAT(0 To LOF(1) \ 4 - 1)
 Get #1, , DAT()
 ReDim DLL(0 To UBound(DAT) - 2) As Long
 
 XORVal = DAT(0)         'RVT9
 CheckSum = DAT(1)
 
 For i = 2 To UBound(DAT)
   DLL(i - 2) = DAT(i) Xor XORVal
   CheckSum = CheckSum Xor DLL(i - 2)
 Next
 Put #2, , DLL()
 List1.AddItem "DLL Written - Length=" & LOF(2) & " bytes"
 Close #1, #2
 
 If CheckSum = 0 Then
    List1.AddItem "Decompile successful"
    List1.AddItem "copy .dll and .tlb to"
    List1.AddItem "C:\WINDOWS\SYSTEM32 (WinXP) or"
    List1.AddItem "C:\WINDOWS\SYSTEM   (Win98)"
 Else
    List1.AddItem "Decompile failed - the .DAT file"
    List1.AddItem "may be corrupted - Checksum wrong"
 End If
 
End Sub


'This is how the DAT was made - its simple XOR encoding with XOR Checksum
Private Sub cmdMakeDAT_Click()
  
 Dim DLL() As Long, DAT() As Long, XORVal As Long, CheckSum As Long, i As Long
 
 Open App.Path & "\rvtDX9vb.dll" For Binary Access Read As #1
 Open App.Path & "\rvtDX9vb.dat" For Binary Access Write As #2
 
 ReDim DLL(0 To LOF(1) \ 4 - 1)
 Get #1, , DLL()
 ReDim DAT(0 To UBound(DLL) + 2) As Long
 
 XORVal = &H39545652        'RVT9
 CheckSum = 0
 
 For i = 0 To UBound(DLL)
   CheckSum = CheckSum Xor DLL(i)
   DAT(i + 2) = DLL(i) Xor XORVal
 Next
 DAT(0) = XORVal
 DAT(1) = CheckSum
 Put #2, , DAT()
 Close #1, #2
 
End Sub

