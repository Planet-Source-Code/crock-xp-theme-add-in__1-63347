VERSION 5.00
Begin VB.Form frmViewRes 
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmViewRes.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmViewRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public sResID As String

Private Sub Form_Activate()

    ' Note.  The returned string from the "ResToString" function is printed to
    ' this form, a textbox or label cannot be used to view the "resXP" resource.
    ' Is there a way to use a textbox and still show all the chars?
    
    Me.Caption = "Resource View - " & sResID
    Me.Print ResToString(sResID)
    
End Sub

Private Sub Form_Load()

    Me.AutoRedraw = True
    
End Sub



