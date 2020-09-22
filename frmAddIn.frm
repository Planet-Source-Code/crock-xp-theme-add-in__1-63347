VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XP Components"
   ClientHeight    =   3495
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6015
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHolder 
      BackColor       =   &H80000005&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   5715
      TabIndex        =   3
      Top             =   1800
      Width           =   5775
      Begin VB.Label lblFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "resXP.res"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   650
         Width           =   690
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   1200
         Stretch         =   -1  'True
         ToolTipText     =   "Double click to view this resource"
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "modXP.bas"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   645
         Width           =   810
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         ToolTipText     =   "Double click to view this resource"
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblCap 
      Caption         =   "lblCap"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Add XP theme components to your project?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' enumerate resource id for icons
Private Enum ridIcon
    ridModule = 101
    ridResource = 102
    ridSelModule = 103
    ridSelResource = 104
End Enum

' resource id for component byte arrays.
Const sIDModule As String = "MODXP"
Const sIDResource As String = "RESXP"

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Sub CancelButton_Click()

    Connect.Hide
    
End Sub

Private Sub Form_Load()

    Dim sBuffer As String
    
    imgIcon(0).Picture = LoadResPicture(ridModule, vbResIcon)
    imgIcon(1).Picture = LoadResPicture(ridResource, vbResIcon)
    imgIcon(0).Tag = sIDModule      ' store resource id
    imgIcon(1).Tag = sIDResource    ' store resource id
    
    sBuffer = "Enables compiled projects to adopt the XP theme." & vbNewLine
    sBuffer = sBuffer & "Requires XP operating system, but remains compatable to non XP O/S." & vbNewLine
    sBuffer = sBuffer & vbNewLine
    
    sBuffer = sBuffer & "Components" & vbNewLine
    sBuffer = sBuffer & "      1.  Standard module." & vbNewLine
    sBuffer = sBuffer & "      2.  Resource file."
       
    lblCap.Caption = sBuffer
    
End Sub

Private Sub imgIcon_DblClick(Index As Integer)

    ' Allow user to view component byte arrays before extracting.
    
    Load frmViewRes
    frmViewRes.sResID = imgIcon(Index).Tag
    frmViewRes.Icon = LoadResPicture(ridResource, vbResIcon)
    frmViewRes.Show vbModal, Me
       
End Sub

Private Sub imgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' toggle the appearance of selected resource file icon
    
    Dim iOtherIndex As Integer
    Dim iIDIcon As Integer
    Dim iIDSelIcon As Integer

    On Error Resume Next
    
    If Index = 0 Then   ' module
        iOtherIndex = 1
        iIDIcon = ridResource       ' use unselected icon
        iIDSelIcon = ridSelModule   ' use selected icon
        
    Else    ' resource
        iIDIcon = ridModule         ' use unselected icon
        iIDSelIcon = ridSelResource ' use selected icon
        
    End If
                        
    ' unhighlight controls with other index
    lblFile(iOtherIndex).BackColor = vbWindowBackground
    lblFile(iOtherIndex).ForeColor = vbWindowText
    imgIcon(iOtherIndex).Picture = LoadResPicture(iIDIcon, vbResIcon)
    
    ' highlight controls with this index
    lblFile(Index).BackColor = vbHighlight
    lblFile(Index).ForeColor = vbHighlightText
    imgIcon(Index).Picture = LoadResPicture(iIDSelIcon, vbResIcon)
    
End Sub

Private Sub OKButton_Click()
    
    Dim cXP As clsXPTheme
    Dim sTemp As String
    Dim sFormName As String
    Dim sFormRef As String
    
    On Error Resume Next
    
    Set cXP = New clsXPTheme
    
    ' pass on active project to class.
    Set cXP.ProjectXP = VBInstance.ActiveVBProject
    
    ' get current startup form's name.
    sFormName = cXP.FormName
    
    ' get the new reference to the startup form.
    sFormRef = cXP.FormRef
    
    ' get errors
    sTemp = cXP.ErrReport
    
    ' job done, release class.
    Set cXP = Nothing
     
    If Len(sTemp) Then
        MsgBox sTemp
        
    Else
        sTemp = "XP components successfuly added to project." & vbNewLine
        sTemp = sTemp & "Project's startup object is now Sub Main." & vbNewLine
        sTemp = sTemp & sFormName & " is now referenced as " & sFormRef & vbNewLine
        sTemp = sTemp & "Remember to edit the XP module if you change the name of " & sFormName
        MsgBox sTemp
        
    End If
       
End Sub


