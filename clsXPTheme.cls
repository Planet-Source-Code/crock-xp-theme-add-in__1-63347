VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXPTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' add xp components to active project.

Option Explicit

Private m_cCollection As VBComponents   ' project's component collection of forms, modules etc.
Private m_sErrReport As String          ' error report.
Private m_sFormName As String           ' current startup form's name.
Private m_sPathMod As String            ' module path.
Private m_sPathRes As String            ' resource path.
Const m_smodXP As String = "modXP"      ' internal module name, file name and also resource ID.
Const m_sresXP As String = "resXP"      ' file name and also resource ID.
Const m_sFormRef As String = "frmStartupXP" ' reference to current startup form

Public Property Get ErrReport() As String

    If Len(m_sErrReport) Then ErrReport = "Error:" & m_sErrReport
       
End Property

Public Property Get FormName() As String

    FormName = m_sFormName

End Property

Public Property Get FormRef() As String

    FormRef = m_sFormRef
    
End Property

Public Property Set ProjectXP(ByVal ActiveProject As VBProject)

    Dim cItem As VBComponent    ' component item
    Dim sError As String
    Dim sPath As String         ' short project path
    
    On Error GoTo ErrHandler
    
    ' reference components collection.
    Set m_cCollection = ActiveProject.VBComponents

    ' Ensure there is a start up form.
    If VarType(m_cCollection.StartUpObject) = vbObject Then
        Set cItem = m_cCollection.StartUpObject ' reference current start up form.
        m_sFormName = cItem.Name ' store startup form's name.
    
    Else
        sError = "No startup form found"
        GoTo ErrHandler
        
    End If
    
    ' get active project's path
    sPath = StripFile(ActiveProject.BuildFileName)
    
    ' has project been saved?
    If Len(sPath) = 0 Then
        sError = "Save project first"
        GoTo ErrHandler
        
    End If
    
    ' parse paths for xp module and xp resource.
    m_sPathMod = sPath & m_smodXP & ".bas"
    m_sPathRes = sPath & m_sresXP & ".res"
    
    ' go add components
    AddComponents
    
    Exit Property
    
ErrHandler:
    If Err.Number Then
        m_sErrReport = Err.Number & vbNewLine & Err.Description
        
    Else
        m_sErrReport = sError
        
    End If
    
End Property

Private Sub AddComponents()

    ' add xp components to active project's component collection.
    
    Dim sError As String
    Dim cItem As VBComponent
    
    On Error GoTo ErrHandler
    
    ' does project already have XP module?
    If ComponentExists(vbext_ct_StdModule, m_smodXP) Then
        sError = "Existing component detected called " & m_smodXP
        GoTo ErrHandler
    
    End If
    
    ' does project already have resource?
    If ComponentExists(vbext_ct_ResFile) Then
        sError = "Existing resource file detected.  There can be only one!"
        GoTo ErrHandler
        
    End If
    
    ' extract xp module from add-in resource.
    ResToFile m_sPathMod, m_smodXP
    
    ' extract xp resource from add-in resource.
    ResToFile m_sPathRes, m_sresXP
    
    ' continue?
    If p_bAbort Then
        sError = "Process aborted"
        GoTo ErrHandler
        
    End If
    
    ' reference and add xp module to project.
    Set cItem = m_cCollection.AddFile(m_sPathMod, False)
       
    ' add xp resource to project.
    m_cCollection.AddFile m_sPathRes, False
    
    ' go add code to xp module.
    AddCode cItem
                   
    Exit Sub
    
ErrHandler:
    If Err.Number Then
        m_sErrReport = Err.Number & vbNewLine & Err.Description
        
    Else
        m_sErrReport = sError
        
    End If
    
End Sub

Private Sub AddCode(ComponentXP As VBComponent)

    ' insert vb code into codemodule.
    ' todo check components for existing m_sFormRef (frmStartupXP) item.
    
    Dim i As Integer
    Dim sCode As String
    Dim cm As CodeModule
    Dim iLineNumber As Integer
    
    On Error GoTo ErrHandler
    
    ' reference codemodule
    Set cm = ComponentXP.CodeModule
    
    ' insert lines after last line of declaration procudure.
    iLineNumber = cm.CountOfDeclarationLines
       
    ' insert blank line
    iLineNumber = iLineNumber + 1   ' Insertion point.
    cm.InsertLines iLineNumber, ""
    
    ' insert line of code.
    sCode = "Public " & m_sFormRef & " As " & m_sFormName
    iLineNumber = iLineNumber + 1   ' Insertion point.
    cm.InsertLines iLineNumber, sCode
    
    ' insert blank line
    iLineNumber = iLineNumber + 1   ' Insertion point.
    cm.InsertLines iLineNumber, ""
        
    ' insert several lines of code (Sub Main procedure).
    sCode = "Sub Main()" & vbNewLine
    sCode = sCode & vbNewLine
    sCode = sCode & vbTab & "InitCommonControlsVB" & vbNewLine
    sCode = sCode & vbNewLine
    sCode = sCode & vbTab & "Set " & m_sFormRef & " = New " & m_sFormName
    sCode = sCode & vbNewLine
    sCode = sCode & vbTab & m_sFormRef & ".Show" & vbNewLine
    sCode = sCode & vbNewLine
    sCode = sCode & "End Sub" & vbNewLine
    
    iLineNumber = iLineNumber + 1   ' Insertion point.
    cm.InsertLines iLineNumber, sCode
    
    ' ensure start from submain
    m_cCollection.StartUpObject = vbext_so_SubMain
    
    Exit Sub
    
ErrHandler:
    m_sErrReport = Err.Number & vbNewLine & Err.Description
    
End Sub

Private Function ComponentExists(ComponentType As vbext_ComponentType, Optional ComponentName As String) As Boolean

    ' does component exists in collection?
    ' if component name is given then check name against any component.
    
    Dim c As VBComponent
        
    On Error GoTo ErrHandler
        
    For Each c In m_cCollection
        ' check type
        If c.Type = ComponentType Then
            Select Case ComponentType
                Case vbext_ct_ResFile
                    ComponentExists = True
                
                Case vbext_ct_RelatedDocument
                    ComponentExists = True
                    
                Case Else
                    If IsMissing(ComponentName) Then
                        ComponentExists = True
                        Exit For

                    Else    'check name
                        If c.Name = ComponentName Then
                            ComponentExists = True
                            Exit For

                        End If

                    End If
                 
            End Select
        End If
        
    Next
    
    Exit Function
    
ErrHandler:
    ComponentExists = False
    m_sErrReport = Err.Number & vbNewLine & Err.Description
    
End Function

Private Sub Class_Initialize()

    p_bAbort = False
    
End Sub
