Attribute VB_Name = "modFunc"
Option Explicit

Public p_bAbort As Boolean

Public Sub ResToFile(sPath As String, ResourceName As String)
 
    ' Write byte array resource to file.
    ' As there is code to be added to xp module any existing xp module must
    ' be deleted first.
    
    Dim ByteArr() As Byte
    Dim sMsg As String
    Dim iFileNum As Integer
    
    On Error GoTo ErrHandler
    
    ' continue?
    If p_bAbort Then Exit Sub
    
    ' get confirmation to kill existing file if found.
    If FileExists(sPath) Then
        sMsg = "File Exists:" & vbNewLine & sPath & vbNewLine & "Overwrite file?"
        
        If MsgBox(sMsg, vbQuestion + vbOKCancel) = vbOK Then
            Kill sPath
            
        Else
            p_bAbort = True ' abort
            Exit Sub
            
        End If
        
    End If
    
    iFileNum = FreeFile
    
    ' initialise byte array with resource
    ByteArr = LoadResData(ResourceName, "CUSTOM")

    ' write byte array to file
    Open sPath For Binary As #iFileNum
        Put #iFileNum, , ByteArr

    Close #iFileNum
    
    Exit Sub
    
ErrHandler:
    Close   ' close any open files
    MsgBox "Error:" & Err.Number & vbNewLine & Err.Description
    
End Sub

Public Function ResToString(ByVal ResourceName As String) As String

    ' Convert byte array resource to string
  
    Dim ByteArr() As Byte
    Dim sBuffer As String
    
    On Error GoTo ErrHandler

    ' initialise byte array with resource
    ByteArr = LoadResData(ResourceName, "CUSTOM")
    
    ' convert byte array to string.
    sBuffer = StrConv(ByteArr, vbUnicode)
    
    ResToString = sBuffer
     
    Exit Function
    
ErrHandler:
    MsgBox "Error:" & Err.Number & vbNewLine & Err.Description
    
End Function

Public Function StripFile(ByVal sPath As String) As String
    
    ' Strip the file name and return with the path.
    
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    For i = Len(sPath) To 1 Step -1
        If Mid$(sPath, i, 1) = "\" Then
            StripFile = Left$(sPath, i)
            Exit For
            
        End If
        
    Next
    
ErrHandler:
   
End Function

Public Function FileExists(ByVal sPath) As Boolean

    On Error GoTo ErrHandler
    
    If FileLen(sPath) >= 0 Then FileExists = True
     
    Exit Function
       
ErrHandler:
    FileExists = False
    
End Function
