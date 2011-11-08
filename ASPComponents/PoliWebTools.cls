VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PoliWebTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function Copy(strSource As String, strTarget As String, nType As Integer) As Integer
    Dim Buffer As String
    Open strSource For Binary As #1
    Open strTarget For Binary As #2
    Buffer = String(1024, " ")
    
    Do
        Get #1, , Buffer
        Put #2, , Buffer
    Loop While Not EOF(1)
    
    Close 1, 2
End Function

Function GetImage(strSource As String) As String
    Dim Buffer As String
    Dim Result As String
    
    Open strSource For Binary As #1
    Buffer = String(1024, " ")
    Result = ""
    
    Do
        Get #1, , Buffer
        Result = Result & Buffer
        
    Loop While Not EOF(1)
    
    Close 1, 2
    
    GetImage = Result
    
End Function


