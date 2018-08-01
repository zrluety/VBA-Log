Attribute VB_Name = "LogCallback"
Option Explicit

Public Const LOG_FILE = "C:\Users\gpwa\Projects\TriangleFile\MasterLog.txt"
Public Const FSO_FOR_READING As Integer = 1
Public Const FSO_FOR_WRITING As Integer = 2
Public Const FSO_FOR_APPEND As Integer = 8

' ============================================= '
' Public Methods
' ============================================= '
''
' @method LogToFile
' @param {Long} Level
' @param {String} Message
' @param {String} [From = ""]
'
' Outputs log messages to a file
''
Public Sub LogToFile(Level As Long, Message As String, Optional From As String = "")
    Dim fso As Object
    Dim log As Object
    Dim log_LevelValue As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Checking if the file exists. If the file exists we open in append mode,
    ' otherwise we create the file. In the case that the LOG_FILE path does not
    ' exist we alert that the path could not be found
    
    On Error GoTo ePathNotFound
    
    If fso.FileExists(LOG_FILE) Then
        Set log = fso.OpenTextFile(LOG_FILE, FSO_FOR_APPEND)
    Else
        Set log = fso.CreateTextFile(LOG_FILE)
    End If
    
    Select Case Level
    Case 1
        log_LevelValue = "Trace"
    Case 2
        log_LevelValue = "Debug"
    Case 3
        log_LevelValue = "Info "
    Case 4
        log_LevelValue = "WARN "
    Case 5
        log_LevelValue = "ERROR"
    End Select
    
    ' After we write the log message, we close the file and cleanup.
    log.WriteLine Now & "|" & log_LevelValue & "|" & IIf(From <> "", From & "|", "") & Message
       
    log.Close
    Set log = Nothing
    Set fso = Nothing
    
Done:
    Exit Sub
    
ePathNotFound:
    MsgBox "The directory " & Left(LOG_FILE, InStrRev(LOG_FILE, "\")) & " could not be found"
End Sub
