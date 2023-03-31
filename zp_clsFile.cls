VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zp_clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'// ========================================================================
'// MIT License
'//
'// Copyright (c) 2023 John Rivett-Carnac +44 7887 570 669
'//
'// Permission is hereby granted, free of charge, to any person obtaining a copy
'// of this software and associated documentation files (the "Software"), to deal
'// in the Software without restriction, including without limitation the rights
'// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'// copies of the Software, and to permit persons to whom the Software is
'// furnished to do so, subject to the following conditions:
'//
'// The above copyright notice and this permission notice shall be included in all
'// copies or substantial portions of the Software.
'//
'// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'// SOFTWARE.
'// ========================================================================

'// Provides the means to write the trace log to a text file. The default location
'// is the path <excel file path>\stacktrace\trace yyyy-mm-dd+hh-mm.txt


Private xTraceFile As TextStream
Private WithEvents xLog As zp_clsLog
Attribute xLog.VB_VarHelpID = -1


Dim fso As New Scripting.FileSystemObject

Public Sub writeFromLog(ByRef log As zp_clsLog, Optional sFilepath As String = "")
    If Not xTraceFile Is Nothing Then xTraceFile.Close
    
    Set xLog = log
    
    If Len(sFilepath) = 0 Then sFilepath = defaultTraceFullName
    
    Set xTraceFile = fso.OpenTextFile(sFilepath, ForAppending, True)
    
End Sub


Public Function defaultTraceFullName() As String
    Dim sPath As String
    sPath = ThisWorkbook.Path
    
    If Not fso.FolderExists(sPath & "\logs") Then
        fso.CreateFolder sPath & "\logs"
    End If
    
    sPath = sPath & "\logs\stacktrace " & sDTS & ".txt"
    
    
    defaultTraceFullName = sPath
End Function


Public Function sDTS() As String
    Dim ss As String: ss = Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "-")
    Dim vv: vv = Split(ss, "-")
    sDTS = vv(2) & "-" & vv(1) & "-" & vv(0) & "+" & vv(3) & "-" & vv(4)
End Function


Private Sub Class_Terminate()
    Set xLog = Nothing
    
    If Not xTraceFile Is Nothing Then xTraceFile.Close
    
End Sub


Private Sub xLog_msg(pmsg As String)
    xTraceFile.WriteLine pmsg
End Sub
