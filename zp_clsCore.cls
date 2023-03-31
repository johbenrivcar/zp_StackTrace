VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zp_clsCore"
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

'// ========================================================================
'// ZP VBATrace provides a neat way to track the call stack of your code, and to output
'// messages to a log that are indented in proportion to the level in the call stack.
'//
'// To use the system you simply insert a line of code at the top of a Sub or Function
'// to create the tracer object. You it a string parameter (normally the name of
'// the sub or function plus any relevant parameter values if needed) for example:
'//
'//    Public Sub Calculate()
'//      dim T: set T = zpL("Calculate")
'//     '.... code in here including calling other functions
'//    End Sub
'//
'// This causes two lines to be written to the log - one on entry to the sub or function
'// and the other on exit. The indentation of the lines is increased on entry and decreased
'// on exit, so any functions that are called by this one with the tracing included
'// will be nested and indented between them.


'// Flag indicates if tracing is turned on
Private xTracing As Boolean

'// Indent level and indent string used in log output
Private xIndentLevel As Integer
Private xIndent As String
Private xPIndent As String
Private xLastTS As String
Private xStartTS As Double

Private xLogFile As zp_clsFile


'// Log object used to generate log messages and publish
'// them through the log.msg Event.
Public log As New zp_clsLog

'// Turn tracing on
Public Sub startLog()
    xTracing = True
End Sub

'// Turn tracing off
Public Sub pauseLog()
    xTracing = False
End Sub

'// Tells the log object to debug.print all messages as well as publishing through Event MSG
Public Sub sendLogToDebug(Optional ByVal YN As Variant = Empty)
    If Not IsEmpty(YN) Then
        log.sendToDebug = YN
    End If
End Sub

Public Property Get logFile()
    
End Property


'// Creates an output text file and directs stackTrace messages into that file
Public Function sendLogToFile(Optional logFilePath As String = "") As zp_clsFile
    Set xLogFile = New zp_clsFile
    xLogFile.writeFromLog log, logFilePath
    
End Function


'// Entry function called to write the entry log message and create a tracer object.
Public Function enter(Optional label As String = "") As zp_clsT

    '// if not tracing, return nothing
    If Not xTracing Then Set enter = Nothing: Exit Function
    
    If xIndentLevel = 0 Then xStartTS = Timer: log.msgOut "LVL " & Now
    
    
    xIndentLevel = xIndentLevel + 1
    
    Dim newTS As String: newTS = Right(Now, 8)
    If newTS <> xLastTS Then
        xLastTS = newTS
        TSMessage
    End If
    
    
    
    '// Create the tracer object that is used to trigger call
    '// to zpL on exit
    If Len(label) = 0 Then label = "[" & xIndentLevel & "]"
    
    
    Dim oT As New zp_clsT
    Set oT.ZP = Me
    
    '// The exit label is always the first part of the label up to a colon:
    Dim xLabel As String: xLabel = Split(label, ":")(0)
    
    '// Set the label to be used on the exit line
    oT.label = xLabel
    oT.indentLevel = xIndentLevel
    
    xPIndent = milliSecs & xIndent
    '// Write the entry line to the log
    log.msgOut il2 & xPIndent & "|>| Enter " & label
    
    '// Add to the indent string
    xIndent = xIndent & " ."
    
    '// Return the tracer object
    Set enter = oT
End Function

'// Function to write the exit log message and decrease the indent. It is
'// called automatically by the tracer object that was instantiated by the
'// .enter function, when the object is Terminated. This happens automatically
'// when the variable referencing the object goes out of scope.
Public Sub exit_(oT As zp_clsT)
    Dim orIndent As Integer
    xIndentLevel = oT.indentLevel
    xIndent = Left(xIndent, xIndentLevel * 2 - 1)
    
    
    xPIndent = milliSecs & xIndent
    log.msgOut il2 & xPIndent & "|<| Exit " & oT.label
    xIndentLevel = xIndentLevel - 1
    If xIndentLevel = 0 Then log.msgOut "    " & Now
    
End Sub

'// Writes a timestamp message to the log at the current indent level.
Public Sub TSMessage()
    Dim tsIndent As String: tsIndent = Replace(xIndent, ".", "+")
    tsIndent = milliSecs & tsIndent
    log.msgOut il2 & tsIndent & "| TSMessage " & xLastTS
    
    
End Sub

Public Function milliSecs() As String
    milliSecs = " " & Right("    " & Round((Timer - xStartTS) * 1000, 0), 4)
End Function

'// Writes a message to the log at the current indent level.
'// Can be used in application code to write to the log while
'// preserving the indentation of the trace.
Public Sub msg(msg As String)
    xPIndent = milliSecs & xIndent
    If xTracing Then log.msgOut il2 & xPIndent & "| " & msg
End Sub


'// Utility to right-justify 2-digit number
Private Function il2() As String
    il2 = Right("  " & xIndentLevel, 2)
End Function

Private Sub Class_Initialize()
    xIndent = "|"
End Sub
