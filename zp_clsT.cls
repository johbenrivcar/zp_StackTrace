VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zp_clsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

'// Implements the tracer object that is used to call the outdenting function
'// when the object is terminated. This makes a call to the outdent happen
'// automatically, when the variable that references the object goes out of
'// scope, i.e. on exit from the containing function.

'// The indent level on creation of this instance
Public indentLevel As Integer
'// The user-assigned label of this instance
'// Normally the name of the sub or function where instantiated
Public label As String
'// Instance of zp_Core that owns this Trace instance
Public ZP As zp_clsCore

'// Termination function called by the system when the object goes out of scope
Private Sub Class_Terminate()
    ZP.exit_ Me
End Sub

