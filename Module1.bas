Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
'always start your program with a Sub Main as part of good
'coding practice.

Form2.Show
DoEvents 'Isn't necessary anymore for the buttons to draw
         'right, but does force form1 in front of form2.
Form1.Show

End Sub
