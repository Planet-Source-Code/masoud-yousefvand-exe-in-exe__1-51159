VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Code by Dr. Y

'This is an example of how you can put an EXE within another and retreive it.
'In this example we put 'MyEXE.exe' inside 'Project1.exe' and retreive it at runtime.
'I added some images within ZIP folder you downloaded, see them.

'You must first add 'Resource editor' to your VB IDE toolbar from 'Add-Ins' menu. [image1]
'then make the first EXE and complie it and save somewhere(in this example 'MyEXE.exe').
'Now open resource editor and click on 'Add custom resource'. [image2]
'Select your EXE and click on 'Open'.
'Now click on 'Save'. [image3]
'Your project must look like [image4].

'OK, if you compile the project now that EXE will be inside this one.
'For retriving that EXE see below:

Private Sub Form_Load()

    Dim Buffer() As Byte
    
    Buffer = LoadResData(101, "CUSTOM")
    
    Open "c:\MyEXE.exe" For Binary As #1 'Don't forget to delete this file from your drive later !
    Put #1, , Buffer()
    Close #1
    Shell "c:\MyEXE.exe", vbNormalFocus
    Me.WindowState = vbMinimized ' let's see our EXE on screen!

End Sub
