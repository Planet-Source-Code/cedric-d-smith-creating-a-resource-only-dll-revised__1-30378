VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Cedric D. Smith (c) 2002
'Purpose: This code demonstrates how to create a resource-only dll to store
'bmps, wavs etc and load them from a VB application.

'Credits: The structure of this code was provided by FATBOYPEET from VBworld.
'However in response to some questions as to how to use the code, I made some
'modifications and added a function that loads the binary data.

'What it does:
'This code demonstrates how to load a picture and a wav from an ActiveX dll.
'

'What to do with this code:
'1. Simply add your resources to Createthedll.res(resource file)
'   (Note**: I added a bmp and wav to it already, you can compile and run or
'    delete and add your own.)
'2. Compile the dll.

'3. Now register the dll in the IDE by going to PROJECT--->>REFERENCES. Open
'   the new dll you just created.

'4.  Close this app and open ViewRes.Vbp to view your resources.


'This function was modified by Cedric D. Smith

Public Function GetPicture(PicIdnum As Integer, ipictype As Integer) As Object

' This code is used to allow you to return a picture as object
  Set GetPicture = LoadResPicture(PicIdnum, ipictype)

End Function

'This function was added by Cedric D. Smith
Public Function GetWAVE(iwavetype As Integer, status As String) As Variant
               
        
    GetWAVE = LoadResData(iwavetype, status) ' passes binary data to GetWave
                                             'function.  Can used to pass an
                                             'avi as well
       
End Function

