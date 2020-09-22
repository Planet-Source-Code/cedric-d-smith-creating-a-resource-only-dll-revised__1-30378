VERSION 5.00
Begin VB.Form frmRes 
   Caption         =   "Resouce Viewer"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Close Viewer"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   3600
      Picture         =   "frmRes.frx":0000
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enter Image#"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter Wav#"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Image"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "frmRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bytsound() As Byte

Dim TheResource As New ResClass
'Cedric D. Smith(c) 2002
'Purpose:  This code accesses the dll you created with Createthedll.vbp
'Note:   If you decided to use my sample resources, their id's are (bmp)101 and (wav)101)
' Make sure that you have registered your dll in the IDE before executing

'NOTE ABOUT AVI's:
'The way that I handle avi's is to download them from the dll file
'and save them to the hard drive(saving as a binary file) and load.  I am still researching how
'to load them direct from the dll, place in a picture file and run.
'If you figure out, send to Ibolink@hotmail.com


Private Sub Command1_Click()
Dim i As Integer

i = Val(Text2.Text)

PlayWaveRes i, SND_ASYNC

End Sub

Private Sub Command2_Click()
'
' Instantiate the ResClass class as New

' Set the Picture of the Picture1 object to the function Return
' Object

Dim i As Integer
i = Val(Text1.Text)

  Set Picture1.Picture = TheResource.GetPicture(i, 0)

End Sub

Private Sub Command3_Click()
End
End Sub

Public Sub PlayWaveRes(ResourceID As Integer, Optional vntFlags)
              
       
      bytsound = TheResource.GetWAVE(ResourceID, "CUSTOM")
       

       If IsMissing(vntFlags) Then
         vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
         
       End If

       If (vntFlags And SND_MEMORY) = 0 Then
          vntFlags = vntFlags Or SND_MEMORY
       End If

      sndsound = sndPlaySound(bytsound(0), vntFlags)
           
 
     End Sub

