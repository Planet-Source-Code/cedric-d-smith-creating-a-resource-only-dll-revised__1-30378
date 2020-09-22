Attribute VB_Name = "Module1"
Option Explicit
       
       Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
       Global sndsound As Long
       
       
       ' Flag values for wFlags parameter.
   
 '*********************************************************************

       Global Const SND_SYNC = &H0        ' Play synchronously (default).
       Global Const SND_ASYNC = &H1      ' Play asynchronously (see
                                         ' note below).
       Global Const SND_NODEFAULT = &H2   ' Do not use default sound.
       Global Const SND_MEMORY = &H4      ' lpszSoundName points to a
                                          ' memory file.
      Global Const SND_LOOP = &H8        ' Loop the sound until next
                                          ' sndPlaySound.
      Global Const SND_NOSTOP = &H10     ' Do not stop any currently
                                          ' playing sound.

'*********************************************************************

       ' Plays a wave file from a resource.

 '*********************************************************************

       




