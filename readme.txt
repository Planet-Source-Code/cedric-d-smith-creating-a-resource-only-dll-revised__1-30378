In response to a many VB programmers wanting a better
alternative to using resource files to store their
bmps,wavs, etc.  This code presents a simple way
to store those resources in a resource-only ActiveX 
dll that can be loaded from their VB application. 

Here is some info about the code:


ViewRes.Vbp      - This the code that loads the resources from
                   the dll. To run this code, you must FIRST load
                   Createthedll.vbp and modify the cREATETHEDLL.RES
                   by adding your resources**. Compile
                   into dll and register it in the IDE going to 
                   PROJECTS-->REFERENCES and loading the dll.   
                                


Createthedll.vbp - This is what you will use to create the
                   dll that will hold your resources. I have
                   created a .res(Createthedll.res) that has
                   a picture and a wave file in it. You can
                   delete those files and add your own**. Once
                   you have finished, compile to dll.



Steps:
1. Open Createthedll.vbp and add resources to .res and compile to dll.
2. Register the dll in the IDE Projects--> References
3. Open ViewRes.Vbp and run to view your resources.




(**Note: I have placed a bmp and wav in the .res file 
   already. You can just compile it as a demonstration). 


I am sure there a lot improvements that can be made, bear in mind 
this was put together to respond to user questions about using VB to
store their wavs, bmps, etc. in a resource only dll.

I would have included a complied dll, but Planet-Source doesn't allow the uploading
of dll's..

Cedric D. Smith
ibolink@hotmail.com
(c)2002