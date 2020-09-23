Attribute VB_Name = "ModPlaySound"
Public Declare Function PlaySound _
 Lib "winmm.dll" Alias "PlaySoundA" ( _
 ByVal lpszName As String, _
 ByVal hModule As Long, _
 ByVal dwFlags As Long) As Long
 
Public Const SND_ASYNC As Long = &H1
Public Const SND_FILENAME As Long = &H20000
Public Const SND_NOSTOP As Long = &H10



