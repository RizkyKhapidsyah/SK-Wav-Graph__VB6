Attribute VB_Name = "ModWav"
Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long


Global DIFF
Global mssg As String * 255
Global TimeChunk As String
