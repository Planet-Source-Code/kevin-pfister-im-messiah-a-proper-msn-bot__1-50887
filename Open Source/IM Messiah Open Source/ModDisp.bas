Attribute VB_Name = "ModDisp"
Public MSNObject As String
Public DisplayPicData As String

Public Function CreateMSNObject()
    Open App.Path & "\data\Avatar.PNG" For Binary As #1
        DisplayPicData = Input(LOF(1), #1)
    Close #1
    
    Dim Obj As String
    Obj = "<msnobj Creator=""" & frmMain.strUserName & """ " & _
          "Size=""" & Len(DisplayPicData) & """ " & _
          "Type=""3"" Location=""dphowto.tmp"" Friendly=""AAA="" " & _
          "SHA1D=""" & Base64Encode(HexToBin(SHAHash(DisplayPicData))) & """ "
    
    'Create the SHA1C hash
    Dim SHA1C As String, SHAArray() As String, TSha As String, I As Long
    SHAArray = Split(Trim$(Obj), " "): SHAArray(0) = ""
    For I = 1 To UBound(SHAArray) ' - 1
        SHAArray(I) = GetBetween(SHAArray(I), , "=") & GetBetween(SHAArray(I), "=""", """")
    Next I
    
    'Finish object
    SHA1C = Join$(SHAArray, "")
    SHA1C = Base64Encode(HexToBin(SHAHash(SHA1C)))
    MSNObject = Obj & "SHA1C=""" & SHA1C & """/>"
End Function
