<%
'在ASP中实现加密与解密,加密方法:根据RSA
'作者:hnsoso@sina.com
Class clsRSA
    
    Public PrivateKey
    Public PublicKey
    Public Modulus
    
    Public Function Crypt(pLngMessage, pLngKey)
        On Error Resume Next
        Dim lLngMod
        Dim lLngResult
        Dim lLngIndex
        If pLngKey Mod 2 = 0 Then
            lLngResult = 1
            For lLngIndex = 1 To pLngKey / 2
                lLngMod = (pLngMessage ^ 2) Mod Modulus
                ' Mod may error on key generation
                lLngResult = (lLngMod * lLngResult) Mod Modulus
                If Err Then Exit Function
            Next
        Else
            lLngResult = pLngMessage
            For lLngIndex = 1 To pLngKey / 2
                lLngMod = (pLngMessage ^ 2) Mod Modulus
                On Error Resume Next
                ' Mod may error on key generation
                lLngResult = (lLngMod * lLngResult) Mod Modulus
                If Err Then Exit Function
            Next
        End If
        Crypt = lLngResult
    End Function
    
    Public Function Encode(ByVal pStrMessage)
        Dim lLngIndex
        Dim lLngMaxIndex
        Dim lBytAscii
        Dim lLngEncrypted
        lLngMaxIndex = Len(pStrMessage)
        If lLngMaxIndex = 0 Then Exit Function
        For lLngIndex = 1 To lLngMaxIndex
            lBytAscii = Asc(Mid(pStrMessage, lLngIndex, 1))
            lLngEncrypted = Crypt(lBytAscii, PublicKey)
            Encode = Encode & NumberToHex(lLngEncrypted, 4)
        Next
    End Function
    
    Public Function Decode(ByVal pStrMessage) 
        Dim lBytAscii
        Dim lLngIndex
        Dim lLngMaxIndex
        Dim lLngEncryptedData
        Decode = ""
        lLngMaxIndex = Len(pStrMessage)
        For lLngIndex = 1 To lLngMaxIndex Step 4
            lLngEncryptedData = HexToNumber(Mid(pStrMessage, lLngIndex, 4))
            lBytAscii = Crypt(lLngEncryptedData, PrivateKey)
            Decode = Decode & Chr(lBytAscii)
        Next
    End Function
    
    Private Function NumberToHex(ByRef pLngNumber, ByRef pLngLength)
        NumberToHex = Right(String(pLngLength, "0") & Hex(pLngNumber), pLngLength)
    End Function
    
    Private Function HexToNumber(ByRef pStrHex)
        HexToNumber = CLng("&h" & pStrHex)
    End Function
End Class
%>
<%
function Encryptstr(Message)
    Dim LngKeyE
    Dim LngKeyD
    Dim LngKeyN
    Dim StrMessage
    Dim ObjRSA
    
    LngKeyE = "32823"
    LngKeyD = "20643"
    LngKeyN = "29893"
    StrMessage = Message
    
    Set ObjRSA = New clsRSA
    ObjRSA.PublicKey = LngKeyE
    ObjRSA.Modulus = LngKeyN
    Encryptstr = ObjRSA.Encode(StrMessage)
    Set ObjRSA = Nothing
end function
 
function decryptstr(Message)
    Dim LngKeyE
    Dim LngKeyD
    Dim LngKeyN
    Dim StrMessage
    Dim ObjRSA
    
    LngKeyE = "32823"
    LngKeyD = "20643"
    LngKeyN = "29893"
    StrMessage = Message
    
    Set ObjRSA = New clsRSA
    ObjRSA.PrivateKey =LngKeyD
    ObjRSA.Modulus=LngKeyN
    decryptstr=ObjRSA.Decode(StrMessage)
    Set ObjRSA = Nothing
end function
 
dim last,first
first = "cftea"
Response.Write "加密前为:"&first
last=Encryptstr(first)
Response.Write "加密后为"&last
Response.Write "解密后为" &decryptstr(last)
%>