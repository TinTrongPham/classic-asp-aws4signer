<%

Const aws_accessKey     = "YOUR ACCESS KEY"
Const aws_secretKey     = "YOUR PRIVATE KEY"
Const aws_waitTimeout   = 28
Const aws_algorithm     = "AWS4-HMAC-SHA256"
Const aws_region        = "ap-southeast-2"
Const aws_serviceName   = "secretsmanager"
'=============================================================================='
'=============================================================================='
Class awsSecretManagerService
    private strServiceTarget 
    private strPayLoad 

    Public Property Get serviceTarget() 
       serviceTarget = strServiceTarget
    End Property

    Public Property Get payload() 
       payload = strPayLoad
    End Property

    Public Property Let serviceTarget(ByVal NewValue) 
       strServiceTarget = NewValue
    End Property
    
    Public Property Let payload(ByVal NewValue) 
       strPayLoad = NewValue
    End Property

    '-- NowInGMT ------------------------------------------------------------------'    
    ' return datetime to yyyyMMddThhmmssZ
    Private Function NowInGMT(now)      
        Dim date: date = datepart("yyyy", now)
        date = date & RIGHT("0" & datepart("m",now),2)
        date = date & RIGHT("0" & datepart("d",now),2)
        date = date & "T"
        date = date & RIGHT("0" & datepart("h",now),2)    
        date = date & RIGHT("0" & datepart("n",now),2)
        date = date & RIGHT("0" & datepart("s",now),2)
        date = date & "Z"
        NowInGMT = date
        Set date= Nothing
        Set dtNowGMT= Nothing
        Set iOffset= Nothing
    End Function

    '-- Hash -----------------------------------------------------------'
    Private Function Hash(data) 
        Dim BytesToHashedBytes: BytesToHashedBytes = Encrypt("sha256","", data, False) 
        For x = 1 To LenB(BytesToHashedBytes)
            HashedBytesToHex = HashedBytesToHex & Right("0" & Hex(AscB(MidB(BytesToHashedBytes, x, 1))), 2)
        Next
        Hash = LCase(HashedBytesToHex)
        Set BytesToHashedBytes = Nothing
    End Function

    '-- Encrypt -----------------------------------------------------------'
    Private Function Encrypt(HashType, key, data, withKey)
        
        ' create UTF-8 string
        With CreateObject("ADODB.Stream")
            .Open
            .CharSet = "Windows-1252"
            .WriteText data
            .Position = 0
            .CharSet = "UTF-8"
            data = .ReadText
            .Close
        End With

        Set UTF8Encoding = CreateObject("System.Text.UTF8Encoding")

        Dim PlainTextToBytes, BytesToHashedBytes, HashedBytesToHex
        PlainTextToBytes = UTF8Encoding.GetBytes_4(data)
    
        Select Case HashType
            Case "md5": Set Cryptography = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider") '< 64 (collisions found)
            Case "ripemd160": Set Cryptography = CreateObject("System.Security.Cryptography.RIPEMD160Managed")
            Case "sha1": Set Cryptography = CreateObject("System.Security.Cryptography.SHA1Managed") '< 80 (collision found)
            Case "sha256": Set Cryptography = CreateObject("System.Security.Cryptography.SHA256Managed")
            Case "sha384": Set Cryptography = CreateObject("System.Security.Cryptography.SHA384Managed")
            Case "sha512": Set Cryptography = CreateObject("System.Security.Cryptography.SHA512Managed")
            Case "md5HMAC": Set Cryptography = CreateObject("System.Security.Cryptography.HMACMD5")
            Case "ripemd160HMAC": Set Cryptography = CreateObject("System.Security.Cryptography.HMACRIPEMD160")
            Case "sha1HMAC": Set Cryptography = CreateObject("System.Security.Cryptography.HMACSHA1")
            Case "sha256HMAC": Set Cryptography = CreateObject("System.Security.Cryptography.HMACSHA256")
        End Select

        Cryptography.Initialize()    
        if withKey = True Then Cryptography.Key = key End If
        BytesToHashedBytes = Cryptography.ComputeHash_2((PlainTextToBytes))    
        Encrypt = BytesToHashedBytes

        Set PlainTextToBytes = Nothing 
        Set BytesToHashedBytes = Nothing
        Set HashedBytesToHex = Nothing
    End Function

    '-- ToHex ----------------------------------------------------------'
    Private Function ToHex(BytesToHashedBytes)
        Dim HashedBytesToHex
        For x = 1 To LenB(BytesToHashedBytes)
            HashedBytesToHex = HashedBytesToHex & Right("0" & Hex(AscB(MidB(BytesToHashedBytes, x, 1))), 2)
        Next
        ToHex = LCase(HashedBytesToHex)
        Set HashedBytesToHex = Nothing
    End Function

    '-- Sign -----------------------------------------------------------'
    Private Function Sign(key, dateStamp, regionName, serviceName)
        Dim UTF8Encoding : Set UTF8Encoding = CreateObject("System.Text.UTF8Encoding")
        Dim kDate: kDate = Hmac(dateStamp, UTF8Encoding.GetBytes_4("AWS4" & key))   
        Dim kRegion: kRegion = Hmac(regionName, kDate)    
        Dim kService: kService = Hmac(serviceName, kRegion)    
        Sign = Hmac("aws4_request", kService)         
        Set kDate = Nothing
        Set kRegion = Nothing
        Set kService = Nothing
        Set UTF8Encoding = Nothing
    End Function

    '-- HMAC ---------------------------------------------------------------------'
    Private Function Hmac(data, key)      
        Hmac = Encrypt("sha256HMAC", key, data, True)    
    End Function

    '-- GetJson ----------------------------------------------------------------'
    Function GetJson() 

        '-- Authentication: --'
        ' make sure dtNowGMT is UTC
        Dim sh: Set sh = Server.CreateObject("WScript.Shell")
        Dim iOffset: iOffset = sh.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
        Dim dtNowGMT: dtNowGMT = DateAdd("n", iOffset, Now())
        Dim strNowInGMT: strNowInGMT = NowInGMT(dtNowGMT)        
        Dim strNowDateOnly: strNowDateOnly = Year(dtNowGMT) & Month(dtNowGMT) & Day(dtNowGMT)    
        Dim credentialScope: credentialScope = strNowDateOnly & "/" & aws_region & "/" & aws_serviceName & "/aws4_request"
        Dim canonicalRequest: canonicalRequest = "POST" & chr(10) & "/" & chr(10) & "" & chr(10) & "host:secretsmanager.ap-southeast-2.amazonaws.com" & chr(10) & "x-amz-date:"& strNowInGMT & chr(10) & "x-amz-target:" & serviceTarget & chr(10) & "" & chr(10) & "host;x-amz-date;x-amz-target" & chr(10) & Hash(strPayLoad)  
        Dim hashedCanonicalRequest: hashedCanonicalRequest = Hash(canonicalRequest)
        Dim stringToSign: stringToSign = _  
            aws_algorithm & chr(10)  & _
            strNowInGMT & chr(10)  & _
            credentialScope & chr(10)  & _
            hashedCanonicalRequest      
        Dim signingKey: signingKey = Sign(aws_secretKey, strNowDateOnly, aws_region, aws_serviceName)             
        Dim signature: signature = ToHex(Hmac(stringToSign, signingKey))        
        Dim authorization: authorization = aws_algorithm & " Credential=" & aws_accessKey & "/" & credentialScope & ",SignedHeaders=host;x-amz-date;x-amz-target, Signature=" & signature
    
        '-- GetValue: --'
        Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
        xhttp.open "POST", "https://secretsmanager.ap-southeast-2.amazonaws.com", TRUE
        xhttp.setRequestHeader "Content-Type", "application/x-amz-json-1.1"
        xhttp.setRequestHeader "X-Amz-Date", strNowInGMT 
        xhttp.setRequestHeader "Authorization", authorization
        xhttp.setRequestHeader "X-Amz-Target", serviceTarget         
        xhttp.send (strPayLoad)
    
        If xhttp.waitForResponse(aws_waitTimeout) THEN       
            If xhttp.status = "200" Then
                GetJson = xhttp.responseText
            End If
        End If
    End Function
End Class
'=============================================================================='
'=============================================================================='

%>
