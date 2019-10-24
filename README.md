# classic-asp-aws4signer-aws-secretmanager
This is aws authentication aws4signer for secret manager service

Straight forward setup with payload and targer property
Depend on target, payload could be empty.
```
Set service = New awsSecretManagerService
service.payload = "{""SecretId"": ""<Your SecretId>"",""VersionStage"": ""AWSCURRENT""}"
service.serviceTarget = "secretsmanager.GetSecretValue"
result = service.GetJson()
```

Url have query string, set up this line to put your query string
```
Dim canonicalRequest: canonicalRequest = "POST" & chr(10) & "/<querystring>" & chr(10) & "" & chr(10) & "host:secretsmanager.ap-southeast-2.amazonaws.com" & chr(10) & "x-amz-date:"& strNowInGMT & chr(10) & "x-amz-target:" & serviceTarget & chr(10) & "" & chr(10) & "host;x-amz-date;x-amz-target" & chr(10) & Hash(strPayLoad) 
```

Some aws service does not require to have x-amz-target in the header. So adjust the http header in <canonicalRequest>, <authorization> and <xhttp>

