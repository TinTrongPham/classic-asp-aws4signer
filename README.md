# classic-asp-aws4signer-aws-secretmanager


Straight forward setup with payload and targer property
Depend on target, payload could be empty.
```
Set service = New awsSecretManagerService
service.payload = "{""SecretId"": ""<Your SecretId>"",""VersionStage"": ""AWSCURRENT""}"
service.serviceTarget = "secretsmanager.GetSecretValue"
result = service.GetJson()
```
