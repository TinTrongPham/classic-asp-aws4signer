# classic-asp-aws4signer
Create http client request to AWS secret manager since there is no AWS sdk for classic asp

Straight forward setup with payload and targer property
Depend on target, payload could be empty.
```
Set service = New awsSecretManagerService
service.payload = "{""SecretId"": ""<Your SecretId>"",""VersionStage"": ""AWSCURRENT""}"
service.serviceTarget = "secretsmanager.GetSecretValue"
result = service.GetJson()
```
