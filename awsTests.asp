
<!-- #Include file="awsSecretManagerService.asp"  -->
<%


Set service = New awsSecretManagerService
    service.payload = "{""SecretId"": ""<Your SecretId>"",""VersionStage"": ""AWSCURRENT""}"
    service.serviceTarget = "secretsmanager.GetSecretValue"
result = service.GetJson()


response.write(result)

%>
