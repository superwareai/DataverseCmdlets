dotnet clean DataverseCmdlets/DataverseCmdlets.csproj 
dotnet build DataverseCmdlets/DataverseCmdlets.csproj -c Release
nuget.exe pack -Properties "Configuration=Release;NoWarn=NU5111" ./DataverseCmdlets/DataverseCmdlets.csproj  -OutputDirectory ./