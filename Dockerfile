FROM microsoft/aspnetcore:2.0
WORKDIR /app
ENTRYPOINT ['dotnet', 'CompilerAPI']
