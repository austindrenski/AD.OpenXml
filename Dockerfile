FROM microsoft/aspnetcore:2.0
WORKDIR /app
RUN 7za compilerapi.zip /app
ENTRYPOINT ['dotnet', 'CompilerAPI']
