FROM microsoft/aspnetcore:2.0
WORKDIR /app
RUN Compress-Archive -LiteralPath /app -DestinationPath /app.zip
ENTRYPOINT ['dotnet', 'CompilerAPI']
