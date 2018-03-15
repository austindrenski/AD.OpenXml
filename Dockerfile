FROM microsoft/aspnetcore:2.0
WORKDIR /app
RUN Compress-Archive -LiteralPath /app -DestinationPath CompilerAPI.zip
ENTRYPOINT ['dotnet', 'CompilerAPI']
