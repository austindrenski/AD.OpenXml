# build the app
FROM microsoft/aspnetcore-build:2.0
WORKDIR /app
COPY . ./
RUN dotnet publish src/CompilerAPI -c Release -f netcoreapp2.0 -r rhel-x64 -o /app/out

# build runtime image
FROM microsoft/aspnetcore:2.0
WORKDIR /app/out
ENTRYPOINT CompilerAPI.exe
