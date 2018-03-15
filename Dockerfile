FROM microsoft/aspnetcore:2.0
WORKDIR /app/rhel-x64
EXPOSE 5000/tcp
ENTRYPOINT ['CompilerAPI']
