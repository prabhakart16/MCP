REM Navigate to project
cd C:\Prabhakar\MCP

REM Build server
cd ExcelMcpServer
dotnet build -c Release
dotnet publish -c Release -o ./publish

REM Build client
cd ..\ExcelMcpClient
dotnet build -c Release

REM Run interactive client
dotnet run -- ..\ExcelMcpServer\publish\ExcelMcpServer.exe ..\sampleFile\LoanReconciliationSample.xlsx