# Convert2PDF

A CLI utility to convert Word, Excel and Powerpoint files to PDF

## How to use
- Open up Powershell (tested on v5 to v7)
- Run this command to grant Powershell scripts executable permission - `powershell.exe -ExecutionPolicy Bypass -File ".\convert2pdf.ps1"`
- Navigate to the directory the `convert2pdf.ps1` file is located in
- `.\convert2pdf.ps1 "Path"`

## Supported File Formats
- docx
- pptx
- xlsx

## Prerequisites
- Powershell 5+
- Microsoft Word, Powerpoint and Excel

## Running Globally
- Add the file to PATH (System Properties → Advanced → Environment Variables)
- Set powershell execution policy (`Set-ExecutionPolicy RemoteSigned -Scope LocalMachine`)
  - Reason for this is that powershell restricts script execution by default
- Now no need to navigate to the directory with the script file, you can directly run `convert2pdf.ps1 "Path"`

### Roadmap
- [ ] Executable .exe file
- [ ] More file formats to support
