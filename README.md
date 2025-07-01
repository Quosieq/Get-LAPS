This was just a project using windows forms for Windows LAPS since default GUI works only with Microsoft LAPS 

Both normal and something noone has asked for - Dark Mode (DM with slightly better UI) 

Basic Functions: 
1. auto copy to clipboard after submitting PC name
2. basic information output if PC is in correct OU (can be removed from ps1 file)/missing from AD

To use those as EXE files just pack tem using IExpress 
Install program custom commadn for semi-classic UI:

> powershell.exe -ExecutionPolicy Bypass -File "Get-Laps.ps1"

For DarkMode: 
> powershell.exe -STA -ExecutionPolicy Bypass -File "Get-Laps_DM.ps1"
