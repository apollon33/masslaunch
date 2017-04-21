$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Welcome to Mass Launch! `n`nBy: Dustin Spears",0,"Mass Launch")

$path = Get-Location
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$file = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the file with computers:", "File Name", "$path\pcs.txt")

If(-Not $file){
    $wshell.Popup("No file was entered.",0,"Mass Launch")
    exit
}

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$site = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the website:", "Website", "$site")

If(-Not $site){
    $wshell.Popup("No site was entered.",0,"Mass Launch")
    exit
}

$domain = (Get-WmiObject Win32_ComputerSystem).Name
$pcs = Get-Content $file -Delimiter ' '

If (-Not $pcs) {
    Write-$wshell.Popup("Error! Make sure the computers names in the file are seperated by 1 space only.",0,"Mass Launch")
    exit
}

Try{
    ForEach ($pc In $pcs) {
	$pcTrimmed = $pc.Trim() + "domain"
        Invoke-Expression '& psexec \\$pcTrimmed -u domain\username -p password -i -d "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "$site"'
    } 
}
Catch{
    $wshell.Popup("Unexpected error has occured.",0,"Mass Launch")
    exit
}

$wshell.Popup("Thank you for using Mass Launch!.",0,"Mass Launch")


