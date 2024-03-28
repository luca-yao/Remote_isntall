<#
文件：Remote_install.ps1
用途：自動化複製、安裝、刪除
創建：2024-03-29
創建人：Luca Yao
#>

$Username = 'domain\administrator'
$Password = 'password'
$pass = ConvertTo-SecureString -AsPlainText $Password -Force
$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass

$Filename = Read-Host -Prompt "Enter FileName"
Remove-Item Error.txt

foreach($serverName in Get-Content Client_list.txt) {
    if($serverName -match $regex){
        echo $serverName
        try { 
            $Session = New-PSSession -ComputerName $serverName -Credential $Cred

            try {
                echo "Copy File"
                Copy-Item -Path "D:\$Filename" -Destination "C:\$Filename"  -ToSession $Session
                echo $Filename
                echo "install & remove"
                Invoke-Command -ComputerName $serverName -ScriptBlock { param($Filename)
                    Start-Process -FilePath "C:\$Filename" -Wait 
                    echo $Filename
                    Remove-Item C:\$Filename
                } -ArgumentList $Filename -credential $Cred
        } finally {
            Remove-PSSession -Session $Session
        }
     } catch {
            "Error occurred on server: $serverName" | Out-File -FilePath "Error.txt" -Append
        }
    }
    
}



