# 前序
實情的起因是醬的，因為某次的斷電，導致我們的防毒主機死機了，又因為剛好沒有分的關係~~原因不好說~~，所以導致我們的Server只能面臨重建的命運。
<br>然而重建，卻導致Agent也要重新安裝，但問題是公司有將近`1000`台設備，分散在各樓層與其他外點(中國、荷蘭、日本)，看著同仁日漸絕望的眼神( º﹃º )與上層不要臉的壓力(對，我就是不爽你們)，我決定提出一個解決方案幫助同仁度過難關。

## Windows Remote Mangement
因為設備都有加入網域管理的關係，剛好可以透過GPO來開啟WinRM

1. 打開組策略管理控制台 ，選擇包含要啟用 WinRM 的計算機的 Active Directory 容器（組織單位），然後創建新的 GPO：corpEnableWinRM; 
![](https://woshub.com/wp-content/uploads/2022/09/enable-winrm-with-gpo.png)
![Imgur](https://imgur.com/0ZNIqy4.png)

2. 打開群組原則進行編輯 對剛新增的GPO，右鍵編輯;
<BR>轉到「電腦設定」-「原則」->「Windows 設定」->”安全設置“->”系統服務“。找到 **Windows 遠端服務 （WS-Management）** 服務併為其啟用自動啟動;
![](https://woshub.com/wp-content/uploads/2022/09/windows-remote-management-ws-management-service.png)
![Imgur](https://imgur.com/ifAXpGt.png)

3. 然後轉到「電腦設定」->「喜好設定」->「控制台設定」->「服務」。選擇「新建 -> 服務」。輸入服務名稱 WinRM，然後在「恢復」選項卡上選擇「重新啟動服務」操作;
![](https://woshub.com/wp-content/uploads/2022/09/restart-winrm-service-on-failure.png)

4. 轉到「電腦設定」->「原則」->「系統管理範本」->「Windows 元件」->「Windows 遠端管理 （WinRM）」->「WinRM 服務」→啟用「允許通過 WinRM 進行遠端伺服器管理」。在“Ipv4/IPv6 篩選器”框中，可以指定必須偵聽 WinRM 連接的IP 位址或子網。如果要允許所有IP位址上的WinRM連接，請在此處保留 ;
![](https://woshub.com/wp-content/uploads/2022/09/enable-gpo-allow-remote-server-management-throug.png)

5. 創建Windows Defender防火牆規則，允許在預設埠 TCP/5985 和 TCP/5986 上建立 WinRM 連接。轉到「電腦設定」->「原則」->「Windows設定」->「安全設置」->「具有進階安全性的 Windows 防火牆」->「具有進階安全性的 Windows 防火牆」->「輸入規則」。選擇「Windows 遠端管理預定義規則」; 
![](https://woshub.com/wp-content/uploads/2022/09/open-windows-remote-management-firewall-ports.png)

6. 轉到「電腦設定」->「原則」->「系統管理範本」->「Windows 元件」->「Windows 遠端 Shell [Windows 遠端殼層]」，然後啟用“允許遠端 Shell 訪問 [允許遠端殼層存取]。
<BR>![](https://woshub.com/wp-content/uploads/2022/09/winrm-group-policy-allow-remote-shell-access.png)

### 到目前為止，終於把GPO開完啦~ 套用到目的部門之後，就可以進入主題了
## PowerShell
開始看之前先聽我說再繼續多嘴一下(ゝ∀･)b

- 修改 Client_list.txt

※檔案須放置D槽根目錄底下

- 點擊打開 Remote_install.ps1，修改Username & Password

- 開啟PowerShell執行，或者對 Remote_install.ps1 按右鍵執行 > 用Powershell執行

- 輸入要進行遠端操作的檔名 ( 含副檔名 )

- 失敗的清單將會寫在Error.txt中


```Powershell
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
```

## 後言
我知道我不厲害，在真正的軟體工程師眼裡，這個跟屎依樣，但這確實能很好的幫我解決問題，這樣足以。
## 保持應有的競爭力，多準備一些能力跟技巧，是好的
## 就是要讓上層氣得牙癢癢又拿我沒辦法(◔౪◔)
# 共勉之 (#`皿´)
