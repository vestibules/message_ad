$StartTime = (Get-Date).ToShortDateString()+", "+(Get-Date).ToLongTimeString()
Import-Module ActiveDirectory -ErrorAction Stop 

$testpath = Test-Path -Path .\log_envoi_message_ad.txt
if ($testpath -match "True")
{
Remove-Item .\log_envoi_message_ad.txt
}

$scope = Read-Host "Envoyer le message à tout le domaine [taper 1] à une UO [taper 2], ou à un Poste [taper 3] ?"

switch ($scope) 
    {

    "2" {
            $ou = Read-Host "Entrez le chemin LDAP de destination du message (format : CN=,DC=,DC= ou OU=,DC=,DC=)"
            $comp = Get-ADComputer -SearchBase $ou -Filter {OperatingSystem -like '*Windows*'} | Select-Object -ExpandProperty name
        }

    "1" {   
            $comp = Get-ADComputer -Filter {OperatingSystem -like '*Windows*'} | Select-Object -ExpandProperty name
        }

    "3" {
            $comp = Read-Host "Entrez le nom du poste"     
        }
    }

$msg = Read-Host 'Entrez le message'

foreach ($computer in $comp) 
    {
        $test = Test-Connection -CN $computer -Count 1 -BufferSize 16 -Quiet

            if ($test -match 'True') { 
                                        Write-Output "Message envoyé : $computer" | Tee-Object -filepath .\log_envoi_message_ad.txt -Append
                                        Invoke-WmiMethod  -Path Win32_Process -Name Create -ArgumentList "msg * $msg" -ComputerName $computer
                                     }

            else {Write-Output "$Computer est hors ligne" | Tee-Object -filepath .\log_envoi_message_ad.txt -Append}

        



    }



$EndTime = (Get-Date).ToShortDateString()+", "+(Get-Date).ToLongTimeString()
$TimeTaken = New-TimeSpan -Start $StartTime -End $EndTime

Write-Output ($Footer = @"

$("-"*79)
   Heure de début : $StartTime
   Heure de fin : $EndTime
   Durée totale : $TimeTaken
$("-"*79)
"@)


Out-File ./log_envoi_message_ad.txt -Append -InputObject $Footer
$StartTime,$EndTime,$TimeTaken=$null
[void](Read-Host 'Appuyer sur Entrée pour quitter le script.')

