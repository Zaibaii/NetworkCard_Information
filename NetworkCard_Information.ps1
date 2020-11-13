function global:_Pause($message)
{
    #Check if script is running by Powershell ISE
    if ($psISE)
    {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("$message")
    }
    else
    {
        Write-Host "$message" -ForegroundColor Yellow
        $x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

function global:_TradENtoFR($string)
{
    $string = $string -replace "True", "Oui"
    $string = $string -replace "False", "Non"
    return $string
}

function global:_Ipv4Only($ip)
{
    return ($ip -split(" "))[0]
}

function global:_DateFR($date)
{
    return $date = $date -replace "^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2}).*", '$3/$2/$1 $4:$5:$6'
}

function global:_TradStatus($number)
{
    switch($number) {
        0 {"D�connect�"}
        1 {"Connexion en cours"}
        2 {"Connect�"}
        3 {"D�connexion en cours"}
        4 {"Carte d�sactiv�"}
        5 {"Mat�riel d�sactiv�"}
        6 {"Dysfonctionnement mat�riel"}
        7 {"M�dia d�connect�"}
        8 {"Authentification"}
        9 {"Authentification r�ussie"}
        10 {"Authentification �chou�e"}
        11 {"Adresse invalide"}
        12 {"Informations requises"}
        default {"N/A"}
    }
}

$ListNetworkAdapter = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object -FilterScript {$_.netconnectionid -ne $null}
ForEach ($NetworkAdapter in $ListNetworkAdapter) {
    $NetworkAdapterConfiguration = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Select-Object -Index $NetworkAdapter.Index
	$Result = [Ordered]@{
		"Nom de la connexion" = ": $($NetworkAdapter.NetConnectionID)";
		"Fabriquant" = ": $($NetworkAdapter.Manufacturer)";
		"Mod�le" = ": $($NetworkAdapter.Name)";
		"Carte physique" =": $(_TradENtoFR($NetworkAdapter.PhysicalAdapter))";
		"Etat de la connexion" = ": $(_TradStatus($NetworkAdapter.NetConnectionStatus))";
		"Adresse MAC" = ": $($NetworkAdapter.MACAddress)";
		"Type de connexion" = ": $($NetworkAdapter.AdapterType)";
		"DHCP activ�" = ": $(_TradENtoFR($NetworkAdapterConfiguration.DHCPEnabled))";
		"Serveur DHCP" = ": $($NetworkAdapterConfiguration.DHCPServer)";
		"Bail DHCP - Obtenu" = ": $(_DateFR($NetworkAdapterConfiguration.DHCPLeaseObtained))";
		"Bail DHCP - Expiration" = ": $(_DateFR($NetworkAdapterConfiguration.DHCPLeaseExpires))";
		"Adresse IP" = ": $(_Ipv4Only($NetworkAdapterConfiguration.IPAddress))";
		"Masque" = ": $(_Ipv4Only($NetworkAdapterConfiguration.IPSubnet))";
		"Passerelle" = ": $(_Ipv4Only($NetworkAdapterConfiguration.DefaultIPGateway))";
		"Serveur DNS" = ": $($NetworkAdapterConfiguration.DNSServerSearchOrder)";
		"Suffixe DNS recherch�" = ": $($NetworkAdapterConfiguration.DNSDomainSuffixSearchOrder)";
		"Nom DNS du PC" = ": $($NetworkAdapterConfiguration.DNSHostName)"
	}
	$Result.GetEnumerator() | Where-Object {$_.Value.Length -gt 2} | Format-table -HideTableHeaders -AutoSize
}
_Pause("Appuyez sur une touche pour quitter le script.")
if (-Not $psISE) {Exit-PSSession}
