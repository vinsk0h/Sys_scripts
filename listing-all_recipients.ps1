# Importez le module Exchange Online s'il n'est pas déjà chargé
if ((Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue) -eq $null) {
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}

# Demandez à l'utilisateur de saisir le nom de son tenant
$tenantName = Read-Host "Entrez le nom de votre tenant (par exemple, contoso.onmicrosoft.com) :"

# Construisez l'adresse e-mail Microsoft 365 en utilisant le tenant
$userPrincipalName = "$tenantName"

# Connectez-vous à votre compte Microsoft 365
Connect-ExchangeOnline -UserPrincipalName $userPrincipalName -ShowProgress $true

# Demandez à l'utilisateur de saisir l'adresse e-mail de l'expéditeur
$expediteur = Read-Host "Entrez l'adresse e-mail de l'expéditeur :"

# Recherchez tous les utilisateurs ayant reçu un e-mail de cet expéditeur
$utilisateurs = Get-MessageTrace -SenderAddress $expediteur | Select-Object -ExpandProperty RecipientAddress

# Affichez la liste des utilisateurs
$utilisateurs | Sort-Object -Unique | Format-Table -Property RecipientAddress

# Exportez la liste des utilisateurs vers un fichier texte
$expediteurFileName = "$expediteur.txt"
$utilisateurs | Sort-Object -Unique | Out-File -FilePath $expediteurFileName

Write-Host "La liste des utilisateurs a été exportée vers $expediteurFileName."


# Déconnectez-vous de votre session Exchange Online
Disconnect-ExchangeOnline -Confirm:$false