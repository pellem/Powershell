#PARAMS
$SiteURL = "https://ChangeMe.sharepoint.com/" <-- MODIFY
$ListTitle = "ChangeMe" <-- MODIFY
$User = "ChangeMe" <-- MODIFY
$csvName = ChangeMe.csv"

#Chargement des librairies Sharepoint 2013 CSOM -- May change for every user
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Crée une variable qui pointe vers le chemin du script // Permet d'aller chercher le CSV dans le même dossier
#New var poiting on script's path // Allow us to find csv in same folder
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

#L'user entre son mot de passe \\ User types his password
$password = Read-Host -Prompt "Enter password" -AsSecureString

#Génère le contexte du client et définit les credentials \\ Handle context and user.creds
#ctx = client context 
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User, $password)
$ctx.Credentials = $credentials

 \\ Get the list wrote in params
$list = $ctx.get_web().get_lists().getByTitle($listTitle);

#Concatène le path du script + le nom du fichier CSV
#Concatenate script's path + csv file name
$csvFilePath = Join-Path -Path $scriptPath -ChildPath $csvName
if (Test-Path "$csvFilePath")
{

	#Récupère les objets pour la liste
	#Handle list items
	Write-Host "Récupération des données contenues dans $csvFilePath "
	$listItems = import-csv -Path "$csvFilePath"

	$recordCount = @($listItems).count;
	Write-Host -ForegroundColor Yellow "Il y a $recordCount objets à traiter"
	Write-Host "Veuillez patienter..."

	#Ajoute les objets à la liste
	for($rowCounter = 0; $rowCounter -le $recordCount - 1; $rowCounter++)
	{ 
		    $curItem = @($listItems)[$rowCounter];

		    Write-Progress -id 1 -activity "Ajout des objets à la liste..." -status "Ajout de l'objet $rowCounter sur $recordCount objets au total." -percentComplete ($rowCounter*(100/$recordCount));
		    #On crée les objets dans la liste
		    $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
		    $newItem = $list.addItem($itemCreateInfo);
			#On fait le lien entre les colonnes de la liste et celle du csv
		    $newItem.set_item('Title', $curItem.Titre);
		    $newItem.set_item('Description', $curItem.Description);
			$newItem.set_item('Age', $curItem.Age);
			$newItem.set_item('Sexe', $curItem.Sexe);
			#On met à jour la liste
		    $newItem.update();
		    $ctx.Load($newItem)
		    $ctx.ExecuteQuery()
	}
}
else
{
		#En cas d'erreur (si on a pas trouvé le ficher csv)
		Write-Host "Could not load file path  $csvFilePath "

}
