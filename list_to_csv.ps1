#title           :csv_to_list.ps1
#description     :Will grab a list from Sharepoint server and save it as a .csv file.
#author		 :GordonAmable
#date            :13/10/2015
#version         :0.1    
#usage		 :./csv-to-list.ps1
#notes           :Need Sharepoint 2013 CSOM librairies.
#==============================================================================

#PARAMS
$SiteURL = "https://ChangeMe.sharepoint.com/" <-- MODIFY
$ListTitle = "ChangeMe" <-- MODIFY
$User = "ChangeMe" <-- MODIFY

#Chargement des librairies Sharepoint 2013 CSOM -- May change for every user
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#L'user entre son mot de passe \\ User types his password
$password = Read-Host -Prompt "Enter password" -AsSecureString

#Génère le contexte du client et définit les credentials \\ Handle context and user.creds
#ctx = client context
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User, $password)
$ctx.Credentials = $credentials

#On récupère la liste définie dans la variable \\ Get the list wrote in params
$list = $ctx.get_web().get_lists().getByTitle($listTitle);

#On définit une requête CAML qui cible tous les éléments dont l'ID > ou = à 1
#Define CAML query like: SELECT * From $listTitle WHERE ID >= 0
#Geq : Greater or equal
$query = New-Object Microsoft.SharePoint.Client.CamlQuery;
$query.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID' /><Value Type='Counter'>1</Value></Geq></Where></Query></View>"

#On éxécute la requête \\ Execute Query
$listItems = $list.GetItems($query);
$ctx.Load($listItems);
$ctx.ExecuteQuery();

#Création de la variable tableau (qu'on va copier dans le fichier de sortie)
#Create new array (which will be copied in output file)
$tableau =@();

#On boucle tant qu'il y a des éléments dans la liste (après l'éxécution de la requête)
#While items exists in list (after query)
foreach ($listItem in $listItems)
{
	#On crée un nouvel objet PowerShell
	#New PowerShell Object
    $result = new-object psobject
	
	#On crée autant de colonnes qu'on veut dans l'objet qu'on vient de créer et on leur donne la valeur correspondante dans la liste
	#Here we set as much columns as we want in the item we just created
  $result | Add-Member -MemberType noteproperty -Name Title -value $listItem['Title'];
	$result | Add-Member -MemberType noteproperty -Name ChangeMe1 -value $listItem['ChangeMe1'];
	$result | Add-Member -MemberType noteproperty -Name ChangeMe2 -value $listItem['ChangeMe2'];
	$result | Add-Member -MemberType noteproperty -Name ChangeMe3 -value $listItem['ChangeMe3'];
    
	#On ajoute l'objet dans la variable tableau
	#Add object in array
	$tableau += $result;
}
#On génère le nom du fichier de sortie en fonction de la liste passée en paramètre en transformant les espaces en '_' (underscore)
#Generate output file name as: "Export_$listTitle" where spaces become '_' (underscore)
$CsvName = "Export_"+$ListTitle.replace(' ','_')+".csv"

#On exporte nôtre tableau dans un fichier CSV tel que définit a la ligne précédente
#Export array in csv file like we defined above
$tableau | export-csv $CsvName -noTypeInformation;

