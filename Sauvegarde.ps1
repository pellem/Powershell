#################################################################################################################################
                                             ##           VARIABLES REQUISES         ##
#################################################################################################################################

#Récupération des données utilisateur
$Username = "USER"
$CredsPath = "Fichier Crypté"
$SiteURL = "URL"
$SiteToSave = "Nom du site" ##Sert uniquement a la gestion des noms de fichiers

#Chargement des librairies Sharepoint 2013 CSOM
Write-Host "Chargement des librairies CSOM" -foregroundcolor black -backgroundcolor yellow
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Write-Host "Librairies chargées avec succès" -foregroundcolor black -backgroundcolor Green 

#################################################################################################################################
                                             ##           FONCTIONS         ##
#################################################################################################################################

#################### - Renvoie l'URL du site sans "https://" ni ".com/". Remplace les "." par "_"
function Epur_URL ## - On se basera dessus pour générer les noms des dossiers
{				  ###########################################################################################	
$clearURL = $SiteURL
$clearURL = $SiteURL.Remove(0,8)
$index = ($clearURL.LastIndexOfAny(".sharepoint") - 1)
$clearURL = $clearURL.Remove($index,4)
$clearURL = $clearURL.replace('.','_')

Write-Output $clearURL
}

################################# - Connexion au site Sharepoint
function Connect_to_Sharepoint ## - Si la connexion est réussie, on stocke le contexte dans la variable globale
{					           ###########################################################################################	
 param (
  [Parameter(Mandatory=$true,Position=1)]
		[string]$Username,
        [Parameter(Mandatory=$true,Position=3)]
		[string]$Url
	   )

  $password = Get-Content $CredsPath | Convertto-SecureString
  $ctx=New-Object Microsoft.SharePoint.Client.ClientContext($Url)
  $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $password)
	try
	{
  	  Write-Host "Connecté !" -foregroundcolor black -backgroundcolor Green	
	}
	catch
	{
      Write-Host "Impossible de se connecter à $SiteURL : mauvais mot de passe ?" -foregroundcolor black -backgroundcolor Red
 	  return
	}
	$ctx.ExecuteQuery()
	$global:ctx=$ctx
}

$global:ctx #Définition de la globale


######################## - Renvoie toutes les listes trouvées via Write-Output -- Write-Output envoie les objets spécifiés sur le pipeline.				 
function Get-SPOList  ## - Si il n'y a pas de commande qui suit Write-Output, les objets sont affichés dans la console. 
{					  ###########################################################################################	
  
   param (
       	   [Parameter(Mandatory=$false,Position=0)]
		   [switch]$IncludeAllProperties
		 )
 
  $ctx.Load($ctx.Web.Lists) #On récupère les listes sur le serveur
  $ctx.ExecuteQuery()
  Write-Host 
  Write-Host $ctx.Url -BackgroundColor White -ForegroundColor DarkGreen
  foreach($ll in $ctx.Web.Lists) 
  {     
        $ctx.Load($ll.RootFolder)
        $ctx.Load($ll.DefaultView)
        $ctx.Load($ll.Views)
        $ctx.Load($ll.WorkflowAssociations)
        try
        {
       	   $ctx.ExecuteQuery()
        }
        catch
        {
        }
        if($IncludeAllProperties) #Si on a lancé la fonction avec -IncludeAllProperties
        {
        
        	$obj = New-Object PSObject
 		 	$obj | Add-Member NoteProperty Title($ll.Title)
 		 	$obj | Add-Member NoteProperty Created($ll.Created)
  		 	$obj | Add-Member NoteProperty Tag($ll.Tag)
  	     	$obj | Add-Member NoteProperty RootFolder.ServerRelativeUrl($ll.RootFolder.ServerRelativeUrl)
 		 	$obj | Add-Member NoteProperty BaseType($ll.BaseType)
		 	$obj | Add-Member NoteProperty BaseTemplate($ll.BaseTemplate)
		 	$obj | Add-Member NoteProperty AllowContenttypes($ll.AllowContenttypes)
		 	$obj | Add-Member NoteProperty ContentTypesEnabled($ll.ContentTypesEnabled)
		 	$obj | Add-Member NoteProperty DefaultView.Title($ll.DefaultView.Title)
		 	$obj | Add-Member NoteProperty Description($ll.Description)
 		 	$obj | Add-Member NoteProperty DocumentTemplateUrl($ll.DocumentTemplateUrl)
		 	$obj | Add-Member NoteProperty DraftVersionVisibility($ll.DraftVersionVisibility)
 		 	$obj | Add-Member NoteProperty EnableAttachments($ll.EnableAttachments)
		 	$obj | Add-Member NoteProperty EnableMinorVersions($ll.EnableMinorVersions)
 		 	$obj | Add-Member NoteProperty EnableFolderCreation($ll.EnableFolderCreation)
 		 	$obj | Add-Member NoteProperty EnableVersioning($ll.EnableVersioning)
 		 	$obj | Add-Member NoteProperty EnableModeration($ll.EnableModeration)
 		 	$obj | Add-Member NoteProperty Fields.Count($ll.Fields.Count)
 		 	$obj | Add-Member NoteProperty ForceCheckout($ll.ForceCheckout)
 		 	$obj | Add-Member NoteProperty Hidden($ll.Hidden)
		 	$obj | Add-Member NoteProperty Id($ll.Id)
		 	$obj | Add-Member NoteProperty IRMEnabled($ll.IRMEnabled)
 		 	$obj | Add-Member NoteProperty IsApplicationList($ll.IsApplicationList)
 		 	$obj | Add-Member NoteProperty IsCatalog($ll.IsCatalog)
 		 	$obj | Add-Member NoteProperty IsPrivate($ll.IsPrivate)
		 	$obj | Add-Member NoteProperty IsSiteAssetsLibrary($ll.IsSiteAssetsLibrary)
		 	$obj | Add-Member NoteProperty ItemCount($ll.ItemCount)
		 	$obj | Add-Member NoteProperty LastItemDeletedDate($ll.LastItemDeletedDate)
 		 	$obj | Add-Member NoteProperty MultipleDataList($ll.MultipleDataList)
		 	$obj | Add-Member NoteProperty NoCrawl($ll.NoCrawl)
 		 	$obj | Add-Member NoteProperty OnQuickLaunch($ll.OnQuickLaunch)
 		 	$obj | Add-Member NoteProperty ParentWebUrl($ll.ParentWebUrl)
		 	$obj | Add-Member NoteProperty TemplateFeatureId($ll.TemplateFeatureId)
 		 	$obj | Add-Member NoteProperty Views.Count($ll.Views.Count)
		 	$obj | Add-Member NoteProperty WorkflowAssociations.Count($ll.WorkflowAssociations.Count)

        Write-Output $obj
        }
        else #Par défaut, on récupère juste le titre de la liste (simplifie la lecture du fichier de sortie par les fonctions suivantes)
        {   
          $obj = New-Object PSObject
  		  $obj | Add-Member NoteProperty Title($ll.Title)
          Write-Output $obj
		
        } 
     }
}

############################# - Renvoie un objet par champs présent dans liste.
function Get-SPOListFields ## - 
{					       ###########################################################################################	
 param (
        [Parameter(Mandatory=$true,Position=3)]
		[string]$ListTitle)

  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.Load($ll.Fields)
  $ctx.ExecuteQuery()

  $fieldsArray=@()
  $fieldslist=@()
 foreach ($fiel in $ll.Fields) #Pour chaque colonne de chaque liste
 {
 	 #Décommenter si on veut la liste très exhaustive des éléments traités, sortie sur la console donc dans le fichier de log...
 	 #Write-Host $fiel.Description `t $fiel.EntityPropertyName `t $fiel.Id `t $fiel.InternalName `t $fiel.StaticName `t $fiel.Tag `t $fiel.Title  `t $fiel.TypeDisplayName
	 
	 #Cette condition permet de réduire la portée du scope; par défaut, la fonction renvoie TOUS les champs de la liste.
	 if ($fiel.InternalName -notcontains "ContentTypeId" -and $fiel.InternalName -notcontains "_ModerationComments" -and $fiel.InternalName -notcontains "File_x0020_Type" -and $fiel.InternalName -notcontains "Author"-and $fiel.InternalName -notcontains "Editor"-and $fiel.InternalName -notcontains "Modified" -and $fiel.InternalName -notcontains "Created" -and $fiel.InternalName -notcontains "LinkTitleNoMenu" -and $fiel.InternalName -notcontains "LinkTitle"-and $fiel.InternalName -notcontains "LinkTitle2"-and $fiel.InternalName -notcontains "ContentType" -and $fiel.InternalName -notcontains "_HasCopyDestinations" -and $fiel.InternalName -notcontains "_CopySource" -and $fiel.InternalName -notcontains "owshiddenversion" -and $fiel.InternalName -notcontains "WorkflowVersion" -and $fiel.InternalName -notcontains "_UIVersion" -and $fiel.InternalName -notcontains "Attachments" -and $fiel.InternalName -notcontains "_ModerationStatus" -and $fiel.InternalName -notcontains "Edit" -and $fiel.InternalName -notcontains "SelectTitle" -and $fiel.InternalName -notcontains "InstanceID" -and $fiel.InternalName -notcontains "Order" -and $fiel.InternalName -notcontains "GUID" -and $fiel.InternalName -notcontains "WorkflowInstanceID" -and $fiel.InternalName -notcontains "FilRef" -and $fiel.InternalName -notcontains "FileDirRef" -and $fiel.InternalName -notcontains "Last_x0020_Modified" -and $fiel.InternalName -notcontains "Created_x0020_Date" -and $fiel.InternalName -notcontains "FSObjType" -and $fiel.InternalName -notcontains "SortBehavior" -and $fiel.InternalName -notcontains "PermMask" -and $fiel.InternalName -notcontains "FileLeafRef"-and $fiel.InternalName -notcontains "UniqueID" -and $fiel.InternalName -notcontains "SyncClientId" -and $fiel.InternalName -notcontains "ProgId" -and $fiel.InternalName -notcontains "ScopeID" -and $fiel.InternalName -notcontains "HTML_x0020_File_x0020_Type" -and $fiel.InternalName -notcontains "_UIVersionString" -and $fiel.InternalName -notcontains "FileRef" -and $fiel.InternalName -notcontains "_EditMenuTableStart" -and $fiel.InternalName -notcontains "_EditMenuTableStart2" -and $fiel.InternalName -notcontains "_EditMenuTableEnd" -and $fiel.InternalName -notcontains "LinkFilenameNoMenu" -and $fiel.InternalName -notcontains "LinkFilename" -and $fiel.InternalName -notcontains "LinkFilename2" -and $fiel.InternalName -notcontains "DocIcon" -and $fiel.InternalName -notcontains "ServerUrl" -and $fiel.InternalName -notcontains "EncodedAbsUrl" -and $fiel.InternalName -notcontains "BaseName" -and $fiel.InternalName -notcontains "MetaInfo" -and $fiel.InternalName -notcontains "_Level" -and $fiel.InternalName -notcontains "_IsCurrentVersion" -and $fiel.InternalName -notcontains "ItemChildCount" -and $fiel.InternalName -notcontains "FolderChildCount" -and $fiel.InternalName -notcontains "Restricted" -and $fiel.InternalName -notcontains "OriginatorId" -and $fiel.InternalName -notcontains "AppAuthor" -and $fiel.InternalName -notcontains "AppEditor" -and $fiel.InternalName -notcontains "SMTotalSize" -and $fiel.InternalName -notcontains "SMLastModifiedDate" -and $fiel.InternalName -notcontains "SMTotalFileStreamSize" -and $fiel.InternalName -notcontains "SMTotalFileCount")
  	{
  		$array=@() #On déclare/reset un tableau, on définit ses 4 premières cases
  		$array+="InternalName"
    	$array+="StaticName"
      	$array+="Tag"
        $array+="Title"

  		$obj = New-Object PSObject #On déclare un nouvel objet, on lui attribue 4 propriétés
  		$obj | Add-Member NoteProperty $array[0]($fiel.InternalName)
  		$obj | Add-Member NoteProperty $array[1]($fiel.StaticName)
  		$obj | Add-Member NoteProperty $array[2]($fiel.Tag)
  		$obj | Add-Member NoteProperty $array[3]($fiel.Title)

  		$fieldsArray+=$obj #On remplit chaque case définie avec l'objet créé au dessus
  		$fieldslist+=$fiel.InternalName
  		Write-Output $obj
  	}
 }
  $ctx.Dispose() #On free la mémoire allouée au contexte
  return $fieldsArray #On renvoie le tableau rempli
}

############################ - Renvoie un tableau contenant tous les champs de la liste passée en paramètre
function Get-SPOListItems ##
{					      ###########################################################################################	
  
   param (
        [Parameter(Mandatory=$true,Position=4)]
		[string]$ListTitle,
        [Parameter(Mandatory=$false,Position=5)]
		[bool]$IncludeAllProperties=$false,
        [switch]$Recursive
		)
    
  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.Load($ll.Fields)
  $ctx.ExecuteQuery()

 $spqQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
	# $spqQuery.ViewAttributes = "Scope='Recursive'"

if($Recursive)
	{
 	  $spqQuery.ViewXml ="<View Scope='RecursiveAll' />";
	}
  $fieldsArray=Get-SPOListFields -ListTitle $ListTitle #On récupère toutes les colonnes de la liste

  $itemsCollection=$ll.GetItems($spqQuery) #On récupère un objet contenant la collection d'éléments retournés par la requête CAML
  $ctx.Load($itemsCollection)
  $ctx.ExecuteQuery()
 
  $objArray=@() #On déclare le tableau de sortie

##Ces boucles imbriquées permettent de remplir toutes les "colonnes" de chaque "ligne" de la liste
  for($j=0;$j -lt $itemsCollection.Count ;$j++) #j = 0 // Tant que j < ou = au nb. d'items dans la liste
  {
    Write-Progress -id 2 -Activity "Récupération des champs..." -Status "pour: $ListTitle..." -percentComplete ($j*(100/$itemsCollection.Count));
    
	$obj = New-Object PSObject    
    for($k=0;$k -lt $fieldsArray.Count ; $k++) #k = 0 // Tant que k < ou = au nb. de champs dans la liste
    {
         $name=$fieldsArray[$k].InternalName #On récupère le nom du champs (situé à la case $k)
         $value=$itemsCollection[$j][$name]	#On récupère la valeur correspondante au $name du champs pour la ligne $j
		 
		 #Ces conditions permettent de gérer les cas particuliers
		 if ($value.LookupValue) #Si la valeur possède un membre "LookupValue", alors c'est un pointeur, on récupère donc son nom
		 {
		    $obj | Add-Member NoteProperty $name($value.LookupValue) -Force
		 }
		 elseif ($name -eq "Th_x00e8_me") ## A gérer ...
		 {
		 	$obj | Add-Member NoteProperty $name([string]$value) -Force
		 }
		  elseif ($value.URL) #Si la valeur possède un membre "LookupValue", alors c'est un pointeur, on récupère donc son nom
		 {
		 	$obj | Add-Member NoteProperty $name($value.Description) -Force
		 }
		 else #Sinon, on remplit directement avec la valeur
		 {
         	$obj | Add-Member NoteProperty $name($value) -Force
         }
    }
    $objArray+=$obj #On rajoute la ligne dans l'objet
  }

  Write-Host "=======>" ((get-date).tostring(‘HH:mm:ss’))" Liste $ListTitle Enregistrée <========"
  return $objArray
  }

#################################################################################################################################
                                             ##           EXECUTION          ##
#################################################################################################################################

$clearURL = Epur_URL;
#On vérifie l'existance de l'arborescence, si elle n'existe pas ou est incomplète, alors on crée les dossiers 
$save_dir = $clearURL + "save"
$save_dir_path = ".\" + $save_dir + "\Sites\" + "$SiteToSave"
If (-not (Test-Path $save_dir)) { New-Item -ItemType Directory -Name $save_dir }
If (-not (Test-Path $save_dir\Sites)) { New-Item -ItemType Directory -Name $save_dir\Sites }
If (-not (Test-Path $save_dir\Sites\$SiteToSave\Logs)) { New-Item -ItemType Directory -Name $save_dir\Sites\$SiteToSave\Logs }
If (-not (Test-Path $save_dir\Sites\$SiteToSave)) { New-Item -ItemType Directory -Name $save_dir\Sites\$SiteToSave }
If (-not (Test-Path $save_dir\Sites\$SiteToSave\Logs\old)) { New-Item -ItemType Directory -Name $save_dir\Sites\$SiteToSave\Logs\old }
If (-not (Test-Path $CredsPath))
{
New-Item -path $CredsPath -ItemType file
Read-Host -Prompt "Password" -AsSecureString | ConvertFrom-SecureString | Out-file $CredsPath -Force
}
Move-Item -Path $save_dir\Sites\$SiteToSave\Logs\*.txt $save_dir\Sites\$SiteToSave\Logs\old

#On se connecte au site sharepoint grâce à la fonction Connect_to_Sharepoint
Connect_to_Sharepoint $Username $SiteURL;

$log = $SiteToSave + "_" + ((get-date).tostring(‘ddMMyy_HHmmss’)) + ".txt" #On génère un nom pour le fichier de log
Start-Transcript -Path  $save_dir\Sites\$SiteToSave\Logs\$log                     #|Enregistre tout ce qui se passe dans la console dans $log
$sw = [Diagnostics.Stopwatch]::StartNew() #|Début du chronomètre (mesure le temps d'execution du script)

#On récupère l'URL épurée (pour générer les path des fichiers de sortie)
$clearURL = Epur_URL;

#On récupère toutes les listes grâce à la fonction Get-SPOList
$AllLists = Get-SPOList;


#On crée un .csv qui contient toutes les listes trouvées sur le sharepoint
$AllLists | export-csv -Path $save_dir_path\lists.csv -notype -Encoding UTF8;

#On stocke toutes les listes dans la variable $listItems
$listItems = import-csv -Path $save_dir_path\lists.csv
$recordCount = @($listItems).count; #On compte le nombre de listes
$tab_lists=@()
for($listCounter = 0; $listCounter -le $recordCount - 1; $listCounter++) #Pour chaque liste
{ 
    $curItem = @($listItems)[$listCounter]; 
	$tab_lists += ($curItem.Title) #On récupère un tableau avec chaque nom de liste
}
for($index = 1; $index -le $recordCount - 1; $index++) #Tant que index < nombre de listes à traiter
{
    Write-Progress -id 1 -activity "Sauvegarde en cours..." -status "(liste n°$index sur $recordCount)" -percentComplete ($index*(100/$recordCount));
	$tab_items=@() #A chaque itération, on crée/reset $tab_items en tant que tableau vide
	$tab_items = Get-SPOListItems $tab_lists[$index]; #On le remplit avec l'objet retourné par Get-SPOListItems
	$tab_path = $save_dir_path + "\" + $tab_lists[$index].replace(' ','_') +".csv" #On crée un path pour chaque liste
	$tab_items | export-csv -Path $tab_path -notype -Encoding UTF8; #On exporte / -NoTypeInformation permet de supprimer la première ligne (obsolete) / On encore en UTF8 pour gérer les accents
}

$sw.Stop()
$sw.Elapsed #On arrête le chronomètre et on l'affiche
Stop-Transcript #On arrête la retransmission du fichier de log
