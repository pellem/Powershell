
#title           :save-site.ps1
#description     :Will save as .csv all lists + all fields for each list found on a Sharepoint site
#author		 :GordonAmable
#date            :20/10/2015
#version         :0.4    
#usage		 :./save-site.ps1
#notes           :Need Sharepoint 2013 CSOM librairies.
#==============================================================================

#################################################################################################################################
                                             ##           VARIABLES REQUISES         ##
#################################################################################################################################



#Récupération des données utilisateur
$Username = Read-Host -Prompt "User"
$SiteURL = "https://johndoe.sharepoint.com/"     #Modify with your own site


#Chargement des librairies Sharepoint 2013 CSOM -- Peut-être à modifier pour chaque nouvelle machine
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"


#################################################################################################################################
                                             ##           FONCTIONS         ##
#################################################################################################################################


function Epur_URL #Renvoie l'URL du site sans "https://" ni ".com/". Remplace les "." par "_"
{
$clearURL = $SiteURL
$clearURL = $SiteURL.Remove(0,8)
$index = ($clearURL.LastIndexOfAny(".sharepoint") - 1)
$clearURL = $clearURL.Remove($index,4)
$clearURL = $clearURL.replace('.','_')

Write-Output $clearURL
}


function Connect_to_Sharepoint #Crée le contexte de connexion au sharepoint et le stocke dans une globale
{
 param (
  [Parameter(Mandatory=$true,Position=1)]
		[string]$Username,
        [Parameter(Mandatory=$true,Position=3)]
		[string]$Url
)

  $password = Read-Host "Password" -AsSecureString
  $ctx=New-Object Microsoft.SharePoint.Client.ClientContext($Url)
  $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $password)
  $ctx.ExecuteQuery()  
$global:ctx=$ctx
}

$global:ctx



					  #Renvoie toutes les listes trouvées via Write-Output -- Write-Output envoie les objets spécifiés sur le pipeline.				 
function Get-SPOList  #Si il n'y a pas de commande qui suit Write-Output, les objets sont affichés dans la console. 
{
  
   param (
        [Parameter(Mandatory=$false,Position=0)]
		[switch]$IncludeAllProperties
		)
 
  $ctx.Load($ctx.Web.Lists)
  $ctx.ExecuteQuery()
  Write-Host 
  Write-Host $ctx.Url -BackgroundColor White -ForegroundColor DarkGreen
  foreach( $ll in $ctx.Web.Lists)
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

        if($IncludeAllProperties)
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
        else
        {

        
       
        $obj = New-Object PSObject
  		$obj | Add-Member NoteProperty Title($ll.Title)
		#$obj | Add-Member -MemberType noteproperty -Name Title -value $ll['Title'];
 # $obj | Add-Member NoteProperty Created($ll.Created)
  #$obj | Add-Member NoteProperty RootFolder.ServerRelativeUrl($ll.RootFolder.ServerRelativeUrl)
        
 
        Write-Output $obj
        
        
     }  
        
        }
}




function Get-SPOListFields #Renvoie un objet par champs présent dans liste.
{
 param (
        [Parameter(Mandatory=$true,Position=3)]
		[string]$ListTitle)
#        [Parameter(Mandatory=$false,Position=4)]
#		[bool]$IncludeSubsites=$false#>
#		)

  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.Load($ll.Fields)
  $ctx.ExecuteQuery()


  $fieldsArray=@()
  $fieldslist=@()
 foreach ($fiel in $ll.Fields)
 {
  #Write-Host $fiel.Description `t $fiel.EntityPropertyName `t $fiel.Id `t $fiel.InternalName `t $fiel.StaticName `t $fiel.Tag `t $fiel.Title  `t $fiel.TypeDisplayName
 if ($fiel.InternalName -notcontains "ContentTypeId" -and $fiel.InternalName -notcontains "_ModerationComments" -and $fiel.InternalName -notcontains "File_x0020_Type" -and $fiel.InternalName -notcontains "Author"-and $fiel.InternalName -notcontains "Editor"-and $fiel.InternalName -notcontains "Modified" -and $fiel.InternalName -notcontains "Created" -and $fiel.InternalName -notcontains "LinkTitleNoMenu" -and $fiel.InternalName -notcontains "LinkTitle"-and $fiel.InternalName -notcontains "LinkTitle2"-and $fiel.InternalName -notcontains "ContentType" -and $fiel.InternalName -notcontains "_HasCopyDestinations" -and $fiel.InternalName -notcontains "_CopySource" -and $fiel.InternalName -notcontains "owshiddenversion" -and $fiel.InternalName -notcontains "WorkflowVersion" -and $fiel.InternalName -notcontains "_UIVersion" -and $fiel.InternalName -notcontains "Attachments" -and $fiel.InternalName -notcontains "_ModerationStatus" -and $fiel.InternalName -notcontains "Edit" -and $fiel.InternalName -notcontains "SelectTitle" -and $fiel.InternalName -notcontains "InstanceID" -and $fiel.InternalName -notcontains "Order" -and $fiel.InternalName -notcontains "GUID" -and $fiel.InternalName -notcontains "WorkflowInstanceID" -and $fiel.InternalName -notcontains "FilRef" -and $fiel.InternalName -notcontains "FileDirRef" -and $fiel.InternalName -notcontains "Last_x0020_Modified" -and $fiel.InternalName -notcontains "Created_x0020_Date" -and $fiel.InternalName -notcontains "FSObjType" -and $fiel.InternalName -notcontains "SortBehavior" -and $fiel.InternalName -notcontains "PermMask" -and $fiel.InternalName -notcontains "FileLeafRef"-and $fiel.InternalName -notcontains "UniqueID" -and $fiel.InternalName -notcontains "SyncClientId" -and $fiel.InternalName -notcontains "ProgId" -and $fiel.InternalName -notcontains "ScopeID" -and $fiel.InternalName -notcontains "HTML_x0020_File_x0020_Type" -and $fiel.InternalName -notcontains "_UIVersionString" -and $fiel.InternalName -notcontains "FileRef" -and $fiel.InternalName -notcontains "_EditMenuTableStart" -and $fiel.InternalName -notcontains "_EditMenuTableStart2" -and $fiel.InternalName -notcontains "_EditMenuTableEnd" -and $fiel.InternalName -notcontains "LinkFilenameNoMenu" -and $fiel.InternalName -notcontains "LinkFilename" -and $fiel.InternalName -notcontains "LinkFilename2" -and $fiel.InternalName -notcontains "DocIcon" -and $fiel.InternalName -notcontains "ServerUrl" -and $fiel.InternalName -notcontains "EncodedAbsUrl" -and $fiel.InternalName -notcontains "BaseName" -and $fiel.InternalName -notcontains "MetaInfo" -and $fiel.InternalName -notcontains "_Level" -and $fiel.InternalName -notcontains "_IsCurrentVersion" -and $fiel.InternalName -notcontains "ItemChildCount" -and $fiel.InternalName -notcontains "FolderChildCount" -and $fiel.InternalName -notcontains "Restricted" -and $fiel.InternalName -notcontains "OriginatorId" -and $fiel.InternalName -notcontains "AppAuthor" -and $fiel.InternalName -notcontains "AppEditor" -and $fiel.InternalName -notcontains "SMTotalSize" -and $fiel.InternalName -notcontains "SMLastModifiedDate" -and $fiel.InternalName -notcontains "SMTotalFileStreamSize" -and $fiel.InternalName -notcontains "SMTotalFileCount")
 {
  $array=@()
  $array+="InternalName"
    $array+="StaticName"
      $array+="Tag"
       $array+="Title"

  $obj = New-Object PSObject
  $obj | Add-Member NoteProperty $array[0]($fiel.InternalName)
  $obj | Add-Member NoteProperty $array[1]($fiel.StaticName)
  $obj | Add-Member NoteProperty $array[2]($fiel.Tag)
  $obj | Add-Member NoteProperty $array[3]($fiel.Title)

  $fieldsArray+=$obj
  $fieldslist+=$fiel.InternalName
  Write-Output $obj
  }
 }
 

 $ctx.Dispose()
  return $fieldsArray

}


  
function Get-SPOListItems #Renvoie un tableau contenant tous les objets (contenant chaque champs) pour la liste passée en paramètres
{
  
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
  $i=0



 $spqQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
# $spqQuery.ViewAttributes = "Scope='Recursive'"

if($Recursive)
{
$spqQuery.ViewXml ="<View Scope='RecursiveAll' />";
}
   $bobo=Get-SPOListFields -ListTitle $ListTitle 


  $itemki=$ll.GetItems($spqQuery)
  $ctx.Load($itemki)
  $ctx.ExecuteQuery()

  
 
 $objArray=@()

  for($j=0;$j -lt $itemki.Count ;$j++)
  {
        Write-Progress -id 2 -Activity "Récupération des champs..." -Status "pour: $ListTitle..." -percentComplete ($j*(100/$itemki.Count));
        $obj = New-Object PSObject
        
#        if($IncludeAllProperties)
#        {

        for($k=0;$k -lt $bobo.Count ; $k++)
        {
          
         # Write-Host $k
         $name=$bobo[$k].InternalName
         $value=$itemki[$j][$name]
          $obj | Add-Member NoteProperty $name($value) -Force
          
        }

#        }
#        else
#        {
#          $obj | Add-Member NoteProperty ID($itemki[$j]["ID"])
#          $obj | Add-Member NoteProperty Title($itemki[$j]["Title"])
#
#        }

      #  Write-Host $obj.ID `t $obj.Title
        $objArray+=$obj
    
   
  }

 
  
  return $objArray
  
  
  }

#################################################################################################################################
                                             ##           EXECUTION          ##
#################################################################################################################################


clear


#On se connecte au site sharepoint grâce à la fonction Connect_to_Sharepoint
Connect_to_Sharepoint $Username $SiteURL;



#On récupère l'URL épurée (pour générer les path des fichiers de sortie)
$clearURL = Epur_URL



#On récupère toutes les listes grâce à la fonction Get-SPOList
$AllLists = Get-SPOList;



#On check si un dossier au nom de l'URL épurée suivi de "save" existe. Si non, on le crée
$save_dir = $clearURL + "save"
$save_dir_path = ".\" + $save_dir
If (-not (Test-Path $save_dir)) { New-Item -ItemType Directory -Name $save_dir }



#On crée un .csv qui contient toutes les listes trouvées sur le sharepoint
$AllLists | export-csv -Path $save_dir_path\lists.csv -Encoding unicode;



#On stocke toutes les listes dans la variable $listItems
$listItems = import-csv -Path $save_dir_path\lists.csv

	$recordCount = @($listItems).count;
	$tab_lists=@()
for($rowCounter = 0; $rowCounter -le $recordCount - 1; $rowCounter++)
	{ 
		    $curItem = @($listItems)[$rowCounter];
			$tab_lists += ($curItem.Title)
	}

for($index = 1; $index -le $recordCount; $index++) #Tant que index < nombre de listes à traiter
{
		    Write-Progress -id 1 -activity "Sauvegarde en cours..." -status "(liste n°$index sur $recordCount)" -percentComplete ($index*(100/$recordCount));
			$tab_items=@() #A chaque itération, on crée/reset $tab_items en tant que tableau vide
			$tab_items = Get-SPOListItems $tab_lists[$index]; #On le remplit avec l'objet retourné par Get-SPOListItems
			$tab_path = $save_dir_path + "\" + $tab_lists[$index].replace(' ','_') +".csv" #On crée un path pour chaque liste			$tab_items | export-csv -Path $tab_path -Encoding unicode; #On exporte
}
