#-------------------------------------------------------------#
#----Initiele Declaraties-------------------------------------#
#-------------------------------------------------------------#
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="800" Height="440" Name="SPIRA_Hoofdvenster" Title="SPIRA v 1.5" WindowStartupLocation="CenterScreen">
<Grid Background="#6c5a82">
  
  
<TabControl Margin="9,9,9,9">
    <TabItem Header="Over SPIRA"><Grid Background="#c2b280">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Script voor Pre-Ingest RAZU (SPIRA)" Margin="15,15,0,0" FontSize="27" FontWeight="Bold"/>
<StackPanel Orientation="Vertical" Width="450" Margin="15,50,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Height="250">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" Width="350" Height="250" TextWrapping="Wrap" Text="SPIRA is een script geschreven door het Regionaal Archief Zuid-Utrecht. Met SPIRA worden een aantal aanpassingen gemaakt aan de export van Zaaksysteem.nl (nadat deze is omgezet naar .xml met de ExportTool van VHIC). Het RAZU heeft SPIRA beschikbaar gesteld onder de Creative Commons Naamsvermelding-NietCommercieel-GelijkDelen 4.0 Internationaal-licentie. Voor meer informatie zie de SPIRA Github pagina." Margin="15,50,15,15" FontSize="12"/>

      </StackPanel>
      <Image Width="200" Source="https://www.razu.nl/wp-content/uploads/2019/11/RAZU-logo-wit.png" VerticalAlignment="Top" Margin="450,75,0,0" Opacity="1"/>
      <TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="SPIRA versie 1.5 | augustus 2020 |  WB en BK" Margin="500,300,0,0"/>
      </Grid></TabItem>
  <TabItem Header="Opties"><Grid Background="#c2b280">
  <CheckBox Content="Originele bestanden niet overschrijven" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="103,15,0,0" Name="Button_Overschrijven" Width="250" Height="35" IsEnabled="True"/>
  <CheckBox Content="Datum splitsen (zaken)" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="103,50,0,0" Name="Button_DS_zaken" Width="250" Height="35" IsEnabled="True"/>
  <CheckBox Content="Datum splitsen (documenten)" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="103,85,0,0" Name="Button_DS_docs" Width="250" Height="35" IsEnabled="True"/>
  <CheckBox Content="Datum splitsen (aangepast)" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="103,120,0,0" Name="Button_DS_custom" Width="250" Height="35" IsEnabled="False"/>
  <CheckBox Content="Relaties vertalen" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="103,155,0,0" Name="Button_Rel_vertalen" Width="250" Height="35" IsEnabled="True"/>
  <CheckBox Content="Relaties samenvoegen" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="103,190,0,0" Name="Button_Rel_koppelen" Width="250" Height="35" IsEnabled="True"/>
  <CheckBox Content="Labels vertalen" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="103,225,0,0" Name="Button_Labels_vertalen" Width="250" Height="35" IsEnabled="False"/>
  <CheckBox Content="Aangepaste velden" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="103,260,0,0" Name="Button_Velden_custom" Width="250" Height="35" IsEnabled="False"/>
  <TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="23" Width="120" TextWrapping="Wrap" Name="Txtbox_alt_naam" Margin="500,15,0,0" Background="#ffffff" IsEnabled="false"/>
  <TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Voer achtervoegsel in:	" Margin="365,19,0,0"/>
  <TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Locatie export:" Margin="365,89,0,0"/>
  <Button Content="Selecteer..." HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Margin="515,89,0,0" Name="Button_loc_export"/>
  <TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="[Nog geen locatie opgegeven]" Margin="365,121,0,0" Name="Loc_export"/>
  <TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Alternatieve opslaglocatie:" Margin="365,159,0,0"/>
  <Button Content="Selecteer..." HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Margin="515,159,0,0" Name="Button_alt_loc" IsEnabled="false"/>
  <TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="[Nog geen locatie opgegeven]" Margin="365,191,0,0" Name="Loc_alt"/>
  <TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Aangepaste waarden-XML:" Margin="365,229,0,0"/>
  <Button Content="Selecteer..." HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Margin="515,229,0,0" Name="Button_alt_velden" IsEnabled="false"/>
  <TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="[Nog geen bestand opgegeven]" Margin="365,261,0,0" Name="Velden_alt"/>
</Grid></TabItem>
<TabItem Header="Uitvoeren"><Grid Background="#c2b280">
<CheckBox HorizontalAlignment="Left" VerticalAlignment="Top" Content="Vink dit aan als u klaar bent met het aanpassen van de opties en de mogelijkheid wilt krijgen om het SPIRA-script te draaien.
" Margin="37,6,0,0" Name="Button_WeetZeker" IsEnabled="True"/>
<Button Content="Start SPIRA" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="36.33331298828125,40,0,0" IsEnabled="False" Name="Button_run_spira"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Uitvoer:" Margin="37,65,0,0" FontWeight="Bold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="[[SPIRA is inactief]]" Margin="37,95,0,0" Name="SPIRA_Uitvoer"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Status:" Margin="275,65,0,0" FontWeight="Bold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="[[SPIRA is inactief]]" Margin="275,95,0,0" Name="SPIRA_Status"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Huidige stap:" Margin="275,130,0,0" FontWeight="Bold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="[[SPIRA is inactief]]" Margin="275,160,0,0" Name="SPIRA_Stap"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Aantaal zaken:" Margin="275,195,0,0" FontWeight="Bold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="[[SPIRA is inactief]]" Margin="275,225,0,0" Name="SPIRA_Aantal_Zaken"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Aantal documenten:" Margin="275,260,0,0" FontWeight="Bold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="[[SPIRA is inactief]]" Margin="275,290,0,0" Name="SPIRA_Aantal_Docs"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Aantal relaties:" Margin="425,195,0,0" FontWeight="Bold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="[[SPIRA is inactief]]" Margin="425,225,0,0" Name="SPIRA_Aantal_relaties"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Aantal splits:" Margin="425,260,0,0" FontWeight="Bold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="[[SPIRA is inactief]]" Margin="425,290,0,0" Name="SPIRA_Aantal_splits"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Aantal labels:" Margin="575,195,0,0" FontWeight="Bold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Deze optie is nog niet beschikbaar" Margin="575,225,0,0" Name="SPIRA_Aantal_labels"/>
  </Grid></TabItem>
  </TabControl>
</Grid>
</Window>
"@

#-------------------------------------------------------------#
#----Functies en event-handlers-------------------------------#
#-------------------------------------------------------------#

function Selecteer_bestand
{
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = $PSScriptroot 
    Filter = 'eXtensible Markup Language file (*.xml)|*.xml'
    Title = 'Selecteer een XML-bestand'
}
$FileBrowser.ShowDialog()
$Velden_alt.Text = $FileBrowser.FileName
}
function Selecteer_pad_export
{
$dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
$dialog.Description = 'Selecteer de map van de export'
$dialog.RootFolder = [System.Environment+specialfolder]::Desktop
$dialog.ShowNewFolderButton = $false
$dialog.ShowDialog()
$Loc_export.Text = $dialog.SelectedPath
$dialog.Dispose()
}
function Selecteer_pad_alt
{
$dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
$dialog.Description = 'Selecteer de map van de export'
$dialog.RootFolder = [System.Environment+specialfolder]::Desktop
$dialog.ShowNewFolderButton = $false
$dialog.ShowDialog()
$Loc_alt.Text = $dialog.SelectedPath
$dialog.Dispose()
}
function Set_active_alts($chk_status)
{
    if($chk_status -eq 1)
    {
        $Button_alt_loc.IsEnabled = $true
        $Txtbox_alt_naam.IsEnabled = $true
    }
    if($chk_status -eq 0)
    {
        $Button_alt_loc.IsEnabled = $false
        $Txtbox_alt_naam.IsEnabled = $false
    }
    
}
function Set_active_altxml ($chk_status )
{
    if($chk_status -eq 1)
    {
        $Button_alt_velden.IsEnabled = $true
    }
    if($chk_status -eq 0)
    {
        $Button_alt_velden.IsEnabled = $false
    }
}
function Activeer_spira ($chk_status )
{
    if($chk_status -eq 1)
    {
        $Button_run_spira.IsEnabled = $true
    }
    if($chk_status -eq 0)
    {
        $Button_run_spira.IsEnabled = $false
    }
}
function Opstarten
{
#Standaard waarden extraheren
foreach($item in (Select-XML -Xml $Global:XML_standaard_waarden -XPath '//SPIRA'))
	{
#Relaties vertalen - inladen relatienamen [[NOG AANPASSEN=>XML moet nog worden ingeladen]]
if($Button_Rel_vertalen.IsChecked -eq $true)
  {
      $global:relatie_child = $item.node.REL_CHILD
      $global:relatie_parent = $item.node.REL_PARENT
      $global:relatie_plain = $item.node.REL_PLAIN
  }else {
    $global:relatie_child = 'child'
    $global:relatie_parent = 'parent'
    $global:relatie_plain = 'plain'
  }

}
  #Extra waarden inladen // LET OP: Extra waarden worden nog niet ondersteund en dus verder ook niet verwerkt!
  if($Button_Velden_custom.IsChecked -eq $true)
{
    
    $Global:XML_extra_waarden = New-Object -TypeName XML
	$XML_extra_waarden_pad = $PSScriptRoot + "\SPIRA-ExtraWaarden-V1.5.xml"
	if (-not (Test-Path -Path $XML_extra_waarden_pad))
	{
    #>>Throw "FOUTMELDING: Extra waarden-bestand niet gevonden. Plaats het extra waarden-bestand in dezelfde map als het script en zorg er voor dat het de volgende syntax heeft: SPIRA-ExtraWaarden-V#.#.xml, waarbij de versienummers identiek zijn aan het SPIRA-scriptbestand." 
	}else{
	Write-Host "Extra waarden-bestand gevonden. Extra waarden inladen..."
	}
	$Global:XML_extra_waarden.Load($XML_extra_waarden_pad)
}
#Een aantal variabelen alvast declareren (niet altijd nodig, maar voorkomt irritaties)
$global:related_cases_volledig = ""    
$global:datumsplits = 0
$global:relatiesgelegd = 0
$global:labelsvertaald = 0

}
function Voorman
{
$SPIRA_Uitvoer.Text = "Opstarten..."
Opstarten
$SPIRA_Status.Text = "Bezig..."
#Zaken en documenten indexeren
$SPIRA_Uitvoer.Text = $SPIRA_Uitvoer.Text + "`n" + "Bestanden zoeken."
$SPIRA_Stap.Text = 'Bestanden zoeken'
BestandenZoeken 
$SPIRA_Aantal_Zaken.Text = $global:aantal_zaken
$SPIRA_Aantal_Docs.Text = $global:aantal_docs
$SPIRA_Uitvoer.Text = $SPIRA_Uitvoer.Text + "`n" + "Bestanden gevonden."
$SPIRA_Uitvoer.Text = $SPIRA_Uitvoer.Text + "`n" + "Zaken verwerken."
#Zaken naar de StappenDoorlopen-functie sturen
$SPIRA_Stap.Text = "Zaak-XMLs doorlopen"

Do
{
$XML_naam = $global:gevonden_zaken[$global:aantal_zaken]
$global:xml_geladen = New-Object -TypeName XML
$global:xml_geladen.Load($XML_naam)
StappenDoorlopen('Zaak')
$global:aantal_zaken--
}While($global:aantal_zaken -igt 0)
$SPIRA_Uitvoer.Text = $SPIRA_Uitvoer.Text + "`n" + "Zaken verwerkt."
$SPIRA_Aantal_splits.Text = $global:datumsplits
$SPIRA_Aantal_relaties.Text = $global:relatiesgelegd
$SPIRA_Uitvoer.Text = $SPIRA_Uitvoer.Text + "`n" + "Documenten verwerken."

#Documenten naar de StappenDoorlopen-functie sturen
$SPIRA_Stap.Text = "Document-XMLs doorlopen"
Do
{
$XML_naam = $global:gevonden_docs[$global:aantal_docs]
$global:xml_geladen = New-Object -TypeName XML
$global:xml_geladen.Load($XML_naam)
StappenDoorlopen('Document')
$global:aantal_docs--
}While($global:aantal_docs -igt 0)
$SPIRA_Uitvoer.Text = $SPIRA_Uitvoer.Text + "`n" + "Documenten verwerkt."
$SPIRA_Aantal_splits.Text = $global:datumsplits

##Afronding
$SPIRA_Stap.Text = "Afgerond"
$SPIRA_Status.Text = "Inactief"
$SPIRA_Uitvoer.Text = $SPIRA_Uitvoer.Text + "`n" + "SPIRA is afgerond. Sluit het venster."
}
function BestandenZoeken
{
#Inlezen bestands- en mappennamen, en het klaarzetten van de variabelen en arrays.
$folders = Get-ChildItem -path $Loc_export.text  -Recurse | Where-Object{ $_.PSIsContainer } | Select-Object FullName, Name, DirectoryName
$bestanden = Get-ChildItem -path $Loc_export.text -Recurse -Filter *.xml | Select-Object FullName, BaseName, Name, DirectoryName
$blobs = Get-ChildItem -path $Loc_export.text -Recurse -Filter *.blob | Select-Object FullName, BaseName, Name, DirectoryName
$count = 1
$global:aantal_zaken = 0
$global:gevonden_zaken = @(0..999)
$global:gevonden_zaken_kort = @(0..999)
$global:gevonden_zaken_dir = @(0..999)
$global:aantal_docs = 0
$global:gevonden_docs = @(0..999)
$global:gevonden_docs_kort = @(0..999)
$global:gevonden_docs_dir = @(0..999)

# Zaken zoeken => vergelijken mapnaam met XML-naam. De zaak-xml heeft namelijk dezelfde naam als de mapnaam van de zaak.
foreach($x in $folders)
{
    
    foreach($y in $bestanden)
    {
    if ($y.BaseName -eq $x.Name)
    {
    $global:gevonden_zaken[$count] = $y.FullName
    $global:gevonden_zaken_kort[$count] = $y.BaseName
    $global:gevonden_zaken_dir[$count] = $y.DirectoryName
    $global:aantal_zaken++
    }else{
    $count--
    }
    $count++
    }
    }
# Foutmelding als er geen zaken zijn gevonden
if(-not ($global:aantal_zaken -ge 1))
{
  Write-Error "Waarschuwing: Geen zaken gevonden! Weet u zeker dat u het op de juiste locatie uitvoert?"
  Write-Host ""
  Write-Host "Deze foutmelding kan voorkomen als u niet eerst de VHIC Export-Tool hebt uitgevoerd om de JSON bestanden naar XML te converten. SPIRA werkt enkel met XML-bestanden."
  Write-Host "U kunt het script voortzetten, als u dit bewust hebt gedaan om enkel documenten te verwerken."
  Tobeornottobe
}
# Bestanden zoeken => vergelijken blobnaam met XML-naam. De blob-xml heeft namelijk dezelfde naam als de blobnaam van het document.
$count = 1
foreach($x in $blobs)
{
  
    foreach($y in $bestanden)
    {
    if ($y.BaseName -eq $x.BaseName)
    {
    $global:gevonden_docs[$count] = $y.FullName
    $global:gevonden_docs_kort[$count] = $y.BaseName
    $global:gevonden_docs_dir[$count] = $y.DirectoryName
    $global:aantal_docs++
    }else{
    $count--
    }
    $count++
    }
    }
}
function StappenDoorlopen($soort)
{
    if(($Button_DS_docs.IsChecked -eq $true -and $soort -eq 'Document')-or ($Button_DS_zaken.IsChecked -eq $true -and $soort -eq 'Zaak'))#nog aanpassen aan checklist-variabelen
    {
        $SDL_split = 1
        
    }else{
        $SDL_split = 0
    }   
#--------------------------------------------
#---------------Events-----------------------
#--------------------------------------------
foreach($global:item_ingeladen in (Select-XML -Xml $global:xml_geladen -XPath '//mse:events' -Namespace @{mse="http://www.nationaalarchief.nl/ToPX/v2.3"}))
    {  
      if($SDL_split -eq 1){   
        Datum_aanpassen_standaard "node" "timestamp" "_tijd" "_datum" #evt. nog dynamisch maken anar std. velden XML
        }
        if($global:vertalen_labels -eq 1){   #nog aanpassen aan checklist vertalenm evt soort checken
            ##nog iets maken
            }
    }
#--------------------------------------------
#---------------Relaties---------------------
#--------------------------------------------
if($soort -eq 'Zaak'){
    foreach($global:item_ingeladen in (Select-XML -Xml $global:xml_geladen -XPath '//mse:relationships' -Namespace @{mse="http://www.nationaalarchief.nl/ToPX/v2.3"}))
    {
      Relaties_verwerken "node" "_vertaald" "_gekoppeld"
    }
    
}
#--------------------------------------------
#---------------Losse velden-----------------
#--------------------------------------------
if($SDL_split -eq 1)
{
    if($soort -eq 'Zaak'){
foreach($global:item_ingeladen in (Select-XML -Xml $global:xml_geladen -XPath '//mse:values' -Namespace @{mse="http://www.nationaalarchief.nl/ToPX/v2.3"}))
    {
        if($SDL_split -eq 1){   
          Datum_aanpassen_standaard "node" "case.date_current" "_tijd" "_datum"  #evt. nog dynamisch maken anar std. velden XML
          Datum_aanpassen_standaard "node" "case.date_of_registration" "_tijd" "_datum"  #evt. nog dynamisch maken anar std. velden XML
          Datum_aanpassen_standaard "node" "case.date_of_registration_full" "_tijd" "_datum"  #evt. nog dynamisch maken anar std. velden XML
          Datum_aanpassen_standaard "node" "case.date_completion" "_tijd" "_datum"  #evt. nog dynamisch maken anar std. velden XML
          Datum_aanpassen_standaard "node" "case.date_completion_full" "_tijd" "_datum"  #evt. nog dynamisch maken anar std. velden XML
          Datum_aanpassen_standaard "node" "case.date_target" "_tijd" "_datum"  #evt. nog dynamisch maken anar std. velden XML
          Datum_aanpassen_standaard "node" "system.current_date" "_tijd" "_datum"  #evt. nog dynamisch maken anar std. velden XML
        }
        if($Button_Rel_koppelen.IsChecked -eq $true) #dit moet zijn aangevinkt, anders is er geen _volledig-variabele aangemaakt bij relaties
        {
          Datum_aanpassen_standaard "Relatie" #Speciale Relatie-call als laatste zaak-actie om de volledige lijst gerelateerde zaken te koppelen
        }
      }
}
if($soort -eq 'Document'){
foreach($global:item_ingeladen in (Select-XML -Xml $global:xml_geladen -XPath '//mse:metadata' -Namespace @{mse="http://www.nationaalarchief.nl/ToPX/v2.3"}))
    {
        if($SDL_split -eq 1){   
          Datum_aanpassen_standaard "node" "date_modified" "_tijd" "_datum"  #evt. nog dynamisch maken anar std. velden XML
          Datum_aanpassen_standaard "node" "date_created" "_tijd" "_datum"  #evt. nog dynamisch maken anar std. velden XML
          Datum_aanpassen_standaard "node.filestore_id" "date_created" "_tijd" "_datum"  #evt. nog dynamisch maken anar std. velden XML
                  }
    }
}
}
#--------------------------------------------
#---------------XML-opslaan------------------
#--------------------------------------------
$opslaan = ""
#Naamgeving bepalen op basis van of een bestand mag worden overschreven of niet
#Bestandsnaam en opslaglocatie bepalen aan de hand van de zaak of documentarray
if($soort -eq 'Zaak')
{
  $bestandsnaam = $global:gevonden_zaken_kort[$global:aantal_zaken]
  $opslaglocatie = $global:gevonden_zaken_dir[$global:aantal_zaken]
}else{
  $bestandsnaam = $global:gevonden_docs_kort[$global:aantal_docs]
  $opslaglocatie = $global:gevonden_docs_dir[$global:aantal_docs]
}
#bestandsnaam samenstellen op basis van of overschrijven wel of niet mag
if ($Button_Overschrijven.IsChecked -eq $true)
{
    $opslaan = "\" + $bestandsnaam + $Txtbox_alt_naam.text + ".xml"
}
else
{
    $opslaan = "\" + $bestandsnaam + ".xml"
}
#Opslaglocatie bepalen op basis van alt_opslag ja of nee
if($Button_Overschrijven.IsChecked -eq $true)
{
  $opslaan = $Loc_alt.text + $opslaan
}else{
  $opslaan = $opslaglocatie + $opslaan
}
$global:xml_geladen.Save($opslaan)
# korte pauze om overlap in het script (n.a.v. evt. trage opslag) te voorkomen
Start-sleep -Millisecond 30
}
function Tobeornottobe
{
#Vraagt de gebruiker of hij/zij door wil gaan met het draaien van het script. Over het algemeen gebruik tijdens het inladen van opties, of als het script een foutmelding tegenkomt welke kan zorgen voor een corrupte verwerking.
Write-Host "Wilt u doorgaan met het draaien van het script? Antwoord 'j' als u wilt doorgaan, antwoord 'n' als u wilt stoppen."
$doorgaan = Read-Host "Doorgaan? (j/n)"
if ($doorgaan -eq "n")
{   $Window.Close()
    throw "Script door gebruiker beëindigd."}
elseif ($doorgaan -eq "j") {
  Write-Host "Script voorgezet..."
}else{
  Write-Host "Verkeerd antwoord. Antwoord met j of n."
  Tobeornottobe
}
}
function Datum_aanpassen_standaard($node,$veldnaam,$suffix_tijd,$suffix_datum)
{
  if($node -ne "Relatie") { #Speciale call om als laatste de relatie_volledig aan de zaak-xml te koppelen
  if ($global:item_ingeladen.$node.$veldnaam -match '(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})')
    {
      #datum in dd-mm-jjjj format
      $datum = '{0:00}-{1:00}-{2:00}' -f [Int]$matches[3], [Int]$matches[2], [Int]$matches[1]
      $element_naam = "" + $veldnaam + $suffix_datum
      $timestamp_2 = $global:xml_geladen.CreateElement($element_naam)
      $timestamp_2.set_InnerText($datum)
      $global:item_ingeladen.$node.AppendChild($timestamp_2)
      #tijd in uu:mm:ss format
      $element_naam = "" + $veldnaam + $suffix_tijd
      $tijd = '{0:00}:{1:00}:{2:00}' -f [Int]$matches[4], [Int]$matches[5], [Int]$matches[6]
      $timestamp_2 = $global:xml_geladen.CreateElement($element_naam)
      $timestamp_2.set_InnerText($tijd)
      $global:item_ingeladen.$node.AppendChild($timestamp_2)
      $global:datumsplits++
    }
  }else{
      $relaties_gekoppeld = $global:xml_geladen.createElement("case.relations_volledig")
      $relaties_gekoppeld.set_InnerText($global:related_cases_volledig)
      $global:item_ingeladen.node.AppendChild($relaties_gekoppeld)
  } 
}
function Relaties_verwerken($node,$suffix_vertaald,$suffix_koppeling)
{
    #Relaties vertalen (als het is aangevinkt)
    if($Button_Rel_vertalen.IsChecked -eq $true)
    {
      if($global:item_ingeladen.$node.relation_type -eq 'child')
      {
        $relatie_vertaling = $global:relatie_child
      }elseif($global:item_ingeladen.$node.relation_type -eq 'parent')
      {
        $relatie_vertaling = $global:relatie_parent
      }elseif($global:item_ingeladen.$node.relation_type -eq 'plain')
      {
        $relatie_vertaling = $global:relatie_plain
      }else {
        Write-Error "FOUTMELDING: Relatie in $($global:zaaknummer) is van onbekende soort!"
        Write-Host "We gebruiken voor nu de definitie van 'plain'"
        $relatie_vertaling = $global:relatie_plain
      }
      $relation_type_vertaald = $global:xml_geladen.CreateElement("relation_type_$($suffix_vertaald)")
      $relation_type_vertaald.set_InnerText($relatie_vertaling)
      $global:item_ingeladen.$node.AppendChild($relation_type_vertaald)
    }
    if($Button_Rel_koppelen.IsChecked -eq $true)
    {
      if($Button_Rel_vertalen.IsChecked -eq $true)
      { 
        $relatie_koppeling = $relatie_vertaling
      }else{
        $relatie_koppeling = $global:item_ingeladen.$node.relation_type
      }
      $relatie_koppeling = "$($global:item_ingeladen.$node.related_object_id) ($($relatie_koppeling))"
      $relatie_gekoppeld = $global:xml_geladen.createElement("relation_volledig")
      $relatie_gekoppeld.set_InnerText($relatie_koppeling)
      $global:item_ingeladen.$node.AppendChild($relatie_gekoppeld)
      $global:relatiesgelegd++
      if($global:related_cases_volledig -eq "")
      {
        $global:related_cases_volledig = $relatie_koppeling
      }else{
        $global:related_cases_volledig = $global:related_cases_volledig + ", " + $relatie_koppeling
      }
    }
}

#-------------------------------------------------------------#
#----Script uitvoer-------------------------------------------#
#-------------------------------------------------------------#
#Start met een aantal controles
Write-Host "SPIRA v1.5"
Write-Host "Opstarten..."
    #Standaard waarden inladen
$Global:XML_standaard_waarden = New-Object -TypeName XML
	$XML_standaard_waarden_pad = $PSScriptRoot + "\SPIRA-StWaarden-V1.5.xml"
	if (-not (Test-Path -Path $XML_standaard_waarden_pad))
	{
    Throw "FOUTMELDING: Standaard waarden-bestand niet gevonden. Plaats het Standaard waarden-bestand in dezelfde map als het script en zorg er voor dat het de volgende syntax heeft: SPIRA-StWaarden-V#.#.xml, waarbij de versienummers identiek zijn aan het SPIRA-scriptbestand." 
	}else{
	Write-Host "Standaard waarden-bestand gevonden. Standaard waarden inladen..."
	}
	$Global:XML_standaard_waarden.Load($XML_standaard_waarden_pad)
$global:datumsplits = 0
#Inladen van de GUI
$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }


$Button_Overschrijven.Add_Checked({Set_active_alts -chk_status '1' $this $_})
$Button_Overschrijven.Add_Unchecked({Set_active_alts -chk_status '0' $this $_})
$Button_Velden_custom.Add_Checked({Set_active_altxml -chk_status '1' $this $_})
$Button_Velden_custom.Add_Unchecked({Set_active_altxml -chk_status '0' $this $_})
$Button_loc_export.Add_Click({Selecteer_pad_export $this $_})
$Button_alt_loc.Add_Click({Selecteer_pad_alt $this $_})
$Button_alt_velden.Add_Click({Selecteer_bestand $this $_})
$Button_WeetZeker.Add_Checked({Activeer_spira -Chk_status '1' $this $_})
$Button_WeetZeker.Add_Unchecked({Activeer_spira -Chk_status '0' $this $_})
$Button_run_spira.Add_Click({Voorman $this $_})

Write-Host "Initiatie compleet, GUI opstarten..."
$Window.ShowDialog()


