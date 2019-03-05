Param (
  [Parameter(mandatory=$true)]
  [string] $InputFile,
  [Parameter(mandatory=$true)]
  [string] $Location
)

$logfile = "license.log"
$debuglog = "debug.log"
# This list is static. Watch ISO ISO 3166 changes at https://www.iso.org/obp/ui/#search/code/
# and update the list as the political scene of the world changes.
$countryCodes = @(`
("Afghanistan","AF"),("�land Islands","AX"),("Albania","AL"),("Algeria","DZ"),("American Samoa","AS"),("Andorra","AD"),("Angola","AO"),("Anguilla","AI"),("Antarctica","AQ"),("Antigua and Barbuda","AG"),`
("Argentina","AR"),("Armenia","AM"),("Aruba","AW"),("Australia","AU"),("Austria","AT"),("Azerbaijan","AZ"),("Bahamas","BS"),("Bahrain","BH"),("Bangladesh","BD"),("Barbados","BB"),`
("Belarus","BY"),("Belgium","BE"),("Belize","BZ"),("Benin","BJ"),("Bermuda","BM"),("Bhutan","BT"),("Bolivia (Plurinational State of)","BO"),("Bonaire, Sint Eustatius and Saba","BQ"),("Bosnia and Herzegovina","BA"),("Botswana","BW"),`
("Bouvet Island","BV"),("Brazil","BR"),("British Indian Ocean Territory","IO"),("Brunei Darussalam","BN"),("Bulgaria","BG"),("Burkina Faso","BF"),("Burundi","BI"),("Cabo Verde","CV"),("Cambodia","KH"),("Cameroon","CM"),`
("Canada","CA"),("Cayman Islands","KY"),("Central African Republic","CF"),("Chad","TD"),("Chile","CL"),("China","CN"),("Christmas Island","CX"),("Cocos (Keeling) Islands","CC"),("Colombia","CO"),("Comoros","KM"),`
("Congo","CG"),("Congo (Democratic Republic of the)","CD"),("Cook Islands","CK"),("Costa Rica","CR"),("C�te d'Ivoire","CI"),("Croatia","HR"),("Cuba","CU"),("Cura�ao","CW"),("Cyprus","CY"),("Czechia","CZ"),`
("Denmark","DK"),("Djibouti","DJ"),("Dominica","DM"),("Dominican Republic","DO"),("Ecuador","EC"),("Egypt","EG"),("El Salvador","SV"),("Equatorial Guinea","GQ"),("Eritrea","ER"),("Estonia","EE"),`
("Ethiopia","ET"),("Falkland Islands (Malvinas)","FK"),("Faroe Islands","FO"),("Fiji","FJ"),("Finland","FI"),("France","FR"),("French Guiana","GF"),("French Polynesia","PF"),("French Southern Territories","TF"),("Gabon","GA"),`
("Gambia","GM"),("Georgia","GE"),("Germany","DE"),("Ghana","GH"),("Gibraltar","GI"),("Greece","GR"),("Greenland","GL"),("Grenada","GD"),("Guadeloupe","GP"),("Guam","GU"),`
("Guatemala","GT"),("Guernsey","GG"),("Guinea","GN"),("Guinea-Bissau","GW"),("Guyana","GY"),("Haiti","HT"),("Heard Island and McDonald Islands","HM"),("Holy See","VA"),("Honduras","HN"),("Hong Kong","HK"),`
("Hungary","HU"),("Iceland","IS"),("India","IN"),("Indonesia","ID"),("Iran (Islamic Republic of)","IR"),("Iraq","IQ"),("Ireland","IE"),("Isle of Man","IM"),("Israel","IL"),("Italy","IT"),`
("Jamaica","JM"),("Japan","JP"),("Jersey","JE"),("Jordan","JO"),("Kazakhstan","KZ"),("Kenya","KE"),("Kiribati","KI"),("Korea (Democratic People's Republic of)","KP"),("Korea (Republic of)","KR"),("Kuwait","KW"),`
("Kyrgyzstan","KG"),("Lao People's Democratic Republic","LA"),("Latvia","LV"),("Lebanon","LB"),("Lesotho","LS"),("Liberia","LR"),("Libya","LY"),("Liechtenstein","LI"),("Lithuania","LT"),("Luxembourg","LU"),`
("Macao","MO"),("Macedonia (the former Yugoslav Republic of)","MK"),("Madagascar","MG"),("Malawi","MW"),("Malaysia","MY"),("Maldives","MV"),("Mali","ML"),("Malta","MT"),("Marshall Islands","MH"),("Martinique","MQ"),`
("Mauritania","MR"),("Mauritius","MU"),("Mayotte","YT"),("Mexico","MX"),("Micronesia (Federated States of)","FM"),("Moldova (Republic of)","MD"),("Monaco","MC"),("Mongolia","MN"),("Montenegro","ME"),("Montserrat","MS"),`
("Morocco","MA"),("Mozambique","MZ"),("Myanmar","MM"),("Namibia","NA"),("Nauru","NR"),("Nepal","NP"),("Netherlands","NL"),("New Caledonia","NC"),("New Zealand","NZ"),("Nicaragua","NI"),`
("Niger","NE"),("Nigeria","NG"),("Niue","NU"),("Norfolk Island","NF"),("Northern Mariana Islands","MP"),("Norway","NO"),("Oman","OM"),("Pakistan","PK"),("Palau","PW"),("Palestine, State of","PS"),`
("Panama","PA"),("Papua New Guinea","PG"),("Paraguay","PY"),("Peru","PE"),("Philippines","PH"),("Pitcairn","PN"),("Poland","PL"),("Portugal","PT"),("Puerto Rico","PR"),("Qatar","QA"),`
("R�union","RE"),("Romania","RO"),("Russian Federation","RU"),("Rwanda","RW"),("Saint Barth�lemy","BL"),("Saint Helena, Ascension and Tristan da Cunha","SH"),("Saint Kitts and Nevis","KN"),("Saint Lucia","LC"),("Saint Martin (French part)","MF"),("Saint Pierre and Miquelon","PM"),`
("Saint Vincent and the Grenadines","VC"),("Samoa","WS"),("San Marino","SM"),("Sao Tome and Principe","ST"),("Saudi Arabia","SA"),("Senegal","SN"),("Serbia","RS"),("Seychelles","SC"),("Sierra Leone","SL"),("Singapore","SG"),`
("Sint Maarten (Dutch part)","SX"),("Slovakia","SK"),("Slovenia","SI"),("Solomon Islands","SB"),("Somalia","SO"),("South Africa","ZA"),("South Georgia and the South Sandwich Islands","GS"),("South Sudan","SS"),("Spain","ES"),("Sri Lanka","LK"),`
("Sudan","SD"),("Suriname","SR"),("Svalbard and Jan Mayen","SJ"),("Swaziland","SZ"),("Sweden","SE"),("Switzerland","CH"),("Syrian Arab Republic","SY"),("Taiwan, Province of China[a]","TW"),("Tajikistan","TJ"),("Tanzania, United Republic of","TZ"),`
("Thailand","TH"),("Timor-Leste","TL"),("Togo","TG"),("Tokelau","TK"),("Tonga","TO"),("Trinidad and Tobago","TT"),("Tunisia","TN"),("Turkey","TR"),("Turkmenistan","TM"),("Turks and Caicos Islands","TC"),`
("Tuvalu","TV"),("Uganda","UG"),("Ukraine","UA"),("United Arab Emirates","AE"),("United Kingdom of Great Britain and Northern Ireland","GB"),("United States of America","US"),("United States Minor Outlying Islands","UM"),("Uruguay","UY"),("Uzbekistan","UZ"),("Vanuatu","VU"),`
("Venezuela (Bolivarian Republic of)","VE"),("Viet Nam","VN"),("Virgin Islands (British)","VG"),("Virgin Islands (U.S.)","VI"),("Wallis and Futuna","WF"),("Western Sahara","EH"),("Yemen","YE"),("Zambia","ZM"),("Zimbabwe","ZW")`
)

$plans=@{
"AAD_BASIC" = "Azure Active Directory Basic"
"AX_DATABASE_STORAGE" = "--- not availablel for selection in the Portal ---"
"CRMENTERPRISE" = "Microsoft Dynamics CRM Online Enterprise"
"CRMPLAN2" = "Microsoft Dynamics CRM Online Basic"
"CRMSTANDARD" = "Microsoft Dynamics CRM Online Professional"
"CRMSTORAGE" = "--- not availablel for selection in the Portal ---"
"CRMTESTINSTANCE" = "--- not availablel for selection in the Portal ---"
"DESKLESSPACK" = "Office 365 F1"
"DYN365_ENTERPRISE_PLAN1" = "Dynamics 365 Customer Engagement Plan Enterprise Edition"
"DYN365_ENTERPRISE_SALES" = "Dynamics 365 for Sales Enterprise Edition"
"DYN365_ENTERPRISE_TEAM_MEMBERS" = "Dynamics 365 for Team Members Enterprise"
"ECAL_SERVICES" = "ECAL Services (EOA, EOP, DLP)"
"EMS" = "Enterprise Mobility + Security E3"
"EMSPREMIUM" = "Enterprise Mobility + Security E5"
"ENTERPRISEPACK" = "Office 365 Enterprise E3"
"ENTERPRISEPREMIUM" = "Office 365 Enterprise E5"
"ENTERPRISEPREMIUM_NOPSTNCONF" = "Office 365 Enterprise E5 without Audio Conferencing (These licenses do not need to be individually assigned)"
"EOP_ENTERPRISE" = "Exchange Online Protection (These licenses do not need to be individually assigned)"
"EXCHANGESTANDARD" = "Exchange Online (Plan 1)"
"FLOW_FREE" = "Microsoft Flow Free"
"INTUNE_A" = "Intune"
"MCOEV" = "Phone System"
"MCOMEETADV" = "Audio Conferencing"
"POWER_BI_PRO" = "Power BI Pro"
"POWER_BI_STANDARD" = "Power BI (free)"
"POWERAPPS_INDIVIDUAL_USER" = "PowerApps and Logic Flows"
"PROJECTCLIENT" = "--- not availablel for selection in the Portal ---"
"SPE_E3" = "Microsoft 365 E3"
"STANDARDPACK" = "Office 365 Enterprise E1"
"VISIOCLIENT" = "Visio Online Plan 2"
}

$services = @{
"AAD_BASIC" = "Azure Active Directory Basic"
"AAD_PREMIUM" = "Azure Active Directory Premium Plan 1"
"ADALLOM_S_DISCOVERY" = "Cloud App Security Discovery"
"ADALLOM_S_O365" = "Office 365 Cloud App Security"
"ATP_ENTERPRISE" = "Exchange Online Advanced Threat Protection (These licenses do not need to be individually assigned)"
"BI_AZURE_P0" = "Power BI (free)"
"BI_AZURE_P2" = "Power BI Pro"
"BPOS_S_TODO_1" = "To-Do (Plan 1)"
"BPOS_S_TODO_2" = "To-Do (Plan 2)"
"BPOS_S_TODO_3" = "To-Do (Plan 3)"
"BPOS_S_TODO_FIRSTLINE" = "To-Do (Firstline)"
"CRM_FIELD_SERVICE_ADDON" = "Microsoft Dynamics CRM Online - Field Service User Add-On"
"CRM_PROJECT_SERVICE_ADDON" = "Microsoft Dynamics CRM Online - Project Service User Add-On"
"CRMENTERPRISE" = "Microsoft Dynamics CRM Online Enterprise"
"CRMPLAN2" = "Microsoft Dynamics CRM Online Basic"
"CRMSTANDARD" = "Microsoft Dynamics CRM Online Professional"
"CRMSTORAGE" = "--- not availablel for selection in the Portal ---"
"CRMTESTINSTANCE" = "--- not availablel for selection in the Portal ---"
"Deskless" = "Microsoft StaffHub"
"DMENTERPRISE" = "Microsoft Dynamics Marketing Online Enterprise"
"DYN365_ENTERPRISE_P1" = "Dynamics 365 Customer Engagement Plan"
"DYN365_ENTERPRISE_SALES" = "Dynamics 365 for Sales"
"DYN365_Enterprise_Talent_Attract_TeamMember" = "Dynamics 365 for Talent - Attract Experience Team Member"
"DYN365_Enterprise_Talent_Onboard_TeamMember" = "Dynamics 365 for Talent - Onboard Experience"
"DYN365_ENTERPRISE_TEAM_MEMBERS" = "Dynamics 365 for Team Members"
"Dynamics 365 for Operations Storage Plan" = "--- not availablel for selection in the Portal ---"
"Dynamics_365_for_Operations_Team_members" = "Dynamics_365_for_Operations_Team_members"
"Dynamics_365_for_Retail_Team_members" = "Dynamics 365 for Retail Team members"
"Dynamics_365_for_Talent_Team_members" = "Dynamics 365 for Talent Team members"
"EOP_ENTERPRISE" = "Exchange Online Protection (These licenses do not need to be individually assigned)"
"EOP_ENTERPRISE_PREMIUM" = "Exchange Enterprise CAL Services (EOP, DLP) (These licenses do not need to be individually assigned)"
"EQUIVIO_ANALYTICS" = "Office 365 Advanced eDiscovery"
"EXCHANGE_ANALYTICS" = "Microsoft MyAnalytics"
"EXCHANGE_S_ARCHIVE" = "Exchange Online Archiving for Exchange Server"
"EXCHANGE_S_DESKLESS" = "Exchange Online Kiosk"
"EXCHANGE_S_ENTERPRISE" = "Exchange Online (Plan 2)"
"EXCHANGE_S_FOUNDATION" = "--- not availablel for selection in the Portal ---"
"EXCHANGE_S_STANDARD" = "Exchange Online (Plan 1)"
"FLOW_DYN_APPS" = "Flow for Dynamics 365"
"FLOW_DYN_P2" = "Flow for Dynamics 365"
"FLOW_DYN_TEAM" = "Flow for Dynamics 365"
"FLOW_O365_P1" = "Flow for Office 365"
"FLOW_O365_P2" = "Flow for Office 365"
"FLOW_O365_P3" = "Flow for Office 365"
"FLOW_O365_S1" = "Flow for Office 365 K1"
"FLOW_P2_VIRAL" = "Microsoft Flow Free"
"FORMS_PLAN_E1" = "Microsoft Forms (Plan E1)"
"FORMS_PLAN_E3" = "Microsoft Forms (Plan E3)"
"FORMS_PLAN_E5" = "Microsoft Forms (Plan E5)"
"FORMS_PLAN_K" = "Microsoft Forms (Plan K)"
"INTUNE_A" = "Intune A Direct"
"INTUNE_O365" = "Mobile Device Management for Office 365 (These licenses do not need to be individually assigned)"
"LOCKBOX_ENTERPRISE" = "Customer Lockbox"
"MCOEV" = "Phone System"
"MCOIMP" = "Skype for Business Online (Plan 1)"
"MCOMEETADV" = "Audio Conferencing"
"MCOSTANDARD" = "Skype for Business Online (Plan 2)"
"MDM_SALES_COLLABORATION" = "Microsoft Dynamics Marketing Sales Collaboration - Eligibility criteria apply"
"MFA_PREMIUM" = "Azure Multi-Factor Authentication "
"NBENTERPRISE" = "Microsoft Social Engagement Enterprise"
"NBPROFESSIONALFORCRM" = "Microsoft Social Engagement Professional � Eligibility Criteria apply"
"OFFICESUBSCRIPTION" = "Office 365 ProPlus"
"ONEDRIVE_BASIC" = "OneDrive for Business Basic"
"PARATURE_ENTERPRISE" = "Parature Enterprise"
"POWERAPPS_DYN_APPS" = "PowerApps for Dynamics 365"
"POWERAPPS_DYN_P2" = "PowerApps for Dynamics 365"
"POWERAPPS_DYN_TEAM" = "PowerApps for Dynamics 365"
"POWERAPPS_O365_P1" = "PowerApps for Office 365"
"POWERAPPS_O365_P2" = "PowerApps for Office 365"
"POWERAPPS_O365_P3" = "PowerApps for Office 365"
"POWERAPPS_O365_S1" = "PowerApps for Office 365 K1"
"POWERAPPSFREE" = "Microsoft PowerApps"
"POWERFLOWSFREE" = "Logic Flows"
"POWERVIDEOSFREE" = "Microsoft Power Videos Basic"
"PROJECT_CLIENT_SUBSCRIPTION" = "Project Online Desktop Client"
"PROJECT_ESSENTIALS" = "Project Online Essentials"
"PROJECTWORKMANAGEMENT" = "Microsoft Planner"
"RMS_S_ENTERPRISE" = "Azure Rights Management"
"RMS_S_PREMIUM" = "Azure Information Protection Plan 1"
"SHAREPOINT_PROJECT" = "Project Online Service"
"SHAREPOINTDESKLESS" = "SharePoint Online Kiosk"
"SHAREPOINTENTERPRISE" = "SharePoint Online (Plan 2)"
"SHAREPOINTSTANDARD" = "SharePoint Online (Plan 1)"
"SHAREPOINTWAC" = "Office Online"
"STREAM_O365_E1" = "Stream for Office 365"
"STREAM_O365_E3" = "Stream for Office 365"
"STREAM_O365_E5" = "Stream for Office 365"
"STREAM_O365_K" = "Stream for Office 365 K1"
"SWAY" = "Sway"
"TEAMS1" = "Microsoft Teams"
"THREAT_INTELLIGENCE" = "Office 365 Threat Intelligence"
"VISIO_CLIENT_SUBSCRIPTION" = "Visio Pro for Office 365"
"VISIOONLINE" = "Visio Online"
"WIN10_PRO_ENT_SUB" = "Windows 10 Enterprise"
"YAMMER_ENTERPRISE" = "Yammer Enterprise"
}

# Regexp UPN/email validation pattern credit:
# https://www.regular-expressions.info/email.html
$UPNpattern = "\A[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\z"
$invalidPlans = @("AX_DATABASE_STORAGE","CRMTESTINSTANCE","CRMSTORAGE","EOP_ENTERPRISE")

function IIF-Plans {
  Param (
    [string]$plan
  )
  if ($plans[$plan]) {
    $tmpResult = $plans[$plan]
  }
  else {
    $tmpResult = $plan + " (no friendly name found)"
  }
  return $tmpResult
}

function IIF-Services {
  Param (
    [string]$service
  )
  if ($services[$service]) {
    $tmpResult = $services[$service]
  }
  else {
    $tmpResult = $service + " (no friendly name found)"
  }
  return $tmpResult
}

function Get-OperationSelection {
  Param (
    [array]$Collection,
    [string]$Message
  )
  $tmpCollection = $Collection[0]
  $tmpMessage = $Collection[1]
  Write-Host $tmpMessage -ForegroundColor White
  for ($i=0; $i -lt $tmpCollection.count; $i++) {
    Write-Host ($i.ToString()).PadLeft(2," "),$tmpCollection[$i]
  }
  $promptMessage = "Selezionare un opzione e premere INVIO, o X per uscire"
  do {
    $selection = Read-Host -Prompt $promptMessage
    $promptMessage = "Selezione non valida. Selezionare un opzione e premere INVIO, o X per uscire"
  }
  until (($selection -in "0"..($tmpCollection.count-1).ToString()) -or ($selection.ToUpper() -eq "X"))
  return $selection
}

function Get-PlanSelection {
  Param (
    [array]$Collection,
    [string]$Message
  )
  $tmpCollection = $Collection[0]
  $tmpMessage = $Collection[1]
  Write-Host $tmpMessage -ForegroundColor White
  for ($i=0; $i -lt $tmpCollection.count; $i++) {
    $tmpPlan = IIF-Plans $tmpCollection[$i]
    Write-Host ($i.ToString()).PadLeft(2," "),$tmpPlan
  }
  $promptMessage = "Selezionare un opzione e premere INVIO, o X per uscire"
  do {
    $selection = Read-Host -Prompt $promptMessage
    $promptMessage = "Selezione non valida. Selezionare un opzione e premere INVIO, o X per uscire"
  }
  until (($selection -in "0"..($tmpCollection.count-1).ToString()) -or ($selection.ToUpper() -eq "X"))
  return $selection
}

function Get-ServiceSelection {
  Param (
    [array]$Collection,
    [string]$Message
  )
  $tmpCollection = $Collection[0]
  $tmpMessage = $Collection[1]
  Write-Host $tmpMessage -ForegroundColor White
  for ($i=0; $i -lt $tmpCollection.count; $i++) {
    $tmpService = IIF-Services $tmpCollection[$i]
#    Write-Host ($i.ToString()).PadLeft(2," "),$services[$tmpCollection[$i]]
    Write-Host ($i.ToString()).PadLeft(2," "),$tmpService
  }
  $promptMessage = "Selezionare un opzione e premere INVIO, o X per uscire"
  do {
    $selection = Read-Host -Prompt $promptMessage
    $promptMessage = "Selezione non valida. Selezionare un opzione e premere INVIO, o X per uscire"
  }
  until (($selection -in "0"..($tmpCollection.count-1).ToString()) -or ($selection.ToUpper() -eq "X"))
  return $selection
}

# Builds the DisabledOptions object where there is only a single service plan is enabled.
# Use it when a new plan is assigned and this is the only service plan enabled.
#########################################################################################
function New-LicenseOptionsSingle {
  Param (
    $Sku,
    $Service
  )
  $disabledServices = @()
  For ($i=0; $i -lt ($Sku.servicestatus).count; $i++) {
    if ((($Sku.servicestatus)[$i]).ServicePlan.ServiceName -ne $Service) {
      $disabledServices += (($Sku.servicestatus)[$i]).ServicePlan.ServiceName
    }
  }
  $tmpLicenseOptions = New-MsolLicenseOptions -AccountSkuId $colAccountSku[$selectedSkuIndex].AccountSkuId -DisabledPlans $disabledServices
  return $tmpLicenseOptions
}

# Turns on a Service Plan when other service plans may also be on.
# Use it when a plan is already assigned and there are enabled service plans.
#########################################################################################
function New-LicenseOptionsMultiple {
  Param (
    $Sku,
    $Service
  )
  $disabledServices = @()
  For ($i=0; $i -lt ($Sku.servicestatus).count; $i++) {
    if (((($Sku.servicestatus)[$i]).ServicePlan.ServiceName -ne $Service) -and ((($Sku.servicestatus)[$i]).ProvisioningStatus -eq "Disabled")) {
      $disabledServices += (($Sku.servicestatus)[$i]).ServicePlan.ServiceName
    }
  }
  $tmpLicenseOptions = New-MsolLicenseOptions -AccountSkuId $colAccountSku[$selectedSkuIndex].AccountSkuId -DisabledPlans $disabledServices
  return $tmpLicenseOptions
}

function New-LicenseOptionsDisableService {
  Param (
    $Sku,
    $Service
  )
  $disabledServices = @()
  For ($i=0; $i -lt ($Sku.servicestatus).count; $i++) {
    if (((($Sku.servicestatus)[$i]).ServicePlan.ServiceName -eq $Service) -or ((($Sku.servicestatus)[$i]).ProvisioningStatus -eq "Disabled")) {
      $disabledServices += (($Sku.servicestatus)[$i]).ServicePlan.ServiceName
    }
  }
  $tmpLicenseOptions = New-MsolLicenseOptions -AccountSkuId $colAccountSku[$selectedSkuIndex].AccountSkuId -DisabledPlans $disabledServices
  return $tmpLicenseOptions
}


###########################
#                         #
# Script main entry point #
#                         #
###########################

# Delete previous log
#######################################################

if (Test-Path $logfile) {
  Remove-Item $logfile
}

if (Test-Path $debuglog) {
  Remove-Item $debuglog
}

# Are MSOL commands available?
#######################################################
if (-not(Get-Command Get-MsolDomain -ErrorAction SilentlyContinue)) {
  Write-Host "Errore: " -ForegroundColor Red -NoNewLine
  Write-Host "Installare il modulo powershell per Azure Active Directory."
  Write-Host "Vdere " -NoNewLine
  Write-Host "https://technet.microsoft.com/en-us/library/dn975125.aspx" -ForegroundColor Cyan -NoNewLine
  Write-Host " per dettagli."
  Write-Host "Uscita."
  Exit
}


# Are we connected to a tenant?
#######################################################
if (-not(Get-MsolDomain -ErrorAction SilentlyContinue)) {
  Try {
    $cred = Get-Credential -ErrorAction Stop
  }
  Catch {
    Write-Host "Errore: " -ForegroundColor Red -NoNewLine
    Write-Host "Non sono state fornite credenziali. Uscita."
    Exit
  }
  Connect-MsolService -Credential $cred -ErrorAction SilentlyContinue
  if (-not(Get-MsolDomain -ErrorAction SilentlyContinue)) {
    Write-Host "Errore: " -ForegroundColor Red -NoNewline
    Write-Host "Credenziali non valide, impossibile collegarrsi a Office 365." Exiting.
    Exit
  }
}

$tmpOrg = (Get-MsolCompanyInformation).DisplayName
Write-Host "Sei attualmente collegato a " -NoNewLine
Write-Host $tmpOrg -ForegroundColor Green


# Check if input file exists. Exit otherwise.
#######################################################
If (-not(Test-Path $InputFile)) {
  Write-Host "Errore: " -ForegroundColor Red -NoNewLine
  Write-Host "Nessun file indicato. uscita."
  Exit
}

# Check if country code has been provided. Exit otherwise.
##########################################################
$location = $location.ToUpper()
$match = $false
For ($i=0; $i -lt $countryCodes.count; $i++) {
  if ($countryCodes[$i][1] -eq $Location) {
    $match = $True
    break
  }
}
if ($match) {
  $strLocation = $countryCodes[$i][1] + ", " + $countryCodes[$i][0]
}
else {
  Write-Host "Errore: " -ForegroundColor Red -NoNewLine
  Write-Host "Codice Nazione non valido. Inserire un codice nazione a due cifre (Es.IT)."
  Write-Host "Riferimenti: " -NoNewLine
  Write-Host "https://technet.microsoft.com/en-us/library/dn771770.aspx" -ForegroundColor Cyan
  Write-Host "Verifica" -NoNewLine
  Write-Host "https://www.iso.org/obp/ui/#search/code/" -ForegroundColor Cyan -NoNewLine
  Write-Host " per un elenco Codici Nazione ISO 3166 Alpha-2."
  Exit
}

Write-Host "Localizzazione: " -NoNewLine
Write-Host $strLocation -ForegroundColor Green
Write-Host "File dati: " -NoNewLine
Write-Host $InputFile -ForegroundColor Green

# Enable or Disable licences:
#######################################################
$colSessionOperation = @("Disabilita Licenza","Abilita Licenza")
$selection = Get-OperationSelection $colSessionOperation, "Selezionare il tipo di operazione:"
if ($selection -eq "X") {
  Write-Host "Uscita."
  exit
}
$selectedLicenseOperationIndex = $selection.ToInt32($null)
Write-Host "Hai Selezionato: " -NoNewLine
Write-Host $colSessionOperation[$selectedLicenseOperationIndex] -ForegroundColor Green

# Select subscription:
#######################################################
# See https://powershellone.wordpress.com/2015/08/04/finding-the-index-of-an-object-within-an-array-by-property-value-using-powershell/
$colAccountSku = [Collections.Generic.List[Object]](Get-MsolAccountSku)
$selection = Get-PlanSelection $colAccountSku.SkuPartNumber, "Selezionare la sottoscrizione:"
if ($selection -eq "X") {
  Write-Host "Uscita."
  exit
}
$selectedSkuIndex = $selection.ToInt32($null)
Write-Host "Hai selezionato: " -NoNewLine
Write-Host $colAccountSku[$selectedSkuIndex].SkuPartNumber -ForegroundColor Green

# Select service plan within the selected subscription:
#######################################################
[array]$colSkuServices = $colAccountSku[$selectedSkuIndex].ServiceStatus.ServicePlan.ServiceName
$selection = Get-ServiceSelection $colSkuServices, "Seleziona il tuo piano:"
if ($selection -eq "X") {
  Write-Host "Uscita."
  exit
}
$selectedSkuServiceIndex = $selection.ToInt32($null)
Write-Host "Hai Selezionato: " -NoNewLine
Write-Host $colSkuServices[$selectedSkuServiceIndex] -ForegroundColor Green

Write-Host "Sommaario Selezioni:" -ForegroundColor White
Write-Host "  Operazione:    " -NoNewLine
Write-Host $colSessionOperation[$selectedLicenseOperationIndex] -ForegroundColor Green
Write-Host "  Sottoscrizione: " -NoNewLine
$tmpPlan = IIF-Plans $colAccountSku[$selectedSkuIndex].SkuPartNumber
Write-Host $tmpPlan -ForegroundColor Green
Write-Host "  Piano: " -NoNewLine
$tmpService = IIF-Services $colSkuServices[$selectedSkuServiceIndex]
Write-Host $tmpService -ForegroundColor Green

if ($invalidPlans -contains $colAccountSku[$selectedSkuIndex].SkuPartNumber) {
  $tmpString = $colAccountSku[$selectedSkuIndex].SkuPartNumber
  Write-Host "Errore: " -ForegroundColor Red -NoNewLine
  Write-Host "Il piano " -NoNewLine
  Write-Host $tmpString -NoNewLine -ForegroundColor White
  Write-Host " non può essere assegnato esplicitamente. Uscita."
  Exit
}

$promptMessage = "Sto per apportare le modifiche. Procedere? "
do {
  Write-Host $promptmessage -ForegroundColor Yellow -NoNewLine
  $selection = (Read-Host -Prompt "[S]i, [N]o").ToUpper()
  $promptMessage = "Selezione non valida. Sto per apportare le modifiche. Procedere? "
}
until ($selection -in ("S","N"))
if ($selection -eq "N") {
  Write-Host "Uscita."
  exit
}
Write-Host "Hai Selezionato: " -NoNewLine
Write-Host $selection -ForegroundColor Green

"Avvio registrazionee: " + (Get-Date).ToString() | Out-File $logfile -Append
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
$msgOperation = "Operazione:    " + $colSessionOperation[$selectedLicenseOperationIndex]
$msgSubscription = "Sottoscrizione: " + $colAccountSku[$selectedSkuIndex].SkuPartNumber
$msgServicePlan = "Piano: " + $colSkuServices[$selectedSkuServiceIndex]
$msgOperation | Out-File $logfile -Append
$msgSubscription | Out-File $logfile -Append
$msgServicePlan | Out-File $logfile -Append
"===================================" | Out-File $logfile -Append

# Getting AccountSKUs that have the selected service
#######################################################
# $colSKUsWithServicePlan stores SKUs that have the selected service.
$colSKUsWithServicePlan = @()
ForEach ($AccountSku in $colAccountSku) {
  For ($i=0; $i -lt $AccountSku.ServiceStatus.Count; $i++) {
    if ($AccountSku.ServiceStatus[$i].ServicePlan.ServiceName -eq $colSkuServices[$selectedSkuServiceIndex]) {
      $colSKUsWithServicePlan += $AccountSku
    }
  }
}

# Getting selected account SKU available licences
#######################################################
$tmpSKU = Get-MsolAccountSku | ?{$_.SkuPartNumber -eq $colAccountSku[$selectedSkuIndex].SkuPartNumber}
$availableLicences = $tmpSku.ActiveUnits - $tmpSku.ConsumedUnits

Write-Host "Sto leggendo gli utenti dal file... " -NoNewLine
$users = @(Get-Content $InputFile)
$usercount = $users.count
$strUsercount = $usercount.ToString()
Write-Host "$usercount utent(i)" -Foreground White
Write-Host "Elaborazione... " -NoNewLine
###
# Credit for manipulating cursor position goes to real_parbold
# https://www.reddit.com/r/PowerShell/comments/34d2sk/move_cursor_to_beginning_of_line/
###
$myX = $Host.UI.RawUI.CursorPosition.X
$myY = $Host.UI.RawUI.CursorPosition.Y

# Initialise statistical counters
#######################################################
$counterSuccess = 0            # SUCCESS license assignment
$counterFailed = 0             # FAIL license assignment, e.g. invalid UPN format or user not found
$counterNoChange = 0           # NOCHANGE license assignment, e.g. user already has a license
$counterNoAvailableLicense = 0 # NOAVAILABLELICENSE license assignment, e.g. user already has a license

# Start processing users
#######################################################
$counter = 1
ForEach ($user in $users) {
  $user = $user.trim()
  $status = $counter.ToString() + " / " + $strUsercount
  $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates $myX,$myY
  Write-Host $status -ForegroundColor White -NoNewLine

# Is the user in the correct format?
  if ($user -notmatch $UPNpattern) {
    $errString = "$user ERRORE Formato UPN non corretto!"
    $errString | Out-File $logfile -Append
    $counterFailed ++
  }
  else {
# Does the user exist?
    $noUser = $false
    try {
      $msolUser = get-msoluser -UserPrincipalName $user -ErrorAction Stop
    }
    catch {
      $noUser = $true
    }
    if ($noUser) {
      $errString = "$user ERRORE Utente non valido!"
      $errString | Out-File $logfile -Append
      $counterFailed ++
    }

    else {
      if ($msoluser.isLicensed) {
#       User is licensed. Will sort out what to do based on what already has.
#       Get user's assigned plans
        $assignedPlans = $msolUser.Licenses
        Switch ($selectedLicenseOperationIndex) {
          0 {    # We are disabling a service
            # User does not have plan. Move on to next user.
            # User does not have service within plan. Move on to next user.
            # User has plan and service. Disable regardless of its current state.
            $licenses = $msolUser.Licenses
            $hasPlan = $false
            For ($i=0; $i -lt $licenses.count; $i++) {
              if ($licenses[$i].AccountSku.SkuPartNumber -eq $colAccountSku[$selectedSkuIndex].SkuPartNumber) {
                $usrSkuIndex = $i  # Index of the plan; check $usrServiceIndex below
                $hasPlan = $true
              }
            }
            if (-not($hasPlan)) {
              $tmpString = $colAccountSku[$selectedSkuIndex].SkuPartNumber
              $errString = "$user Nessuna modifica l'utente non ha $tmpString un piano. Nessuna modifica eseguita."
              $counterNoChange ++
            }
            else {
              $hasPlanAndService = $false
              For ($i=0; $i -lt ($licenses[$usrSkuIndex].ServiceStatus).count; $i++) {
                if (($licenses[$usrSkuIndex].ServiceStatus[$i].ServicePlan.ServiceName -eq $colSkuServices[$selectedSkuServiceIndex]) -and `
                  ($licenses[$usrSkuIndex].ServiceStatus[$i].ProvisioningStatus -ne "Disabled")) {
                  $hasPlanAndService = $True
                  $usrServiceIndex = $i  # Index of the service within the plan identified by $usrSkuIndex
                }
              }
              Switch ($hasPlanAndService) {
                $false {
                  $tmpString1 = $colAccountSku[$selectedSkuIndex].SkuPartNumber
                  $tmpString2 = $colSkuServices[$selectedSkuServiceIndex]
                  $errString = "$user Nessuna modifica l'utente non ha un servizio $tmpString2 abilitato nel piano $tmpString1.  Nessuna modifica eseguita."
                  $counterNoChange ++
                }
                $true {
                  $tmpString1 = $licenses[$usrSkuIndex].AccountSku.SkuPartNumber
                  $tmpString2 = $licenses[$usrSkuIndex].ServiceStatus[$usrServiceIndex].ServicePlan.ServiceName
                  $licenseOptions = New-LicenseOptionsDisableService -Sku $licenses[$usrSkuIndex] -Service $tmpString2
                  Try {
                    Set-MsolUserLicense -UserPrincipalName $user -LicenseOptions $licenseOptions -ErrorAction Stop
                    $errString = "$user Completato Servizio $tmpString2 nel piano $tmpString1 è stato disabilitato."
                    # Checking whether there are any services that aren't "Disabled".
                    # Will delete the license if all services are "Disabled".
                    $updatedLicenses = (get-msoluser -UserPrincipalName $user).Licenses
                    $tmpAllServicesCount = ($updatedLicenses[$usrSkuIndex].ServiceStatus).count
                    $tmpPendingActivationServicesCount = ($updatedLicenses[$usrSkuIndex].ServiceStatus | ?{$_.ProvisioningStatus -eq "PendingActivation"}).count
                    $tmpDisabledServicesCount = ($updatedLicenses[$usrSkuIndex].ServiceStatus | ?{$_.ProvisioningStatus -eq "Disabled"}).count
                    if (($tmpPendingActivationServicesCount + $tmpDisabledServicesCount) -eq $tmpAllServicesCount) {
                      Set-MsolUserLicense -UserPrincipalName $user -RemoveLicense $licenses[$usrSkuIndex].AccountSkuId -ErrorAction Stop
                      $errString = "$errString Il Piano $tmpstring1 è stato rimosso dall'utente."
                    }
                    $counterSuccess ++
                  }
                  Catch {
                    "==================== ERROR ====================" | Out-File $debuglog -Append
                    "Affected user: $user" | Out-File $debuglog -Append
                    $dbgMsg = "Command: " + ($_.InvocationInfo.Line).Trim()
                    $dbgMsg | Out-File $debuglog -Append
                    $dbgMsg = "Line: " + $_.InvocationInfo.ScriptLineNumber
                    $dbgMsg | Out-File $debuglog -Append
                    $dbgMsg = "Message: " + $_.Exception.Message
                    $dbgMsg | Out-File $debuglog -Append
                    $errString = "$user FAIL $($_.Exception.Message)"
                    $counterFailed ++
                  }
                }
              }
            }



          }
          1 {    # We are enabling a service
#           Does the user have the selected plan?
            $licenses = $msolUser.Licenses
            $hasPlan = $false
            For ($i=0; $i -lt $licenses.count; $i++) {
              if ($licenses[$i].AccountSku.SkuPartNumber -eq $colAccountSku[$selectedSkuIndex].SkuPartNumber) {
                $usrSkuIndex = $i
                $hasPlan = $true
              }
            }
#           Does the user already have the service enabled?
            $hasService = $false
            For ($i=0; $i -lt $licenses.count; $i++) {
              For ($j=0; $j -lt ($licenses[$i].ServiceStatus).count; $j++) {
                if (($licenses[$i].ServiceStatus[$j].ServicePlan.ServiceName -eq $colSkuServices[$selectedSkuServiceIndex]) `
                -and ($licenses[$i].ServiceStatus[$j].ProvisioningStatus -eq "Success")) {
                  $hasService = $true
                  $usrPlanAlreadyHasService = $licenses[$i].AccountSku.SkuPartNumber
                }
              }
            }
            if ($hasService) {
#             User already has the selected service enabled in one of the assigned plans. No changes, moving on to the next user.
              $errString = "$user Nessuna modifica il Servizio "+ $colSkuServices[$selectedSkuServiceIndex] + " è già nel piano " + $usrPlanAlreadyHasService + ". Nessuna modifica eseguita."
              $counterNoChange ++
            }
            else {
              if ($hasPlan) {
#               User already has the plan for the selected service assigned. It will be updated.
                $licenseOptions = New-LicenseOptionsMultiple -Sku $licenses[$usrSkuIndex] -Service $colSkuServices[$selectedSkuServiceIndex]
                Try {
                  Set-MsolUserLicense -UserPrincipalName $user -LicenseOptions $licenseOptions -ErrorAction Stop
                  $errString = "$user Comepletato il Servizio  "+ $colSkuServices[$selectedSkuServiceIndex] + " è stato abilitato nel piano " + $colAccountSku[$selectedSkuIndex].SkuPartNumber
                  $counterSuccess ++
                }
                Catch {
                  "==================== ERROR ====================" | Out-File $debuglog -Append
                  "Affected user: $user" | Out-File $debuglog -Append
                  $dbgMsg = "Command: " + ($_.InvocationInfo.Line).Trim()
                  $dbgMsg | Out-File $debuglog -Append
                  $dbgMsg = "Line: " + $_.InvocationInfo.ScriptLineNumber
                  $dbgMsg | Out-File $debuglog -Append
                  $dbgMsg = "Message: " + $_.Exception.Message
                  $dbgMsg | Out-File $debuglog -Append
                  $errString = "$user FAIL $($_.Exception.Message)"
                  $counterFailed ++
                }
              }
              else {
#               User has no plan for the selected service. It will be assigned and the selected service enabled.
                if ($availableLicences -gt 0) {
                  $licenseOptions = New-LicenseOptionsSingle -Sku $colAccountSku[$selectedSkuIndex] -Service $colSkuServices[$selectedSkuServiceIndex]
                  Try {
                    Set-MsolUserLicense -UserPrincipalName $user -AddLicenses $colAccountSku[$selectedSkuIndex].AccountSkuId -LicenseOptions $licenseOptions -ErrorAction Stop
                    $errString = "$user Completato il Servizio "+ $colSkuServices[$selectedSkuServiceIndex] + " è stato abilitato nel piano " + $colAccountSku[$selectedSkuIndex].SkuPartNumber
                    $counterSuccess ++
                    $availablelicenses --
                  }
                  Catch {
                    "==================== ERROR ====================" | Out-File $debuglog -Append
                    "Affected user: $user" | Out-File $debuglog -Append
                    $dbgMsg = "Command: " + ($_.InvocationInfo.Line).Trim()
                    $dbgMsg | Out-File $debuglog -Append
                    $dbgMsg = "Line: " + $_.InvocationInfo.ScriptLineNumber
                    $dbgMsg | Out-File $debuglog -Append
                    $dbgMsg = "Message: " + $_.Exception.Message
                    $dbgMsg | Out-File $debuglog -Append
                    $errString = "$user FAIL $($_.Exception.Message)"
                    $counterFailed ++
                  }
                }
                else {
                  $errString = "$user NOAVAILABLELICENSE Nessuna licenza disponibile. Nessuna modifica eseguita."
                  $counterNoAvailableLicense ++
                }
              }
            }
          }
        }
        $errString | Out-File $logfile -Append
      } # ELSE of MSOLUser is licensed check

      else {
#       User is not licensed.
#       The selected plan will be assigned and the service enabled.
        if ($availableLicences -gt 0) {
          $disabledServices = @()
          For ($i=0; $i -lt ($colAccountSku[$selectedSkuIndex].servicestatus).count; $i++) {
            if ((($colAccountSku[$selectedSkuIndex].servicestatus)[$i]).ServicePlan.ServiceName -ne $colSkuServices[$selectedSkuServiceIndex]) {
              $disabledServices += (($colAccountSku[$selectedSkuIndex].servicestatus)[$i]).ServicePlan.ServiceName
            }
          }
          $licenseOptions = New-MsolLicenseOptions -AccountSkuId $colAccountSku[$selectedSkuIndex].AccountSkuId -DisabledPlans $disabledServices
          Try {
            Set-MsolUser -UserPrincipalName $user -UsageLocation $location -ErrorAction Stop
            Set-MsolUserLicense -UserPrincipalName $user -AddLicenses $colAccountSku[$selectedSkuIndex].AccountSkuId
            Set-MsolUserLicense -UserPrincipalName $user -LicenseOptions $licenseOptions
            $errString = "$user Completato il servizio "+ $colSkuServices[$selectedSkuServiceIndex] + " è stato abilitato nel piano " + $colAccountSku[$selectedSkuIndex].SkuPartNumber
            $availablelicenses --
            $counterSuccess ++
          }
          Catch {
            "==================== ERROR ====================" | Out-File $debuglog -Append
            "Affected user: $user" | Out-File $debuglog -Append
            $dbgMsg = "Command: " + ($_.InvocationInfo.Line).Trim()
            $dbgMsg | Out-File $debuglog -Append
            $dbgMsg = "Line: " + $_.InvocationInfo.ScriptLineNumber
            $dbgMsg | Out-File $debuglog -Append
            $dbgMsg = "Message: " + $_.Exception.Message
            $dbgMsg | Out-File $debuglog -Append
            $errString = "$user FAIL $($_.Exception.Message)"
            $counterFailed ++
          }
        }
        else {
          $errString = "$user NOAVAILABLELICENSE Nessuna licenza disponibile. Nessuna modifica eseguita."
          $counterNoAvailableLicense ++
        }
        $errString | Out-File $logfile -Append
      } # ELSE of MSOLUser is NOT licensed check
    }
    $counter++
  } # ELSE user is valid and found.
} # End of FOR loop through users read from file
Write-Host

# Make stats printable
##############################
$errStringTotal = "Numero Totale Utenti Processati: " + ($users.count).ToString()
$errStringSuccess = "  Completati: " + $counterSuccess.ToString()
$errStringNoChange = "  Nessuna Modifica: " + $counterNoChange.ToString()
$errStringFailed = "  Falliti: " + $counterFailed.ToString()
$errStringNoAvailableLicense = "  Nessuna Licenza Disponibile: " + $counterNoAvailableLicense.ToString()

# Stats on screen
##############################
"==================================="
$errStringTotal
$errStringSuccess
$errStringNoChange
$errStringFailed
$errStringNoAvailableLicense

# Stats in log file
##############################
"===================================" | Out-File $logfile -Append
$errStringTotal | Out-File $logfile -Append
$errStringSuccess | Out-File $logfile -Append
$errStringNoChange | Out-File $logfile -Append
$errStringNoAvailableLicense | Out-File $logfile -Append
$errStringFailed | Out-File $logfile -Append

"Logging ended: " + (Get-Date).ToString() | Out-File $logfile -Append
"Run time: " + ("{0:N2}" -f $stopwatch.Elapsed.TotalMinutes).ToString() + " minutes" | Out-File $logfile -Append

Write-Host "Fine."
