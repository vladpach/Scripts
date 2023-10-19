using assembly "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.Office.Server.dll"
using assembly "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.Office.Server.UserProfiles.dll"

Param
(
    # Site collection URL
    [Parameter(Mandatory=$true)][string]
    $Url
)

enum Localization {
  English = 1033
  Russian = 1049
  Ukrainian = 1058
}

### about_Classes https://learn.microsoft.com/ru-ru/powershell/module/microsoft.powershell.core/about/about_classes?view=powershell-7.3
class UserProfileProperty {
  [string] $SiteUrl
  [string] $Name
  [string] $DisplayNameEN
  [string] $DisplayNameUA
  [string] $TypeName
  [int] $DisplayOrder
  [int] $Length
  [boolean] $IsVisible

  # Конструктор
  UserProfileProperty (
    [string] $siteUrl,
    [string] $name,
    [string] $displayNameEN,
    [string] $displayNameUA,
    [string] $typeName,
    [int] $displayOrder,
    [int] $length,
    [boolean] $isVisible
  ) {
    $this.SiteUrl = $siteUrl
    $this.Name = $name
    $this.DisplayNameEN = $displayNameEN
    $this.DisplayNameUA = $displayNameUA
    $this.TypeName = $typeName
    $this.DisplayOrder = $displayOrder
    $this.Length = $length
    $this.IsVisible = $isVisible
  }

  # Метод экземпляра
  [void] CreateProperty() {
    $site = Get-SPSite $this.SiteUrl
    $context = Get-SPServiceContext $site
    $upcm = New-Object Microsoft.Office.Server.UserProfiles.UserProfileConfigManager($context)
    $ppm = $upcm.ProfilePropertyManager

    #create core property
    $cpm = $ppm.GetCoreProperties()
    $pn = $cpm.GetPropertyByName($this.Name)
    if ($null -eq $pn) {
      $cp = $cpm.Create($false)
      $cp.Name = $this.Name
      $cp.DisplayName = $this.DisplayNameEN
      $cp.DisplayNameLocalized[[Localization]::English.value__] = $this.DisplayNameEN
      $cp.DisplayNameLocalized[[Localization]::Ukrainian.value__] = $this.DisplayNameUA
      $dataType = $this.TypeName
      $cp.Type = [Microsoft.Office.Server.UserProfiles.PropertyDataType]::$dataType
      if ($this.TypeName -gt 0) {
        $cp.Length = $this.Length
      }
      $cp.IsSearchable = $true

      $cpm.Add($cp)

      #create profile type property
      $ptpm = $ppm.GetProfileTypeProperties([Microsoft.Office.Server.UserProfiles.ProfileType]::User)
      $ptp = $ptpm.Create($cp)
      $ptp.IsVisibleOnEditor = $this.IsVisible
      #$ptp.IsVisibleOnViewer = $false

      $ptpm.Add($ptp)

      #create profile subtype property
      $psm = [Microsoft.Office.Server.UserProfiles.ProfileSubtypeManager]::Get($context)
      $ps = $psm.GetProfileSubtype([Microsoft.Office.Server.UserProfiles.ProfileSubtypeManager]::GetDefaultProfileName([Microsoft.Office.Server.UserProfiles.ProfileType]::User))
      $pspm = $ps.Properties
      $psp = $pspm.Create($ptp)

      $psp.PrivacyPolicy = [Microsoft.Office.Server.UserProfiles.PrivacyPolicy]::OptIn
      $psp.DefaultPrivacy = [Microsoft.Office.Server.UserProfiles.Privacy]::Public

      $pspm.Add($psp)

      if ($this.DisplayOrder -gt 0) {
        $pspm.SetDisplayOrderByPropertyName($this.Name, $this.DisplayOrder)
        $pspm.CommitDisplayOrder()
      }
    }
  }
}

$upp = [UserProfileProperty]::new($Url, "OrganizationStructure", "Organizational structure", "Організаційна структура", "StringSingleValue", 10, 3000, $true)
$upp.CreateProperty()

$upp = [UserProfileProperty]::new($Url, "BranchName", "Branch", "Філія", "StringSingleValue", 11, 250, $true)
$upp.CreateProperty()

$upp = [UserProfileProperty]::new($Url, "CP-DataHash", "DataHash", "", "BigInteger", 0, 0, $false)
$upp.CreateProperty()

$upp = [UserProfileProperty]::new($Url, "CP-WorkPhoneLongSuffix", "WorkPhoneLongSuffix", "", "StringSingleValue", 0, 7, $false)
$upp.CreateProperty()

$upp = [UserProfileProperty]::new($Url, "CP-WorkPhoneShortSuffix", "WorkPhoneShortSuffix", "", "StringSingleValue", 0, 4, $false)
$upp.CreateProperty()

$upp = [UserProfileProperty]::new($Url, "CP-PreferredName", "PreferredName", "", "StringSingleValue", 0, 256, $false)
$upp.CreateProperty()