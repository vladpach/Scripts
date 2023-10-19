$LogTime = Get-Date -Format yyyy-MM-dd
$LogFile = ".\UpdateManagedProperties-$LogTime.log"

#Managed Properties
$mpNames = ("WorkPhone", "OrganizationStructure", "BranchName", "CP-DataHash", "CP-WorkPhoneLongSuffix", "CP-WorkPhoneShortSuffix", "CP-PreferredName")

Start-Transcript $LogFile

#Add SharePoint PowerShell Snapin
if ($null -eq (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue))
{
    Add-PSSnapin Microsoft.SharePoint.Powershell
}

# try {
    $ssa = Get-SPEnterpriseSearchServiceApplication
    $proxy = Get-SPEnterpriseSearchServiceApplicationProxy

    $owner = [Microsoft.Office.Server.Search.Administration.SearchObjectOwner]::new([Microsoft.Office.Server.Search.Administration.SearchObjectLevel]::Ssa)
    $catInfo = $ssa.GetAllCategories($owner) | Where-Object { $_.Name -eq "People" }

    #Set some values of Managed Properties
    foreach ($mpName in $mpNames) {
        $mp = Get-SPEnterpriseSearchMetadataManagedProperty -SearchApplication $ssa -Identity $mpName.Replace("CP-", "").Replace("PreferredName", "PreferredNameCustom") -ea silentlycontinue

        if ($null -eq $mp) {
            <# The type of managed property must be one of the following data types:
            1 = Text
            2 = Integer
            3 = Decimal
            4 = DateTime
            5 = YesNo
            6 = Binary
            7 = Double #>
            if ($mpName -eq "CP-DataHash") {
                $mp = New-SPEnterpriseSearchMetadataManagedProperty -Name $mpName.Replace("CP-", "") -SearchApplication $ssa -Type 3 -Queryable $true -Retrievable $false
            }
            else {
                $isRetrievable = if (("OrganizationStructure", "BranchName", "CP-PreferredName") -contains $mpName ) { $true } else { $false }

                $mp = New-SPEnterpriseSearchMetadataManagedProperty -Name $mpName.Replace("CP-", "").Replace("PreferredName", "PreferredNameCustom") -SearchApplication $ssa -Type 1 -Queryable $true -Retrievable $isRetrievable
                $mp.Searchable = $true

                if ($mpName -eq "BranchName") {
                    $mp.Refinable = $true
                }
                if ($mpName -eq "CP-PreferredName") {
                    $mp.SafeForAnonymous = $true
                }
            }
            $mp.Update()

            #CrawledPropertyInfo Implicit
            $cpii = $proxy.QueryCrawledProperties($catInfo, $mpName, 1000, [GUID]"00000000-0000-0000-0000-000000000000", [string]::Empty, $true, $false, $owner)
            $cpi = $ssa.PromoteImplicitCrawledProperty($cpii[0], $owner)

            $cat = Get-SPEnterpriseSearchMetadataCategory -SearchApplication $ssa -Identity People
            $cp = Get-SPEnterpriseSearchMetadataCrawledProperty -Name $cpi.Name -SearchApplication $ssa -Category $cat
            New-SPEnterpriseSearchMetadataMapping -SearchApplication $ssa -ManagedProperty $mp -CrawledProperty $cp
            Write-Host "The managed property" $mpName "has been created successfully... Done!" -fore green
            <# else
        {
            Write-Host -f Yellow "The specified managed property " $mpName " does not exists... Please check whether you have given valid managed property name."
        }  #>
        }
        else {
            if ($mpName -eq "WorkPhone" -and $mp.Searchable -eq $false) {
                $mp.Searchable = $true
                $mp.Update()

                Write-Host "The Searchable for the managed property" $mpName "has been enabled successfully... Done!" -f green
            }

            <#
            $ManagedProperty.Queryable = $true
            $ManagedProperty.Sortable = $true
            $ManagedProperty.Retrievable = $true
            $ManagedProperty.SafeForAnonymous = $true
        #>

        }
    }
# }
# catch {
#     Write-Host "Exception: $PSItem" -f red
# }


#Create new Query Rules
$searchOwner = Get-SPEnterpriseSearchOwner -Level Ssa
$searchFilter = New-Object Microsoft.Office.Server.Search.Administration.SearchObjectFilter($searchOwner)
$ruleMgr = New-Object Microsoft.Office.Server.Search.Query.Rules.QueryRuleManager($ssa)
$federManager = New-Object Microsoft.Office.Server.Search.Administration.Query.FederationManager($ssa)
$commonResultSource = $federManager.GetSourceByName("Local SharePoint Results", $searchOwner)
$peopleResultSource = $federManager.GetSourceByName("Local People Results", $searchOwner)
$qrCollection = $ruleMgr.GetQueryRules($searchFilter)
$qrgCollection = $ruleMgr.GetQueryRuleGroups($searchFilter)
$qrGroup = $qrgCollection["CorpPortal Employees Group"]

#Common search
if ($null -eq $qrCollection["Append wildcard"])
{
    $rule = $qrCollection.CreateQueryRule("Append wildcard", $null, $null, $true)

    #Add Source
    $rule.CreateSourceContextCondition($commonResultSource) | Out-Null

    #Advanced Query Text Match
    $rule.QueryConditions.CreateRegularExpressionCondition("(.+)", $true) | Out-Null

    #Change ranked results by changing the query
    $rule.CreateQueryAction([Microsoft.Office.Server.Search.Query.Rules.QueryActionType]::ChangeQuery)
    $rule.ChangeQueryAction.QueryTransform.QueryTemplate = "{searchTerms}*"

    $rule.Update()

    Write-Host "`"Append wildcard`" query rule has been created successfully!" -f green
}

#People search
if ($null -eq $qrGroup) {
    #Create a Query Rule Group
    $qrGroup = $qrgCollection.CreateQueryRuleGroup("CorpPortal Employees Group")
}
else {
    $qrGroup = $qrGroup[0]
}

if ($null -eq $qrCollection["Search by WorkPhoneShortSuffix"])
{
    $rule = $qrCollection.CreateQueryRule("Search by WorkPhoneShortSuffix", $null, $null, $true)

    #Add Source
    $rule.CreateSourceContextCondition($peopleResultSource) | Out-Null

    #Advanced Query Text Match
    $rule.QueryConditions.CreateRegularExpressionCondition("^\d{4}$", $true) | Out-Null

    #Change ranked results by changing the query
    $rule.CreateQueryAction([Microsoft.Office.Server.Search.Query.Rules.QueryActionType]::ChangeQuery)
    $rule.ChangeQueryAction.QueryTransform.QueryTemplate = "WorkPhoneShortSuffix:{searchTerms}"

    $rule.Update()

    #Add rule to group
    $rule.MoveToGroup($qrGroup, [Microsoft.Office.Server.Search.Query.Rules.GroupProcessingDirective]::Stop)

    Write-Host "`"Search by WorkPhoneShortSuffix`" query rule has been created successfully!" -f green
}

if ($null -eq $qrCollection["Search by WorkPhoneLongSuffix"])
{
    $rule = $qrCollection.CreateQueryRule("Search by WorkPhoneLongSuffix", $null, $null, $true)

    #Add Source
    $rule.CreateSourceContextCondition($peopleResultSource) | Out-Null

    #Advanced Query Text Match
    $rule.QueryConditions.CreateRegularExpressionCondition("^\d{7}$", $true) | Out-Null

    #Change ranked results by changing the query
    $rule.CreateQueryAction([Microsoft.Office.Server.Search.Query.Rules.QueryActionType]::ChangeQuery)
    $rule.ChangeQueryAction.QueryTransform.QueryTemplate = "WorkPhoneLongSuffix:{searchTerms}"

    $rule.Update()

    #Add rule to group
    $rule.MoveToGroup($qrGroup, [Microsoft.Office.Server.Search.Query.Rules.GroupProcessingDirective]::Stop)

    Write-Host "`"Search by WorkPhoneLongSuffix`" query rule has been created successfully!" -f green
}

if ($null -eq $qrCollection["Filter by Ulcimus presence"])
{
    $rule = $qrCollection.CreateQueryRule("Filter by Ulcimus presence", $null, $null, $true)

    #Add Source
    $rule.CreateSourceContextCondition($peopleResultSource) | Out-Null

    #Advanced Query Text Match
    $rule.QueryConditions.CreateRegularExpressionCondition("(.+)", $true) | Out-Null

    #Change ranked results by changing the query
    $rule.CreateQueryAction([Microsoft.Office.Server.Search.Query.Rules.QueryActionType]::ChangeQuery)
    $rule.ChangeQueryAction.QueryTransform.QueryTemplate = "{searchTerms}* DataHash>0 DataHash<0"

    $rule.Update()

    #Add rule to group
    $rule.MoveToGroup($qrGroup, [Microsoft.Office.Server.Search.Query.Rules.GroupProcessingDirective]::Stop)

    Write-Host "`"Filter by Ulcimus presence`" query rule has been created successfully!" -f green
}

Stop-Transcript