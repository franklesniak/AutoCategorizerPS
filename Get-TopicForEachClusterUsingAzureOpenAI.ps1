# Get-TopicForEachClusterUsingAzureOpenAI.ps1
# Version: 1.1.20250412.0

<#
.SYNOPSIS
Inputs a CSV containing cluster information (e.g., k-means) and another CSV containing
unstructured text data attempts to determine the topic, category, or central theme of
each cluster.

.DESCRIPTION
This script first reads in a CSV file containing cluster metadata (e.g., for each
cluster, the most representative data points and the number of data points in the
cluster). It then reads in a CSV file containing unstructured text data. By relating
the two CSVs and then querying a large language model (e.g., OpenAI GPT-3) for the
topic, category, or central theme of each cluster, the script outputs a new CSV file
containing the cluster metadata and the topic, category, or central theme of each
cluster.

.PARAMETER ClusterMetadataInputCSVPath
Specifies the path to the input CSV file containing the cluster metadata.

.PARAMETER ClusterMetadataFieldNameIndexToMostRepresentativeItem
Specifies the name of the field in the cluster metadata input CSV file containing the
index of the most representative item in the cluster. The default value is
'MostRepresentativeItem'.

.PARAMETER ClusterMetadataFieldNameNumberOfNMostRepresentativeDataPoints
Specifies the name of the field in the cluster metadata input CSV file containing the
number of most representative data points. The default value is
'CountOfNMostRepresentativeItems'.

.PARAMETER ClusterMetadataFieldNameIndiciesOfNMostRepresentativeItems
Specifies the name of the field in the cluster metadata input CSV file containing the
indices of the most representative data points. The indices are specified as a
semicolon-separated list. The default value is 'NMostRepresentativeItems'.

.PARAMETER ClusterMetadataFieldNameCountOfItemsInCluster
Specifies the name of the field in the cluster metadata input CSV file containing the
number of items in the cluster. The default value is 'CountOfItemsInCluster'.

.PARAMETER ClusterMetadataFieldNameIndiciesOfAllItemsInCluster
Specifies the name of the field in the cluster metadata input CSV file containing the
indices of all items in the cluster. The indices are specified as a semicolon-separated
list. The default value is 'ItemsInCluster'.

.PARAMETER ClusterMetadataIndicesSeparator
Specifies the separator used when outputting multiple indices in the cluster (i.e.,
for when the 'n' most representative items or all items in the cluster are more than
one). The default value is '; '.

.PARAMETER UnstructuredTextDataInputCSVPath
Specifies the path to the input CSV file containing the unstructured text data. The CSV
file must contain a field that contains the unstructured text data.

.PARAMETER UnstructuredTextDataFieldNameContainingTextData
Specifies the name of the field in the unstructured text data input CSV file containing
the unstructured text data.

.PARAMETER OutputDataFromMostRepresentativeItem
Specifies whether the most representative item's data should be looked up and included
in the output CSV file. The default value is $true.

.PARAMETER OutputDataFieldNameForMostRepresentativeItem
If the most representative item's data is being included in the output CSV file, this
parameter specifies the name of the field in the output CSV file containing the most
representative item's data. The default value is 'MostRepresentativeItem'.

.PARAMETER UseMostRepresentativeItemForTopicExtraction
Specifies whether to use the most representative data point in the cluster for topic
extraction. The default value is $true.

.PARAMETER OutputDataFieldNameForTopicFromMostRepresentativeItem
If the most representative item is being used to generate the topic, this parameter
specifies the name of the field in the output CSV file containing the topic, category,
or central theme of the cluster derived from the most representative data point in the
cluster. The default value is 'TopicFromMostRepresentativeItem'.

.PARAMETER OutputDataFromNMostRepresentativeItems
Specifies whether the 'n' most representative items' data should be looked up and
included in the output CSV file. The default value is $true.

.PARAMETER OutputDataFieldNameForNMostRepresentativeItems
If the 'n' most representative items' data is being included in the output CSV file,
this parameter specifies the name of the field in the output CSV file containing the
'n' most representative items' data. The default value is 'NMostRepresentativeItems'.

.PARAMETER UseNMostRepresentativeItemsInClusterForTopicExtraction
Specifies whether to use the 'n' most representative data points in the cluster for
topic extraction. The default value is $true.

.PARAMETER OutputDataFieldNameForTopicFromNMostRepresentativeItems
If the 'n' most representative items are being used to generate the topic, this
parameter specifies the name of the field in the output CSV file containing the topic,
category, or central theme of the cluster derived from the 'n' most representative data
points in the cluster. The default value is 'TopicFromNMostRepresentativeItems'.

.PARAMETER OutputDataFromAllItemsInCluster
Specifies whether all items in the cluster's data should be looked up and included in
the output CSV file. The default value is $false.

.PARAMETER OutputDataFieldNameForAllItemsInCluster
If all items in the cluster are being included in the output CSV file, this parameter
specifies the name of the field in the output CSV file containing all items in the
cluster. The default value is 'ItemsInCluster'.

.PARAMETER UseAllItemsInClusterForTopicExtraction
Specifies whether to use all items in the cluster for topic extraction. The default
value is $false.

.PARAMETER OutputDataFieldNameForTopicFromAllItemsInCluster
If all items in the cluster are being used to generate the topic, this parameter
specifies the name of the field in the output CSV file containing the topic, category,
or central theme of the cluster derived from all items in the cluster. The default
value is 'TopicFromAllItemsInCluster'.

.PARAMETER SeparatorForItemsInCluster
Specifies the separator to use when outputting multiple items in the cluster (i.e.,
for use when OutputDataFromNMostRepresentativeItems or OutputDataFromAllItemsInCluster
are set to $true). The default value is ' /// '. The script will scan the unstructured
text data and make sure that the separator is not present in the data. If it is, the
script will throw a warning and quit.

.PARAMETER OutputCSVPath
Specifies the path to the output CSV file that will list the K-means cluster metadata.

.PARAMETER DoNotCheckForModuleUpdates
If supplied, the script will skip the check for PowerShell module updates. This can
speed up the script's execution time, but it is not recommended unless the user knows
that the computer's modules are already up-to-date.

.PARAMETER ReferenceToAzureOpenAIEndpoint
This parameter is required; it is a reference to a string containing the endpoint
for the Azure OpenAI service. To view the endpoint, for an Azure OpenAI resource,
go to the Azure portal and select the resource. Then, navigate to "Keys and
Endpoint" in the left-hand menu. The endpoint will be in the format
'https://<resource-name>.openai.azure.com/' where <resource-name> is the name of
the Azure OpenAI resource. Supply the complete endpoint URL, including the
https:// prefix, the .openai.azure.com suffix, and the trailing slash.

.PARAMETER ReferenceToAzureOpenAIDeploymentName
This parameter is required; it is a reference to a string that specifies the
deployment name in the Azure OpenAI service instance that represents the
chat response model to be used. The model deployments can be viewed in Azure AI
Foundry. To view the model deployments, go to
https://ai.azure.com/resource/deployments, then verify that the correct Azure
OpenAI instance is selected at the top. The model deployments are listed in the
middle pane. For this parameter, supply the name of the deployment that
represents the chat response model to be used. The deployment name is case-
sensitive.

.PARAMETER AzureOpenAIAPIVersion
Specifies the API version to use when connecting to the Azure OpenAI service. The
API version is supplied in YYYY-MM-DD format, and, if this parameter is omitted, the
script defaults to version 2024-12-01-preview. The latest GA API version can be viewed here:
https://learn.microsoft.com/en-us/azure/ai-services/openai/api-version-deprecation?source=recommendations#latest-ga-api-release

.PARAMETER EntraIdTenantId
Specifies the tenant ID to use when authenticating to Entra ID to access the Azure Key
Vault. The default tenant ID is the one used in Frank and Danny's demo.

.PARAMETER AzureSubscriptionId
Specifies the subscription ID that holds the Azure Key Vault. The default subscription
ID is the one used in Frank and Danny's demo.

.PARAMETER AzureKeyVaultName
Specifies the name of the Azure Key Vault that holds the API key (secret). The default
Key Vault name is the one used in Frank and Danny's demo.

.PARAMETER SecretName
Specifies the name of the secret in the Azure Key Vault. The secret must contain the
Azure OpenAI API key.

.EXAMPLE
PS C:\> .\Get-TopicForEachCluster.ps1 -ClusterMetadataInputCSVPath 'C:\Users\jdoe\Documents\West Monroe Pulse Survey Comments Aug 2021 - Cluster Metadata.csv' -UnstructuredTextDataInputCSVPath 'C:\Users\jdoe\Documents\West Monroe Pulse Survey Comments Aug 2021 - With Embeddings.csv' -UnstructuredTextDataFieldNameContainingTextData 'Comment' -OutputCSVPath 'C:\Users\jdoe\Documents\West Monroe Pulse Survey Comments Aug 2021 - Cluster Metadata with Topics.csv'

This example reads in a CSV file containing cluster metadata and another CSV file
containing unstructured text data. The script then queries a large language model
(e.g., OpenAI GPT-3) for the topic, category, or central theme of each cluster and
outputs a new CSV file containing the cluster metadata and the topic, category, or
central theme of each cluster.

.OUTPUTS
None
#>

#region License ################################################################
# Copyright (c) 2025 Frank Lesniak and Daniel Stutz
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of
# this software and associated documentation files (the "Software"), to deal in the
# Software without restriction, including without limitation the rights to use,
# copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
# Software, and to permit persons to whom the Software is furnished to do so,
# subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
# FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
# COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
# AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
# WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#endregion License ################################################################

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][string]$ClusterMetadataInputCSVPath,
    [Parameter(Mandatory = $false)][string]$ClusterMetadataFieldNameIndexToMostRepresentativeItem = 'MostRepresentativeItemIndex',
    [Parameter(Mandatory = $false)][string]$ClusterMetadataFieldNameNumberOfNMostRepresentativeDataPoints = 'CountOfNMostRepresentativeItems',
    [Parameter(Mandatory = $false)][string]$ClusterMetadataFieldNameIndiciesOfNMostRepresentativeItems = 'NMostRepresentativeItemIndices',
    [Parameter(Mandatory = $false)][string]$ClusterMetadataFieldNameCountOfItemsInCluster = 'CountOfItemsInCluster',
    [Parameter(Mandatory = $false)][string]$ClusterMetadataFieldNameIndiciesOfAllItemsInCluster = 'ItemsInClusterIndices',
    [Parameter(Mandatory = $false)][string]$ClusterMetadataIndicesSeparator = '; ',
    [Parameter(Mandatory = $true)][string]$UnstructuredTextDataInputCSVPath,
    [Parameter(Mandatory = $true)][string]$UnstructuredTextDataFieldNameContainingTextData,
    [Parameter(Mandatory = $false)][bool]$OutputDataFromMostRepresentativeItem = $true,
    [Parameter(Mandatory = $false)][string]$OutputDataFieldNameForMostRepresentativeItem = 'MostRepresentativeItem',
    [Parameter(Mandatory = $false)][bool]$UseMostRepresentativeItemForTopicExtraction = $true,
    [Parameter(Mandatory = $false)][string]$OutputDataFieldNameForTopicFromMostRepresentativeItem = 'TopicFromMostRepresentativeItem',
    [Parameter(Mandatory = $false)][bool]$OutputDataFromNMostRepresentativeItems = $true,
    [Parameter(Mandatory = $false)][string]$OutputDataFieldNameForNMostRepresentativeItems = 'NMostRepresentativeItems',
    [Parameter(Mandatory = $false)][bool]$UseNMostRepresentativeItemsInClusterForTopicExtraction = $true,
    [Parameter(Mandatory = $false)][string]$OutputDataFieldNameForTopicFromNMostRepresentativeItems = 'TopicFromNMostRepresentativeItems',
    [Parameter(Mandatory = $false)][bool]$OutputDataFromAllItemsInCluster = $false,
    [Parameter(Mandatory = $false)][string]$OutputDataFieldNameForAllItemsInCluster = 'ItemsInCluster',
    [Parameter(Mandatory = $false)][bool]$UseAllItemsInClusterForTopicExtraction = $false,
    [Parameter(Mandatory = $false)][string]$OutputDataFieldNameForTopicFromAllItemsInCluster = 'TopicFromAllItemsInCluster',
    [Parameter(Mandatory = $false)][string]$SeparatorForItemsInCluster = ' /// ',
    [Parameter(Mandatory = $true)][string]$OutputCSVPath,
    [Parameter(Mandatory = $false)][switch]$DoNotCheckForModuleUpdates,
    [Parameter(Mandatory = $true)][string]$AzureOpenAIEndpoint,
    [Parameter(Mandatory = $true)][string]$AzureOpenAIDeploymentName,
    [Parameter(Mandatory = $false)][string]$AzureOpenAIAPIVersion = '2024-12-01-preview',
    [Parameter(Mandatory = $false)][string]$EntraIdTenantId = '4cb2f1c9-c771-4ce5-a581-9376e59ea807',
    [Parameter(Mandatory = $false)][string]$AzureSubscriptionId = 'b337b2c0-fe35-4e3c-9434-7b7a15da61b7',
    [Parameter(Mandatory = $false)][string]$AzureKeyVaultName = 'powershell-conf-2024',
    [Parameter(Mandatory = $false)][string]$SecretName = 'powershell-saturday-openai-key'
)

function Get-PSVersion {
    # Returns the version of PowerShell that is running, including on the original
    # release of Windows PowerShell (version 1.0)
    #
    # Example:
    # Get-PSVersion
    #
    # This example returns the version of PowerShell that is running. On versions
    # of PowerShell greater than or equal to version 2.0, this function returns the
    # equivalent of $PSVersionTable.PSVersion
    #
    # The function outputs a [version] object representing the version of
    # PowerShell that is running
    #
    # PowerShell 1.0 does not have a $PSVersionTable variable, so this function
    # returns [version]('1.0') on PowerShell 1.0

    #region License ############################################################
    # Copyright (c) 2024 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy
    # of this software and associated documentation files (the "Software"), to deal
    # in the Software without restriction, including without limitation the rights
    # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    # copies of the Software, and to permit persons to whom the Software is
    # furnished to do so, subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in
    # all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    # SOFTWARE.
    #endregion License ############################################################

    #region DownloadLocationNotice #############################################
    # The most up-to-date version of this script can be found on the author's
    # GitHub repository at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice #############################################

    $versionThisFunction = [version]('1.0.20240326.0')

    if (Test-Path variable:\PSVersionTable) {
        return ($PSVersionTable.PSVersion)
    } else {
        return ([version]('1.0'))
    }
}

function Get-PowerShellModuleUsingHashtable {
    <#
    .SYNOPSIS
    Gets a list of installed PowerShell modules for each entry in a hashtable.

    .DESCRIPTION
    The Get-PowerShellModuleUsingHashtable function steps through each entry in the
    supplied hashtable and gets a list of installed PowerShell modules for each entry.

    .PARAMETER ReferenceToHashtable
    Is a reference to a hashtable. The value of the reference should be a hashtable
    with keys that are the names of PowerShell modules and values that are initialized
    to be enpty arrays.

    .EXAMPLE
    $hashtableModuleNameToInstalledModules = @{}
    $hashtableModuleNameToInstalledModules.Add('PnP.PowerShell', @())
    $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Authentication', @())
    $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Groups', @())
    $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Users', @())
    $refHashtableModuleNameToInstalledModules = [ref]$hashtableModuleNameToInstalledModules
    Get-PowerShellModuleUsingHashtable -ReferenceToHashtable $refHashtableModuleNameToInstalledModules

    This example gets the list of installed PowerShell modules for each of the four
    modules listed in the hashtable. The list of each respective module is stored in
    the value of the hashtable entry for that module.

    .OUTPUTS
    None
    #>

    #region License ################################################################
    # Copyright (c) 2024 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################################

    #region DownloadLocationNotice  ################################################
    # The most up-to-date version of this script can be found on the author's GitHub
    # repository at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice  ################################################

    # Version 1.0.20240401.0

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtable
    )

    $VerbosePreferenceAtStartOfFunction = $VerbosePreference

    $arrModulesToGet = @(($ReferenceToHashtable.Value).Keys)

    for ($intCounter = 0; $intCounter -lt $arrModulesToGet.Count; $intCounter++) {
        Write-Verbose ('Checking for ' + $arrModulesToGet[$intCounter] + ' module...')
        $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
        ($ReferenceToHashtable.Value).Item($arrModulesToGet[$intCounter]) = @(Get-Module -Name ($arrModulesToGet[$intCounter]) -ListAvailable)
        $VerbosePreference = $VerbosePreferenceAtStartOfFunction
    }
}

function Test-PowerShellModuleInstalledUsingHashtable {
    <#
    .SYNOPSIS
    Tests to see if a PowerShell module is installed based on entries in a hashtable.
    If the PowerShell module is not installed, an error or warning message may
    optionally be displayed.

    .DESCRIPTION
    The Test-PowerShellModuleInstalledUsingHashtable function steps through each entry
    in the supplied hashtable and, if there are any modules not installed, it
    optionally throws an error or warning for each module that is not installed. If all
    modules are installed, the function returns $true; otherwise, if any module is not
    installed, the function returns $false.

    .PARAMETER ReferenceToHashtableOfInstalledModules
    Is a reference to a hashtable. The hashtable must have keys that are the names of
    PowerShell modules with each key's value populated with arrays of
    ModuleInfoGrouping objects (the result of Get-Module).

    .PARAMETER ThrowErrorIfModuleNotInstalled
    Is a switch parameter. If this parameter is specified, an error is thrown for each
    module that is not installed. If this parameter is not specified, no error is
    thrown.

    .PARAMETER ThrowWarningIfModuleNotInstalled
    Is a switch parameter. If this parameter is specified, a warning is thrown for each
    module that is not installed. If this parameter is not specified, or if the
    ThrowErrorIfModuleNotInstalled parameter was specified, no warning is thrown.

    .PARAMETER ReferenceToHashtableOfCustomNotInstalledMessages
    Is a reference to a hashtable. The hashtable must have keys that are custom error
    or warning messages (string) to be displayed if one or more modules are not
    installed. The value for each key must be an array of PowerShell module names
    (strings) relevant to that error or warning message.

    If this parameter is not supplied, or if a custom error or warning message is not
    supplied in the hashtable for a given module, the script will default to using the
    following message:

    <MODULENAME> module not found. Please install it and then try again.
    You can install the <MODULENAME> PowerShell module from the PowerShell Gallery by
    running the following command:
    Install-Module <MODULENAME>;

    If the installation command fails, you may need to upgrade the version of
    PowerShellGet. To do so, run the following commands, then restart PowerShell:
    Set-ExecutionPolicy Bypass -Scope Process -Force;
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;
    Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;

    .PARAMETER ReferenceToArrayOfMissingModules
    Is a reference to an array. The array must be initialized to be empty. If any
    modules are not installed, the names of those modules are added to the array.

    .EXAMPLE
    $hashtableModuleNameToInstalledModules = @{}
    $hashtableModuleNameToInstalledModules.Add('PnP.PowerShell', @())
    $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Authentication', @())
    $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Groups', @())
    $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Users', @())
    $refHashtableModuleNameToInstalledModules = [ref]$hashtableModuleNameToInstalledModules
    Get-PowerShellModuleUsingHashtable -ReferenceToHashtable $refHashtableModuleNameToInstalledModules
    $hashtableCustomNotInstalledMessageToModuleNames = @{}
    $strGraphNotInstalledMessage = 'Microsoft.Graph.Authentication, Microsoft.Graph.Groups, and/or Microsoft.Graph.Users modules were not found. Please install the full Microsoft.Graph module and then try again.' + [System.Environment]::NewLine + 'You can install the Microsoft.Graph PowerShell module from the PowerShell Gallery by running the following command:' + [System.Environment]::NewLine + 'Install-Module Microsoft.Graph;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    $hashtableCustomNotInstalledMessageToModuleNames.Add($strGraphNotInstalledMessage, @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users'))
    $refhashtableCustomNotInstalledMessageToModuleNames = [ref]$hashtableCustomNotInstalledMessageToModuleNames
    $boolResult = Test-PowerShellModuleInstalledUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ThrowErrorIfModuleNotInstalled -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames

    This example checks to see if the PnP.PowerShell, Microsoft.Graph.Authentication,
    Microsoft.Graph.Groups, and Microsoft.Graph.Users modules are installed. If any of
    these modules are not installed, an error is thrown for the PnP.PowerShell module
    or the group of Microsoft.Graph modules, respectively, and $boolResult is set to
    $false. If all modules are installed, $boolResult is set to $true.

    .OUTPUTS
    [boolean] - Returns $true if all modules are installed; otherwise, returns $false.
    #>

    #region License ################################################################
    # Copyright (c) 2024 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################################

    #region DownloadLocationNotice  ################################################
    # The most up-to-date version of this script can be found on the author's GitHub
    # repository at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice  ################################################

    # Version 1.1.20240401.0

    [CmdletBinding()]
    [OutputType([Boolean])]
    param (
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfInstalledModules,
        [switch]$ThrowErrorIfModuleNotInstalled,
        [switch]$ThrowWarningIfModuleNotInstalled,
        [Parameter(Mandatory = $false)][ref]$ReferenceToHashtableOfCustomNotInstalledMessages,
        [Parameter(Mandatory = $false)][ref]$ReferenceToArrayOfMissingModules
    )

    $boolThrowErrorForMissingModule = $false
    $boolThrowWarningForMissingModule = $false

    if ($ThrowErrorIfModuleNotInstalled.IsPresent -eq $true) {
        $boolThrowErrorForMissingModule = $true
    } elseif ($ThrowWarningIfModuleNotInstalled.IsPresent -eq $true) {
        $boolThrowWarningForMissingModule = $true
    }

    $boolResult = $true

    $hashtableMessagesToThrowForMissingModule = @{}
    $hashtableModuleNameToCustomMessageToThrowForMissingModule = @{}
    if ($null -ne $ReferenceToHashtableOfCustomNotInstalledMessages) {
        $arrMessages = @(($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Keys)
        foreach ($strMessage in $arrMessages) {
            $hashtableMessagesToThrowForMissingModule.Add($strMessage, $false)

            ($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Item($strMessage) | ForEach-Object {
                $hashtableModuleNameToCustomMessageToThrowForMissingModule.Add($_, $strMessage)
            }
        }
    }

    $arrModuleNames = @(($ReferenceToHashtableOfInstalledModules.Value).Keys)
    foreach ($strModuleName in $arrModuleNames) {
        $arrInstalledModules = @(($ReferenceToHashtableOfInstalledModules.Value).Item($strModuleName))
        if ($arrInstalledModules.Count -eq 0) {
            $boolResult = $false

            if ($hashtableModuleNameToCustomMessageToThrowForMissingModule.ContainsKey($strModuleName) -eq $true) {
                $strMessage = $hashtableModuleNameToCustomMessageToThrowForMissingModule.Item($strModuleName)
                $hashtableMessagesToThrowForMissingModule.Item($strMessage) = $true
            } else {
                $strMessage = $strModuleName + ' module not found. Please install it and then try again.' + [System.Environment]::NewLine + 'You can install the ' + $strModuleName + ' PowerShell module from the PowerShell Gallery by running the following command:' + [System.Environment]::NewLine + 'Install-Module ' + $strModuleName + ';' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
                $hashtableMessagesToThrowForMissingModule.Add($strMessage, $true)
            }

            if ($null -ne $ReferenceToArrayOfMissingModules) {
                ($ReferenceToArrayOfMissingModules.Value) += $strModuleName
            }
        }
    }

    if ($boolThrowErrorForMissingModule -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForMissingModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingModule.Item($strMessage) -eq $true) {
                Write-Error $strMessage
            }
        }
    } elseif ($boolThrowWarningForMissingModule -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForMissingModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingModule.Item($strMessage) -eq $true) {
                Write-Warning $strMessage
            }
        }
    }

    return $boolResult
}

function Test-PowerShellModuleUpdatesAvailableUsingHashtable {
    <#
    .SYNOPSIS
    Tests to see if updates are available for a PowerShell module based on entries in a
    hashtable. If updates are available for a PowerShell module, an error or warning
    message may optionally be displayed.

    .DESCRIPTION
    The Test-PowerShellModuleUpdatesAvailableUsingHashtable function steps through each
    entry in the supplied hashtable and, if there are updates available, it optionally
    throws an error or warning for each module that has updates available. If all
    modules are installed and up to date, the function returns $true; otherwise, if any
    module is not installed or not up to date, the function returns $false.

    .PARAMETER ReferenceToHashtableOfInstalledModules
    Is a reference to a hashtable. The hashtable must have keys that are the names of
    PowerShell modules with each key's value populated with arrays of
    ModuleInfoGrouping objects (the result of Get-Module).

    .PARAMETER ThrowErrorIfModuleNotInstalled
    Is a switch parameter. If this parameter is specified, an error is thrown for each
    module that is not installed. If this parameter is not specified, no error is
    thrown.

    .PARAMETER ThrowWarningIfModuleNotInstalled
    Is a switch parameter. If this parameter is specified, a warning is thrown for each
    module that is not installed. If this parameter is not specified, or if the
    ThrowErrorIfModuleNotInstalled parameter was specified, no warning is thrown.

    .PARAMETER ThrowErrorIfModuleNotUpToDate
    Is a switch parameter. If this parameter is specified, an error is thrown for each
    module that is not up to date. If this parameter is not specified, no error is
    thrown.

    .PARAMETER ThrowWarningIfModuleNotUpToDate
    Is a switch parameter. If this parameter is specified, a warning is thrown for each
    module that is not up to date. If this parameter is not specified, or if the
    ThrowErrorIfModuleNotUpToDate parameter was specified, no warning is thrown.

    .PARAMETER ReferenceToHashtableOfCustomNotInstalledMessages
    Is a reference to a hashtable. The hashtable must have keys that are custom error
    or warning messages (string) to be displayed if one or more modules are not
    installed. The value for each key must be an array of PowerShell module names
    (strings) relevant to that error or warning message.

    If this parameter is not supplied, or if a custom error or warning message is not
    supplied in the hashtable for a given module, the script will default to using the
    following message:

    <MODULENAME> module not found. Please install it and then try again.
    You can install the <MODULENAME> PowerShell module from the PowerShell Gallery by
    running the following command:
    Install-Module <MODULENAME>;

    If the installation command fails, you may need to upgrade the version of
    PowerShellGet. To do so, run the following commands, then restart PowerShell:
    Set-ExecutionPolicy Bypass -Scope Process -Force;
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;
    Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;

    .PARAMETER ReferenceToHashtableOfCustomNotUpToDateMessages
    Is a reference to a hashtable. The hashtable must have keys that are custom error
    or warning messages (string) to be displayed if one or more modules are not
    up to date. The value for each key must be an array of PowerShell module names
    (strings) relevant to that error or warning message.

    If this parameter is not supplied, or if a custom error or warning message is not
    supplied in the hashtable for a given module, the script will default to using the
    following message:

    A newer version of the <MODULENAME> PowerShell module is available. Please consider
    updating it by running the following command:
    Install-Module <MODULENAME> -Force;

    If the installation command fails, you may need to upgrade the version of
    PowerShellGet. To do so, run the following commands, then restart PowerShell:
    Set-ExecutionPolicy Bypass -Scope Process -Force;
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;
    Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;

    .PARAMETER ReferenceToArrayOfMissingModules
    Is a reference to an array. The array must be initialized to be empty. If any
    modules are not installed, the names of those modules are added to the array.

    .PARAMETER ReferenceToArrayOfOutOfDateModules
    Is a reference to an array. The array must be initialized to be empty. If any
    modules are not up to date, the names of those modules are added to the array.

    .EXAMPLE
    $hashtableModuleNameToInstalledModules = @{}
    $hashtableModuleNameToInstalledModules.Add('PnP.PowerShell', @())
    $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Authentication', @())
    $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Groups', @())
    $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Users', @())
    $refHashtableModuleNameToInstalledModules = [ref]$hashtableModuleNameToInstalledModules
    Get-PowerShellModuleUsingHashtable -ReferenceToHashtable $refHashtableModuleNameToInstalledModules
    $hashtableCustomNotInstalledMessageToModuleNames = @{}
    $strGraphNotInstalledMessage = 'Microsoft.Graph.Authentication, Microsoft.Graph.Groups, and/or Microsoft.Graph.Users modules were not found. Please install the full Microsoft.Graph module and then try again.' + [System.Environment]::NewLine + 'You can install the Microsoft.Graph PowerShell module from the PowerShell Gallery by running the following command:' + [System.Environment]::NewLine + 'Install-Module Microsoft.Graph;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    $hashtableCustomNotInstalledMessageToModuleNames.Add($strGraphNotInstalledMessage, @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users'))
    $refhashtableCustomNotInstalledMessageToModuleNames = [ref]$hashtableCustomNotInstalledMessageToModuleNames
    $hashtableCustomNotUpToDateMessageToModuleNames = @{}
    $strGraphNotUpToDateMessage = 'A newer version of the Microsoft.Graph.Authentication, Microsoft.Graph.Groups, and/or Microsoft.Graph.Users modules was found. Please consider updating it by running the following command:' + [System.Environment]::NewLine + 'Install-Module Microsoft.Graph -Force;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    $hashtableCustomNotUpToDateMessageToModuleNames.Add($strGraphNotUpToDateMessage, @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users'))
    $refhashtableCustomNotUpToDateMessageToModuleNames = [ref]$hashtableCustomNotUpToDateMessageToModuleNames
    $boolResult = Test-PowerShellModuleUpdatesAvailableUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ThrowErrorIfModuleNotInstalled -ThrowWarningIfModuleNotUpToDate -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames -ReferenceToHashtableOfCustomNotUpToDateMessages $refhashtableCustomNotUpToDateMessageToModuleNames

    This example checks to see if the PnP.PowerShell, Microsoft.Graph.Authentication,
    Microsoft.Graph.Groups, and Microsoft.Graph.Users modules are installed. If any of
    these modules are not installed, an error is thrown for the PnP.PowerShell module
    or the group of Microsoft.Graph modules, respectively, and $boolResult is set to
    $false. If any of these modules are installed but not up to date, a warning
    message is thrown for the PnP.PowerShell module or the group of Microsoft.Graph
    modules, respectively, and $boolResult is set to false. If all modules are
    installed and up to date, $boolResult is set to $true.

    .OUTPUTS
    [boolean] - Returns $true if all modules are installed and up to date; otherwise,
    returns $false.

    .NOTES
    Requires PowerShell v5.0 or newer
    #>

    #region License ################################################################
    # Copyright (c) 2024 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################################

    #region DownloadLocationNotice  ################################################
    # The most up-to-date version of this script can be found on the author's GitHub
    # repository at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice  ################################################

    # Version 1.1.20240401.0

    [CmdletBinding()]
    [OutputType([Boolean])]
    param (
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfInstalledModules,
        [switch]$ThrowErrorIfModuleNotInstalled,
        [switch]$ThrowWarningIfModuleNotInstalled,
        [switch]$ThrowErrorIfModuleNotUpToDate,
        [switch]$ThrowWarningIfModuleNotUpToDate,
        [Parameter(Mandatory = $false)][ref]$ReferenceToHashtableOfCustomNotInstalledMessages,
        [Parameter(Mandatory = $false)][ref]$ReferenceToHashtableOfCustomNotUpToDateMessages,
        [Parameter(Mandatory = $false)][ref]$ReferenceToArrayOfMissingModules,
        [Parameter(Mandatory = $false)][ref]$ReferenceToArrayOfOutdatedModules
    )

    function Get-PSVersion {
        # Returns the version of PowerShell that is running, including on the original
        # release of Windows PowerShell (version 1.0)
        #
        # Example:
        # Get-PSVersion
        #
        # This example returns the version of PowerShell that is running. On versions
        # of PowerShell greater than or equal to version 2.0, this function returns the
        # equivalent of $PSVersionTable.PSVersion
        #
        # The function outputs a [version] object representing the version of
        # PowerShell that is running
        #
        # PowerShell 1.0 does not have a $PSVersionTable variable, so this function
        # returns [version]('1.0') on PowerShell 1.0

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/PowerShell_Resources
        #endregion DownloadLocationNotice #############################################

        $versionThisFunction = [version]('1.0.20240326.0')

        if (Test-Path variable:\PSVersionTable) {
            return ($PSVersionTable.PSVersion)
        } else {
            return ([version]('1.0'))
        }
    }

    $versionPS = Get-PSVersion
    if ($versionPS -lt ([version]'5.0')) {
        Write-Warning 'Test-PowerShellModuleUpdatesAvailableUsingHashtable requires PowerShell version 5.0 or newer.'
        return $false
    }

    $boolThrowErrorForMissingModule = $false
    $boolThrowWarningForMissingModule = $false

    if ($ThrowErrorIfModuleNotInstalled.IsPresent -eq $true) {
        $boolThrowErrorForMissingModule = $true
    } elseif ($ThrowWarningIfModuleNotInstalled.IsPresent -eq $true) {
        $boolThrowWarningForMissingModule = $true
    }

    $boolThrowErrorForOutdatedModule = $false
    $boolThrowWarningForOutdatedModule = $false

    if ($ThrowErrorIfModuleNotUpToDate.IsPresent -eq $true) {
        $boolThrowErrorForOutdatedModule = $true
    } elseif ($ThrowWarningIfModuleNotUpToDate.IsPresent -eq $true) {
        $boolThrowWarningForOutdatedModule = $true
    }

    $VerbosePreferenceAtStartOfFunction = $VerbosePreference

    $boolResult = $true

    $hashtableMessagesToThrowForMissingModule = @{}
    $hashtableModuleNameToCustomMessageToThrowForMissingModule = @{}
    if ($null -ne $ReferenceToHashtableOfCustomNotInstalledMessages) {
        $arrMessages = @(($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Keys)
        foreach ($strMessage in $arrMessages) {
            $hashtableMessagesToThrowForMissingModule.Add($strMessage, $false)

            ($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Item($strMessage) | ForEach-Object {
                $hashtableModuleNameToCustomMessageToThrowForMissingModule.Add($_, $strMessage)
            }
        }
    }

    $hashtableMessagesToThrowForOutdatedModule = @{}
    $hashtableModuleNameToCustomMessageToThrowForOutdatedModule = @{}
    if ($null -ne $ReferenceToHashtableOfCustomNotUpToDateMessages) {
        $arrMessages = @(($ReferenceToHashtableOfCustomNotUpToDateMessages.Value).Keys)
        foreach ($strMessage in $arrMessages) {
            $hashtableMessagesToThrowForOutdatedModule.Add($strMessage, $false)

            ($ReferenceToHashtableOfCustomNotUpToDateMessages.Value).Item($strMessage) | ForEach-Object {
                $hashtableModuleNameToCustomMessageToThrowForOutdatedModule.Add($_, $strMessage)
            }
        }
    }

    $arrModuleNames = @(($ReferenceToHashtableOfInstalledModules.Value).Keys)
    foreach ($strModuleName in $arrModuleNames) {
        $arrInstalledModules = @(($ReferenceToHashtableOfInstalledModules.Value).Item($strModuleName))
        if ($arrInstalledModules.Count -eq 0) {
            $boolResult = $false

            if ($hashtableModuleNameToCustomMessageToThrowForMissingModule.ContainsKey($strModuleName) -eq $true) {
                $strMessage = $hashtableModuleNameToCustomMessageToThrowForMissingModule.Item($strModuleName)
                $hashtableMessagesToThrowForMissingModule.Item($strMessage) = $true
            } else {
                $strMessage = $strModuleName + ' module not found. Please install it and then try again.' + [System.Environment]::NewLine + 'You can install the ' + $strModuleName + ' PowerShell module from the PowerShell Gallery by running the following command:' + [System.Environment]::NewLine + 'Install-Module ' + $strModuleName + ';' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
                $hashtableMessagesToThrowForMissingModule.Add($strMessage, $true)
            }

            if ($null -ne $ReferenceToArrayOfMissingModules) {
                ($ReferenceToArrayOfMissingModules.Value) += $strModuleName
            }
        } else {
            $versionNewestInstalledModule = ($arrInstalledModules | ForEach-Object { [version]($_.Version) } | Sort-Object)[-1]

            $arrModuleNewestInstalledModule = @($arrInstalledModules | Where-Object { ([version]($_.Version)) -eq $versionNewestInstalledModule })

            # In the event there are multiple installations of the same version, reduce to a
            # single instance of the module
            if ($arrModuleNewestInstalledModule.Count -gt 1) {
                $moduleNewestInstalled = @($arrModuleNewestInstalledModule | Select-Object -Unique)[0]
            } else {
                $moduleNewestInstalled = $arrModuleNewestInstalledModule[0]
            }

            $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
            $moduleNewestAvailable = Find-Module -Name $strModuleName -ErrorAction SilentlyContinue
            $VerbosePreference = $VerbosePreferenceAtStartOfFunction

            if ($null -ne $moduleNewestAvailable) {
                if ($moduleNewestAvailable.Version -gt $moduleNewestInstalled.Version) {
                    # A newer version is available
                    $boolResult = $false

                    if ($hashtableModuleNameToCustomMessageToThrowForOutdatedModule.ContainsKey($strModuleName) -eq $true) {
                        $strMessage = $hashtableModuleNameToCustomMessageToThrowForOutdatedModule.Item($strModuleName)
                        $hashtableMessagesToThrowForOutdatedModule.Item($strMessage) = $true
                    } else {
                        $strMessage = 'A newer version of the ' + $strModuleName + ' PowerShell module is available. Please consider updating it by running the following command:' + [System.Environment]::NewLine + 'Install-Module ' + $strModuleName + ' -Force;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
                        $hashtableMessagesToThrowForOutdatedModule.Add($strMessage, $true)
                    }

                    if ($null -ne $ReferenceToArrayOfOutdatedModules) {
                        ($ReferenceToArrayOfOutdatedModules.Value) += $strModuleName
                    }
                }
            } else {
                # Couldn't find the module in the PowerShell Gallery
            }
        }
    }

    if ($boolThrowErrorForMissingModule -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForMissingModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingModule.Item($strMessage) -eq $true) {
                Write-Error $strMessage
            }
        }
    } elseif ($boolThrowWarningForMissingModule -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForMissingModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingModule.Item($strMessage) -eq $true) {
                Write-Warning $strMessage
            }
        }
    }

    if ($boolThrowErrorForOutdatedModule -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForOutdatedModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForOutdatedModule.Item($strMessage) -eq $true) {
                Write-Error $strMessage
            }
        }
    } elseif ($boolThrowWarningForOutdatedModule -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForOutdatedModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForOutdatedModule.Item($strMessage) -eq $true) {
                Write-Warning $strMessage
            }
        }
    }
    return $boolResult
}

function Split-StringOnLiteralString {
    # Split-StringOnLiteralString is designed to split a string the way the way that I
    # expected it to be done - using a literal string (as opposed to regex). It's also
    # designed to be backward-compatible with all versions of PowerShell and has been
    # tested successfully on PowerShell v1. My motivation for creating this function
    # was 1) I wanted a split function that behaved more like VBScript's Split
    # function, 2) I did not want to be messing around with RegEx, and 3) I needed code
    # that was backward-compatible with all versions of PowerShell.
    #
    # This function takes two positional arguments
    # The first argument is a string, and the string to be split
    # The second argument is a string or char, and it is that which is to split the string in the first parameter
    #
    # Note: This function always returns an array, even when there is zero or one element in it.
    #
    # Example:
    # $result = Split-StringOnLiteralString 'foo' ' '
    # # $result.GetType().FullName is System.Object[]
    # # $result.Count is 1
    #
    # Example 2:
    # $result = Split-StringOnLiteralString 'What do you think of this function?' ' '
    # # $result.Count is 7

    #region License ################################################################
    # Copyright 2023 Frank Lesniak

    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:

    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.

    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################################

    #region DownloadLocationNotice #################################################
    # The most up-to-date version of this script can be found on the author's GitHub
    # repository at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice #################################################

    $strThisFunctionVersionNumber = [version]'2.0.20230708.0'

    trap {
        Write-Error 'An error occurred using the Split-StringOnLiteralString function. This was most likely caused by the arguments supplied not being strings'
    }

    if ($args.Length -ne 2) {
        Write-Error 'Split-StringOnLiteralString was called without supplying two arguments. The first argument should be the string to be split, and the second should be the string or character on which to split the string.'
        $result = @()
    } else {
        $objToSplit = $args[0]
        $objSplitter = $args[1]
        if ($null -eq $objToSplit) {
            $result = @()
        } elseif ($null -eq $objSplitter) {
            # Splitter was $null; return string to be split within an array (of one element).
            $result = @($objToSplit)
        } else {
            if ($objToSplit.GetType().Name -ne 'String') {
                Write-Warning 'The first argument supplied to Split-StringOnLiteralString was not a string. It will be attempted to be converted to a string. To avoid this warning, cast arguments to a string before calling Split-StringOnLiteralString.'
                $strToSplit = [string]$objToSplit
            } else {
                $strToSplit = $objToSplit
            }

            if (($objSplitter.GetType().Name -ne 'String') -and ($objSplitter.GetType().Name -ne 'Char')) {
                Write-Warning 'The second argument supplied to Split-StringOnLiteralString was not a string. It will be attempted to be converted to a string. To avoid this warning, cast arguments to a string before calling Split-StringOnLiteralString.'
                $strSplitter = [string]$objSplitter
            } elseif ($objSplitter.GetType().Name -eq 'Char') {
                $strSplitter = [string]$objSplitter
            } else {
                $strSplitter = $objSplitter
            }

            $strSplitterInRegEx = [regex]::Escape($strSplitter)

            # With the leading comma, force encapsulation into an array so that an array is
            # returned even when there is one element:
            $result = @([regex]::Split($strToSplit, $strSplitterInRegEx))
        }
    }

    # The following code forces the function to return an array, always, even when there are
    # zero or one elements in the array
    $intElementCount = 1
    if ($null -ne $result) {
        if ($result.GetType().FullName.Contains('[]')) {
            if (($result.Count -ge 2) -or ($result.Count -eq 0)) {
                $intElementCount = $result.Count
            }
        }
    }
    $strLowercaseFunctionName = $MyInvocation.InvocationName.ToLower()
    $boolArrayEncapsulation = $MyInvocation.Line.ToLower().Contains('@(' + $strLowercaseFunctionName + ')') -or $MyInvocation.Line.ToLower().Contains('@(' + $strLowercaseFunctionName + ' ')
    if ($boolArrayEncapsulation) {
        $result
    } elseif ($intElementCount -eq 0) {
        , @()
    } elseif ($intElementCount -eq 1) {
        , (, ($args[0]))
    } else {
        $result
    }
}

function Get-AzureOpenAISingleChatResponseRobust {
    # .SYNOPSIS
    # This function submits text (a prompt) to Azure OpenAI and returns the chat
    # response. Designed for o1 and newer models.
    #
    # .DESCRIPTION
    # This function submits text (a prompt) to Azure OpenAI and returns the chat
    # response. It is designed to be robust, with retry logic for transient errors. The
    # function will retry the request up to a specified maximum number of attempts if
    # it encounters certain error codes. The function also includes parameters to
    # customize the Azure OpenAI endpoint, deployment name, API key, and other
    # settings. The function is designed to be used in Windows PowerShell 3 and later
    # versions.
    #
    # .PARAMETER ReferenceToResponse
    # This parameter is required; it is a reference to a string that will be used to
    # store the chat response retrieved from the Azure OpenAI service.
    #
    # .PARAMETER CurrentAttemptNumber
    # This parameter is required; it is an integer indicating the current attempt
    # number. When calling this function for the first time, it should be 1.
    #
    # .PARAMETER MaxAttempts
    # This parameter is required; it is an integer representing the maximum number
    # of attempts that the function will observe before giving up.
    #
    # .PARAMETER ReferenceToAzureOpenAIEndpoint
    # This parameter is required; it is a reference to a string containing the endpoint
    # for the Azure OpenAI service. To view the endpoint, for an Azure OpenAI resource,
    # go to the Azure portal and select the resource. Then, navigate to "Keys and
    # Endpoint" in the left-hand menu. The endpoint will be in the format
    # 'https://<resource-name>.openai.azure.com/' where <resource-name> is the name of
    # the Azure OpenAI resource. Supply the complete endpoint URL, including the
    # https:// prefix, the .openai.azure.com suffix, and the trailing slash.
    #
    # .PARAMETER ReferenceToAzureOpenAIDeploymentName
    # This parameter is required; it is a reference to a string that specifies the
    # deployment name in the Azure OpenAI service instance that represents the
    # chat response model to be used. The model deployments can be viewed in Azure AI
    # Foundry. To view the model deployments, go to
    # https://ai.azure.com/resource/deployments, then verify that the correct Azure
    # OpenAI instance is selected at the top. The model deployments are listed in the
    # middle pane. For this parameter, supply the name of the deployment that
    # represents the chat response model to be used. The deployment name is case-
    # sensitive.
    #
    # .PARAMETER ReferenceToAPIKey
    # This parameter is required; is a reference to a string containing a valid Azure
    # OpenAI API key that the function will use to submit the text (prompt) and receive
    # the chat response.
    #
    # .PARAMETER ReferenceToPrompt
    # This parameter is required; it is a reference to a string containing the text
    # that the function will prompt to (ask) the chat API.
    #
    # .PARAMETER ReferenceToAzureOpenAIAPIVersion
    # This parameter is optional; if supplied, it is a memory reference (pointer) to a
    # string that specifies the API version to use when connecting to the Azure OpenAI
    # service. The API version is supplied in YYYY-MM-DD format, and, if this parameter
    # is omitted, the script defaults to version 2024-12-01-preview. The latest GA API
    # version can be viewed here:
    # https://learn.microsoft.com/en-us/azure/ai-services/openai/api-version-deprecation?source=recommendations#latest-ga-api-release
    #
    # .PARAMETER PSVersion
    # This parameter is optional; if supplied, it is a System.Version object that
    # contains the version of PowerShell that is running. If this parameter is
    # omitted, the function will determine the version of PowerShell that is
    # currently running. This parameter is used to determine whether the script is
    # running on a supported version of PowerShell.
    #
    # .EXAMPLE
    # $strTopic = ''
    # $strAzureOpenAIEndpoint = 'https://flesniak-dstutz.openai.azure.com/'
    # $strAzureOpenAIDeploymentName = 'private-chat'
    # $strAzureOpenAIAPIVersion = '2024-12-01-preview'
    # $strAPIKey = 'abcdefghijklmnopqrstuvwxyzabcdef'
    # $strPrompt = 'In as few words as possible (certainly no more than 1-3 words), '
    # $strPrompt += 'describe the topic, main idea, or theme of the following four '
    # $strPrompt += 'texts, treated as a set. Each text is separated by three forward '
    # $strPrompt += 'slashes (///): '
    # $arrTexts = @()
    # $arrTexts += 'I would really like to be able to work out during the work day but there are no showers at my office.'
    # $arrTexts += 'Having exercise equipment in the building is great!'
    # $arrTexts += 'Someone stole my iPad from the exercise room and the building manager is not doing anything about it.'
    # $arrTexts += 'Can we work out during our lunch break? I know some people do it, but I''m not sure it''s permitted. Can leadership make an announcement about this?'
    # $strPrompt += $arrTexts -join '///'
    # $boolSuccess = Get-AzureOpenAISingleChatResponseRobust -ReferenceToResponse ([ref]$strTopic) -CurrentAttemptNumber 1 -MaxAttempts 8 -ReferenceToAzureOpenAIEndpoint ([ref]$strAzureOpenAIEndpoint) -ReferenceToAzureOpenAIDeploymentName ([ref]$strAzureOpenAIDeploymentName) -ReferenceToAPIKey ([ref]$strAPIKey) -ReferenceToAzureOpenAIAPIVersion ([ref]$strAzureOpenAIAPIVersion) -ReferenceToPrompt ([ref]$strPrompt)
    #
    # .INPUTS
    # None. You can't pipe objects to Get-AzureOpenAISingleChatResponseRobust.
    #
    # .OUTPUTS
    # System.Boolean. Get-AzureOpenAISingleChatResponseRobust returns a boolean value
    # indiciating whether the process completed successfully. $true means the
    # process completed successfully; $false means there was an error.
    #
    # .NOTES
    # Version: 2.0.20250410.0

    #region License ############################################################
    # Copyright (c) 2025 Frank Lesniak and Daniel Stutz
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy
    # of this software and associated documentation files (the "Software"), to deal
    # in the Software without restriction, including without limitation the rights
    # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    # copies of the Software, and to permit persons to whom the Software is
    # furnished to do so, subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in
    # all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    # SOFTWARE.
    #endregion License ############################################################

    param (
        [ref]$ReferenceToResponse = ([ref]$null),
        [int]$CurrentAttemptNumber = 1,
        [int]$MaxAttempts = 4,
        [ref]$ReferenceToAzureOpenAIEndpoint = ([ref]$null),
        [ref]$ReferenceToAzureOpenAIDeploymentName = ([ref]$null),
        [ref]$ReferenceToAPIKey = ([ref]$null),
        [ref]$ReferenceToPrompt = ([ref]$null),
        [ref]$ReferenceToAzureOpenAIAPIVersion = ([ref]'2024-12-01-preview'),
        [version]$PSVersion = ([version]'0.0')
    )

    #region FunctionsToSupportErrorHandling ####################################
    function Get-ReferenceToLastError {
        # .SYNOPSIS
        # Gets a reference (memory pointer) to the last error that
        # occurred.
        #
        # .DESCRIPTION
        # Returns a reference (memory pointer) to $null ([ref]$null) if no
        # errors on on the $error stack; otherwise, returns a reference to
        # the last error that occurred.
        #
        # .EXAMPLE
        # # Intentionally empty trap statement to prevent terminating
        # # errors from halting processing
        # trap { }
        #
        # # Retrieve the newest error on the stack prior to doing work:
        # $refLastKnownError = Get-ReferenceToLastError
        #
        # # Store current error preference; we will restore it after we do
        # # some work:
        # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
        #
        # # Set ErrorActionPreference to SilentlyContinue; this will suppress
        # # error output. Terminating errors will not output anything, kick
        # # to the empty trap statement and then continue on. Likewise, non-
        # # terminating errors will also not output anything, but they do not
        # # kick to the trap statement; they simply continue on.
        # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
        #
        # # Do something that might trigger an error
        # Get-Item -Path 'C:\MayNotExist.txt'
        #
        # # Restore the former error preference
        # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
        #
        # # Retrieve the newest error on the error stack
        # $refNewestCurrentError = Get-ReferenceToLastError
        #
        # $boolErrorOccurred = $false
        # if (($null -ne $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
        #     # Both not $null
        #     if (($refLastKnownError.Value) -ne ($refNewestCurrentError.Value)) {
        #         $boolErrorOccurred = $true
        #     }
        # } else {
        #     # One is $null, or both are $null
        #     # NOTE: $refLastKnownError could be non-null, while
        #     # $refNewestCurrentError could be null if $error was cleared;
        #     # this does not indicate an error.
        #     #
        #     # So:
        #     # If both are null, no error.
        #     # If $refLastKnownError is null and $refNewestCurrentError is
        #     # non-null, error.
        #     # If $refLastKnownError is non-null and $refNewestCurrentError
        #     # is null, no error.
        #     #
        #     if (($null -eq $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
        #         $boolErrorOccurred = $true
        #     }
        # }
        #
        # .INPUTS
        # None. You can't pipe objects to Get-ReferenceToLastError.
        #
        # .OUTPUTS
        # System.Management.Automation.PSReference ([ref]).
        # Get-ReferenceToLastError returns a reference (memory pointer) to
        # the last error that occurred. It returns a reference to $null
        # ([ref]$null) if there are no errors on on the $error stack.
        #
        # .NOTES
        # Version: 2.0.20250215.0

        #region License ################################################
        # Copyright (c) 2025 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person
        # obtaining a copy of this software and associated documentation
        # files (the "Software"), to deal in the Software without
        # restriction, including without limitation the rights to use,
        # copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the
        # Software is furnished to do so, subject to the following
        # conditions:
        #
        # The above copyright notice and this permission notice shall be
        # included in all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
        # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
        # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
        # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
        # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
        # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
        # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
        # OTHER DEALINGS IN THE SOFTWARE.
        #endregion License ################################################

        if ($Error.Count -gt 0) {
            return ([ref]($Error[0]))
        } else {
            return ([ref]$null)
        }
    }

    function Test-ErrorOccurred {
        # .SYNOPSIS
        # Checks to see if an error occurred during a time period, i.e.,
        # during the execution of a command.
        #
        # .DESCRIPTION
        # Using two references (memory pointers) to errors, this function
        # checks to see if an error occurred based on differences between
        # the two errors.
        #
        # To use this function, you must first retrieve a reference to the
        # last error that occurred prior to the command you are about to
        # run. Then, run the command. After the command completes, retrieve
        # a reference to the last error that occurred. Pass these two
        # references to this function to determine if an error occurred.
        #
        # .PARAMETER ReferenceToEarlierError
        # This parameter is required; it is a reference (memory pointer) to
        # a System.Management.Automation.ErrorRecord that represents the
        # newest error on the stack earlier in time, i.e., prior to running
        # the command for which you wish to determine whether an error
        # occurred.
        #
        # If no error was on the stack at this time,
        # ReferenceToEarlierError must be a reference to $null
        # ([ref]$null).
        #
        # .PARAMETER ReferenceToLaterError
        # This parameter is required; it is a reference (memory pointer) to
        # a System.Management.Automation.ErrorRecord that represents the
        # newest error on the stack later in time, i.e., after to running
        # the command for which you wish to determine whether an error
        # occurred.
        #
        # If no error was on the stack at this time, ReferenceToLaterError
        # must be a reference to $null ([ref]$null).
        #
        # .EXAMPLE
        # # Intentionally empty trap statement to prevent terminating
        # # errors from halting processing
        # trap { }
        #
        # # Retrieve the newest error on the stack prior to doing work
        # if ($Error.Count -gt 0) {
        #     $refLastKnownError = ([ref]($Error[0]))
        # } else {
        #     $refLastKnownError = ([ref]$null)
        # }
        #
        # # Store current error preference; we will restore it after we do
        # # some work:
        # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
        #
        # # Set ErrorActionPreference to SilentlyContinue; this will
        # # suppress error output. Terminating errors will not output
        # # anything, kick to the empty trap statement and then continue
        # # on. Likewise, non- terminating errors will also not output
        # # anything, but they do not kick to the trap statement; they
        # # simply continue on.
        # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
        #
        # # Do something that might trigger an error
        # Get-Item -Path 'C:\MayNotExist.txt'
        #
        # # Restore the former error preference
        # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
        #
        # # Retrieve the newest error on the error stack
        # if ($Error.Count -gt 0) {
        #     $refNewestCurrentError = ([ref]($Error[0]))
        # } else {
        #     $refNewestCurrentError = ([ref]$null)
        # }
        #
        # if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
        #     # Error occurred
        # } else {
        #     # No error occurred
        # }
        #
        # .INPUTS
        # None. You can't pipe objects to Test-ErrorOccurred.
        #
        # .OUTPUTS
        # System.Boolean. Test-ErrorOccurred returns a boolean value
        # indicating whether an error occurred during the time period in
        # question. $true indicates an error occurred; $false indicates no
        # error occurred.
        #
        # .NOTES
        # This function also supports the use of positional parameters
        # instead of named parameters. If positional parameters are used
        # instead of named parameters, then two positional parameters are
        # required:
        #
        # The first positional parameter is a reference (memory pointer) to
        # a System.Management.Automation.ErrorRecord that represents the
        # newest error on the stack earlier in time, i.e., prior to running
        # the command for which you wish to determine whether an error
        # occurred. If no error was on the stack at this time, the first
        # positional parameter must be a reference to $null ([ref]$null).
        #
        # The second positional parameter is a reference (memory pointer)
        # to a System.Management.Automation.ErrorRecord that represents the
        # newest error on the stack later in time, i.e., after to running
        # the command for which you wish to determine whether an error
        # occurred. If no error was on the stack at this time,
        # ReferenceToLaterError must be a reference to $null ([ref]$null).
        #
        # Version: 2.0.20250215.0

        #region License ################################################
        # Copyright (c) 2025 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person
        # obtaining a copy of this software and associated documentation
        # files (the "Software"), to deal in the Software without
        # restriction, including without limitation the rights to use,
        # copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the
        # Software is furnished to do so, subject to the following
        # conditions:
        #
        # The above copyright notice and this permission notice shall be
        # included in all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
        # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
        # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
        # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
        # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
        # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
        # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
        # OTHER DEALINGS IN THE SOFTWARE.
        #endregion License ################################################
        param (
            [ref]$ReferenceToEarlierError = ([ref]$null),
            [ref]$ReferenceToLaterError = ([ref]$null)
        )

        # TODO: Validate input

        $boolErrorOccurred = $false
        if (($null -ne $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
            # Both not $null
            if (($ReferenceToEarlierError.Value) -ne ($ReferenceToLaterError.Value)) {
                $boolErrorOccurred = $true
            }
        } else {
            # One is $null, or both are $null
            # NOTE: $ReferenceToEarlierError could be non-null, while
            # $ReferenceToLaterError could be null if $error was cleared;
            # this does not indicate an error.
            # So:
            # - If both are null, no error.
            # - If $ReferenceToEarlierError is null and
            #   $ReferenceToLaterError is non-null, error.
            # - If $ReferenceToEarlierError is non-null and
            #   $ReferenceToLaterError is null, no error.
            if (($null -eq $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                $boolErrorOccurred = $true
            }
        }

        return $boolErrorOccurred
    }
    #endregion FunctionsToSupportErrorHandling ####################################

    function Get-PSVersion {
        # .SYNOPSIS
        # Returns the version of PowerShell that is running.
        #
        # .DESCRIPTION
        # The function outputs a [version] object representing the version of
        # PowerShell that is running.
        #
        # On versions of PowerShell greater than or equal to version 2.0, this
        # function returns the equivalent of $PSVersionTable.PSVersion
        #
        # PowerShell 1.0 does not have a $PSVersionTable variable, so this
        # function returns [version]('1.0') on PowerShell 1.0.
        #
        # .EXAMPLE
        # $versionPS = Get-PSVersion
        # # $versionPS now contains the version of PowerShell that is running.
        # # On versions of PowerShell greater than or equal to version 2.0,
        # # this function returns the equivalent of $PSVersionTable.PSVersion.
        #
        # .INPUTS
        # None. You can't pipe objects to Get-PSVersion.
        #
        # .OUTPUTS
        # System.Version. Get-PSVersion returns a [version] value indiciating
        # the version of PowerShell that is running.
        #
        # .NOTES
        # Version: 1.0.20250106.0

        #region License ####################################################
        # Copyright (c) 2025 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining
        # a copy of this software and associated documentation files (the
        # "Software"), to deal in the Software without restriction, including
        # without limitation the rights to use, copy, modify, merge, publish,
        # distribute, sublicense, and/or sell copies of the Software, and to
        # permit persons to whom the Software is furnished to do so, subject to
        # the following conditions:
        #
        # The above copyright notice and this permission notice shall be
        # included in all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
        # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
        # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
        # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
        # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
        # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
        # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ####################################################

        if (Test-Path variable:\PSVersionTable) {
            return ($PSVersionTable.PSVersion)
        } else {
            return ([version]('1.0'))
        }
    }

    trap {
        # Intentionally left empty to prevent terminating errors from halting
        # processing
    }

    if ([string]::IsNullOrEmpty($ReferenceToAzureOpenAIEndpoint.Value)) {
        Write-Error 'Get-AzureOpenAISingleChatResponseRobust must be called with the ReferenceToAzureOpenAIEndpoint parameter, which is a memory reference to a string containing the endpoint for the Azure OpenAI service. To view the endpoint, for an Azure OpenAI resource, go to the Azure portal and select the resource. Then, navigate to "Keys and Endpoint" in the left-hand menu. The endpoint will be in the format https://<resource-name>.openai.azure.com/ where <resource-name> is the name of the Azure OpenAI resource. Supply the complete endpoint URL, including the https:// prefix, the .openai.azure.com suffix, and the trailing slash.'
        return $false
    }
    if ($ReferenceToAzureOpenAIEndpoint.Value.Substring($ReferenceToAzureOpenAIEndpoint.Value.Length - 1) -ne '/') {
        Write-Error 'Get-AzureOpenAISingleChatResponseRobust must be called with the ReferenceToAzureOpenAIEndpoint parameter, which is a memory reference to a string containing the endpoint for the Azure OpenAI service. To view the endpoint, for an Azure OpenAI resource, go to the Azure portal and select the resource. Then, navigate to "Keys and Endpoint" in the left-hand menu. The endpoint will be in the format https://<resource-name>.openai.azure.com/ where <resource-name> is the name of the Azure OpenAI resource. Supply the complete endpoint URL, including the https:// prefix, the .openai.azure.com suffix, and the trailing slash.'
        return $false
    }
    if ([string]::IsNullOrEmpty($ReferenceToAzureOpenAIDeploymentName.Value)) {
        Write-Error 'Get-AzureOpenAISingleChatResponseRobust must be called with the ReferenceToAzureOpenAIDeploymentName parameter, which is a memory reference to a string that specifies the deployment name in the Azure OpenAI service instance that represents the chat response model to be used. The model deployments can be viewed in Azure AI Foundry. To view the model deployments, go to https://ai.azure.com/resource/deployments, then verify that the correct Azure OpenAI instance is selected at the top. The model deployments are listed in the middle pane. For this parameter, supply the name of the deployment that represents the chat response model to be used. The deployment name is case-sensitive.'
        return $false
    }
    if ([string]::IsNullOrEmpty($ReferenceToAPIKey.Value)) {
        Write-Error 'Get-AzureOpenAISingleChatResponseRobust must be called with the ReferenceToAPIKey parameter, which is a memory reference (pointer) to a string containing a valid Azure OpenAI API key that the function will use to submit the text (prompt) and receive the chat response.'
        return $false
    }
    if ([string]::IsNullOrEmpty($ReferenceToPrompt.Value)) {
        Write-Error 'Get-AzureOpenAISingleChatResponseRobust must be called with the ReferenceToPrompt parameter, which is a memory reference to a string containing the text that the function will prompt to (ask) the chat API. The text must be a non-empty string.'
        return $false
    }
    if (-not $ReferenceToResponse.Value.GetType().FullName -eq 'System.String') {
        Write-Error 'Get-AzureOpenAISingleChatResponseRobust must be called with the ReferenceToResponse parameter, which is a memory reference to a string that will be used to store the chat response retrieved from the Azure OpenAI service.'
        return $false
    }

    if ($PSVersion -eq ([version]'0.0')) {
        # If the PSVersion parameter is not supplied, set it to the current version
        $versionPS = $PSVersionTable.PSVersion
        $refPSVersion = [ref]$versionPS
    } else {
        $refPSVersion = [ref]$PSVersion
    }

    if ($refPSVersion.Value.Major -lt 3) {
        Write-Error 'Get-AzureOpenAISingleChatResponseRobust requires PowerShell v3.0 or newer. Please upgrade PowerShell, upgrade your operating system, or run this script from another computer.'
        return $false
    }

    $strDescriptionOfWhatWeAreDoingInThisFunction = 'retrieving a response from the ChatGPT (Azure OpenAI chat completions) API'

    $hashtableAzureOpenAIHeaders = [ordered]@{
        'api-key' = $ReferenceToAPIKey.Value
    }

    $arrMessages = @(
        @{
            'role' = 'system'
            'content' = 'You are ChatGPT, a large language model trained by OpenAI.'
        },
        @{
            'role' = 'user'
            'content' = $ReferenceToPrompt.Value
        }
    )

    $strJSONRequestBody = @{
        messages = $arrMessages
    } | ConvertTo-Json
    $strURL = $ReferenceToAzureOpenAIEndpoint.Value + 'openai/deployments/' + $ReferenceToAzureOpenAIDeploymentName.Value + '/chat/completions?api-version=' + $ReferenceToAzureOpenAIAPIVersion.Value

    # Retrieve the newest error on the stack prior to doing work
    $refLastKnownError = Get-ReferenceToLastError

    # Store current error preference; we will restore it after we do the work of
    # this function
    $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

    # Set ErrorActionPreference to SilentlyContinue; this will suppress error
    # output. Terminating errors will not output anything, kick to the empty trap
    # statement and then continue on. Likewise, non-terminating errors will also
    # not output anything, but they do not kick to the trap statement; they simply
    # continue on.
    $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

    # Call the Azure OpenAI API to retrieve the chat response
    $PSCustomObjectResponse = Invoke-RestMethod -Uri $strURL -Headers $hashtableAzureOpenAIHeaders -Method Post -Body $strJSONRequestBody -ContentType 'application/json' -TimeoutSec 180 -DisableKeepAlive

    # Restore the former error preference
    $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

    # Retrieve the newest error on the error stack
    $refNewestCurrentError = Get-ReferenceToLastError

    if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
        # Error occurred
        if ($CurrentAttemptNumber -lt $MaxAttempts) {
            Write-Verbose ("An error occurred " + $strDescriptionOfWhatWeAreDoingInThisFunction + ". Waiting for " + [string]([math]::Pow(2, $CurrentAttemptNumber)) + " seconds, then retrying...")
            Start-Sleep -Seconds ([math]::Pow(2, $CurrentAttemptNumber))

            $objResultIndicator = Get-AzureOpenAISingleChatResponseRobust -ReferenceToResponse $ReferenceToResponse -CurrentAttemptNumber ($CurrentAttemptNumber + 1) -MaxAttempts $MaxAttempts -ReferenceToAzureOpenAIEndpoint $ReferenceToAzureOpenAIEndpoint -ReferenceToAzureOpenAIDeploymentName $ReferenceToAzureOpenAIDeploymentName -ReferenceToAPIKey $ReferenceToAPIKey -ReferenceToPrompt $ReferenceToPrompt -ReferenceToAzureOpenAIAPIVersion $ReferenceToAzureOpenAIAPIVersion -PSVersion $refPSVersion.Value
            return $objResultIndicator
        } else {
            # Number of attempts exceeded maximum
            if ($MaxAttempts -ge 2) {
                Write-Verbose ("An error occurred " + $strDescriptionOfWhatWeAreDoingInThisFunction + ". Giving up after too many attempts!")
            } else {
                Write-Verbose ("An error occurred " + $strDescriptionOfWhatWeAreDoingInThisFunction + ".")
            }

            return $false
        }
    } else {
        # No error occurred
        $boolChatResponseReturned = $false
        if (-not [string]::IsNullOrEmpty(@($PSCustomObjectResponse.choices)[0].message.content)) {
            # A chat response was returned
            $boolChatResponseReturned = $true
        }

        if (-not $boolChatResponseReturned) {
            if ($CurrentAttemptNumber -lt $MaxAttempts) {
                Write-Verbose ("An error occurred " + $strDescriptionOfWhatWeAreDoingInThisFunction + ". Waiting for " + [string]([math]::Pow(2, $CurrentAttemptNumber)) + " seconds, then retrying...")
                Start-Sleep -Seconds ([math]::Pow(2, $CurrentAttemptNumber))

                $objResultIndicator = Get-AzureOpenAISingleChatResponseRobust -ReferenceToResponse $ReferenceToResponse -CurrentAttemptNumber ($CurrentAttemptNumber + 1) -MaxAttempts $MaxAttempts -ReferenceToAzureOpenAIEndpoint $ReferenceToAzureOpenAIEndpoint -ReferenceToAzureOpenAIDeploymentName $ReferenceToAzureOpenAIDeploymentName -ReferenceToAPIKey $ReferenceToAPIKey -ReferenceToPrompt $ReferenceToPrompt -ReferenceToAzureOpenAIAPIVersion $ReferenceToAzureOpenAIAPIVersion -PSVersion $refPSVersion.Value
                return $objResultIndicator
            } else {
                # Number of attempts exceeded maximum
                if ($MaxAttempts -ge 2) {
                    Write-Verbose ("An error occurred " + $strDescriptionOfWhatWeAreDoingInThisFunction + ". Giving up after too many attempts!")
                } else {
                    Write-Verbose ("An error occurred " + $strDescriptionOfWhatWeAreDoingInThisFunction + ".")
                }

                return $false
            }
        } else {
            # A chat response was returned successfully!
            $ReferenceToResponse.Value = @($PSCustomObjectResponse.choices)[0].message.content.Trim()
            return $true
        }
    }
}

function Get-SingleAzureOpenAIChatResponse {
    # Get-SingleAzureOpenAIChatResponse
    # Version: 1.0.20250412.0

    <#
    .SYNOPSIS
    This function submits text (a prompt) to Azure OpenAI and returns the chat
    response

    .DESCRIPTION
    This function submits text (a prompt) to Azure OpenAI and returns the chat
    response. The function also includes parameters to customize the Azure OpenAI
    endpoint, deployment name, API key, and other settings. The function is
    designed to be used in Windows PowerShell 3 and later versions.

    .PARAMETER MaximumAttempts
    Specifies the maximum number of attempts that the function will observe before
    giving up. The default value is 8.

    .PARAMETER AzureOpenAIEndpoint
    A string containing the endpoint for the Azure OpenAI service. The endpoint
    should be in the format 'https://<resource-name>.openai.azure.com/' where
    <resource-name> is the name of the Azure OpenAI resource. The endpoint must
    include the 'https://' prefix, the '.openai.azure.com' suffix, and the trailing
    slash.

    .PARAMETER AzureOpenAIDeploymentName
    A string that specifies the deployment name in the Azure OpenAI service instance
    that represents the chat response model to be used. The model deployments can be
    viewed in Azure AI Foundry. To view the model deployments, go to
    https://ai.azure.com/resource/deployments, then verify that the correct Azure
    OpenAI instance is selected at the top. The model deployments are listed in the
    middle pane. For this parameter, supply the name of the deployment that
    represents the chat response model to be used. The deployment name is case-
    sensitive.

    .PARAMETER AzureOpenAIAPIKey
    A string containing a valid Azure OpenAI API key that the function will use to
    communicate with the Azure OpenAI API.

    .PARAMETER AzureOpenAIAPIVersion
    A string that specifies the API version to use when connecting to the Azure OpenAI
    service. The API version is supplied in YYYY-MM-DD format, and, if this parameter
    is omitted, the script defaults to version 2024-12-01-preview. The latest GA API
    version can be viewed here:
    https://learn.microsoft.com/en-us/azure/ai-services/openai/api-version-deprecation?source=recommendations#latest-ga-api-release

    .PARAMETER Prompt
    Specifies the text that the function will send to Azure OpenAI's chat completions
    API.

    .PARAMETER PSVersion
    This parameter is optional; if supplied, it is a System.Version object that
    contains the version of PowerShell that is running. If this parameter is
    omitted, the function will determine the version of PowerShell that is
    currently running. This parameter is used to determine whether the script is
    running on a supported version of PowerShell.

    .EXAMPLE
    PS C:\> $strResponse = Get-SingleAzureOpenAIChatResponse -MaximumAttempts 3 -AzureOpenAIEndpoint 'https://my-openai-resource.openai.azure.com/' -AzureOpenAIDeploymentName 'gpt-4-turbo-preview' -AzureOpenAIAPIKey 'my-api-key' -AzureOpenAIAPIVersion '2024-12-01-preview' -Prompt 'What is the capital of France?'
    PS C:\> $strResponse

    .OUTPUTS
    The function returns a string containing the response from Azure OpenAI's chat
    completions API
    #>

    #region License ############################################################
    # Copyright 2025 Frank Lesniak and Daniel Stutz
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy
    # of this software and associated documentation files (the "Software"), to deal
    # in the Software without restriction, including without limitation the rights
    # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    # copies of the Software, and to permit persons to whom the Software is
    # furnished to do so, subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in
    # all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    # SOFTWARE.
    #endregion License ############################################################

    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $false)][ValidateRange(1, 10)][int]$MaximumAttempts = 8,
        [Parameter(Mandatory = $true)][string]$AzureOpenAIEndpoint,
        [Parameter(Mandatory = $true)][string]$AzureOpenAIDeploymentName,
        [Parameter(Mandatory = $true)][string]$AzureOpenAIAPIKey,
        [Parameter(Mandatory = $false)][string]$AzureOpenAIAPIVersion = '2024-12-01-preview',
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $false)][version]$PSVersion = ([version]'0.0')
    )

    $strResponse = ''
    $boolSuccess = Get-AzureOpenAISingleChatResponseRobust -ReferenceToResponse ([ref]$strResponse) -CurrentAttemptNumber 1 -MaxAttempts $MaximumAttempts -ReferenceToAzureOpenAIEndpoint ([ref]$AzureOpenAIEndpoint) -ReferenceToAzureOpenAIDeploymentName ([ref]$AzureOpenAIDeploymentName) -ReferenceToAPIKey ([ref]$AzureOpenAIAPIKey) -ReferenceToPrompt ([ref]$Prompt) -ReferenceToAzureOpenAIAPIVersion ([ref]$AzureOpenAIAPIVersion) -PSVersion $PSVersion
    if ($boolSuccess -eq $true) {
        return $strResponse
    } else {
        return ''
    }
}

function Get-TopicFromDataSetUsingAzureOpenAI {
    # Get-TopicFromDataSetUsingAzureOpenAI
    # Version: 1.0.20250412.0

    <#
    .SYNOPSIS
    Uses Azure OpenAI to generate a topic for a set of unstructured text data.

    .DESCRIPTION
    This function uses Azure OpenAI to generate a topic for a set of unstructured
    text data. The function is designed to be used in conjunction with the Azure OpenAI
    chat completions API.

    .PARAMETER MaximumAttempts
    Specifies the maximum number of attempts that the function will observe before
    giving up. The default value is 8.

    .PARAMETER AzureOpenAIEndpoint
    A string containing the endpoint for the Azure OpenAI service. The endpoint
    should be in the format 'https://<resource-name>.openai.azure.com/' where
    <resource-name> is the name of the Azure OpenAI resource. The endpoint must
    include the 'https://' prefix, the '.openai.azure.com' suffix, and the trailing
    slash.

    .PARAMETER AzureOpenAIDeploymentName
    A string that specifies the deployment name in the Azure OpenAI service instance
    that represents the chat response model to be used. The model deployments can be
    viewed in Azure AI Foundry. To view the model deployments, go to
    https://ai.azure.com/resource/deployments, then verify that the correct Azure
    OpenAI instance is selected at the top. The model deployments are listed in the
    middle pane. For this parameter, supply the name of the deployment that
    represents the chat response model to be used. The deployment name is case-
    sensitive.

    .PARAMETER AzureOpenAIAPIKey
    A string containing a valid Azure OpenAI API key that the function will use to
    communicate with the Azure OpenAI API.

    .PARAMETER AzureOpenAIAPIVersion
    A string that specifies the API version to use when connecting to the Azure OpenAI
    service. The API version is supplied in YYYY-MM-DD format, and, if this parameter
    is omitted, the script defaults to version 2024-12-01-preview. The latest GA API
    version can be viewed here:
    https://learn.microsoft.com/en-us/azure/ai-services/openai/api-version-deprecation?source=recommendations#latest-ga-api-release

    .PARAMETER ReferenceToArrayOfUnstructuredTextData
    A reference to an array of strings containing the unstructured text data that the
    function will use in its prompt to Azure OpenAI to generate the topic.

    .PARAMETER PSVersion
    This parameter is optional; if supplied, it is a System.Version object that
    contains the version of PowerShell that is running. If this parameter is
    omitted, the function will determine the version of PowerShell that is
    currently running. This parameter is used to determine whether the script is
    running on a supported version of PowerShell.

    .EXAMPLE
    PS C:\> $arrTexts = @()
    PS C:\> $arrTexts += 'I would really like to be able to work out during the work day but there are no showers at my office.'
    PS C:\> $arrTexts += 'Having exercise equipment in the building is great!'
    PS C:\> $arrTexts += 'Someone stole my iPad from the exercise room and the building manager is not doing anything about it.'
    PS C:\> $arrTexts += 'Can we work out during our lunch break? I know some people do it, but I''m not sure it''s permitted. Can leadership make an announcement about this?'
    PS C:\> $strTopic = Get-TopicFromDataSetUsingAzureOpenAI -MaximumAttempts 3 -AzureOpenAIEndpoint 'https://my-openai-resource.openai.azure.com/' -AzureOpenAIDeploymentName 'gpt-4-turbo-preview' -AzureOpenAIAPIKey 'my-api-key' -AzureOpenAIAPIVersion '2024-12-01-preview' -ReferenceToArrayOfUnstructuredTextData ([ref]$arrTexts)

    This example generates a topic for a set of unstructured text data. The function
    will make up to 3 attempts to retrieve a response.  The unstructured text data is
    specified in the -ReferenceToArrayOfUnstructuredTextData parameter.

    .OUTPUTS
    The function returns a string containing the topic generated by Azure OpenAI.
    #>

    #region License ############################################################
    # Copyright 2025 Frank Lesniak and Daniel Stutz
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy
    # of this software and associated documentation files (the "Software"), to deal
    # in the Software without restriction, including without limitation the rights
    # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    # copies of the Software, and to permit persons to whom the Software is
    # furnished to do so, subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in
    # all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    # SOFTWARE.
    #endregion License ############################################################

    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $false)][ValidateRange(1, 10)][int]$MaximumAttempts = 8,
        [Parameter(Mandatory = $true)][string]$AzureOpenAIEndpoint,
        [Parameter(Mandatory = $true)][string]$AzureOpenAIDeploymentName,
        [Parameter(Mandatory = $true)][string]$AzureOpenAIAPIKey,
        [Parameter(Mandatory = $false)][string]$AzureOpenAIAPIVersion = '2024-12-01-preview',
        [Parameter(Mandatory = $true)][ref]$ReferenceToArrayOfUnstructuredTextData,
        [Parameter(Mandatory = $false)][version]$PSVersion = ([version]'0.0')
    )

    $intCountOfItems = @($ReferenceToArrayOfUnstructuredTextData.Value).Count

    if ($intCountOfItems -lt 1) {
        Write-Warning 'The array of unstructured text data is empty.'
        return ''
    }

    if ($intCountOfItems -eq 1) {
        $strPrompt = 'In as few words as possible (certainly no more than 1-3 words), describe the topic, main idea, or theme of the following text: ' + @($ReferenceToArrayOfUnstructuredTextData.Value)[0]
    } else {
        # More than one item in a set
        switch ($intCountOfItems) {
            2 { $strNumberOfItems = 'two' }
            3 { $strNumberOfItems = 'three' }
            4 { $strNumberOfItems = 'four' }
            5 { $strNumberOfItems = 'five' }
            6 { $strNumberOfItems = 'six' }
            7 { $strNumberOfItems = 'seven' }
            8 { $strNumberOfItems = 'eight' }
            9 { $strNumberOfItems = 'nine' }
            10 { $strNumberOfItems = 'ten' }
            11 { $strNumberOfItems = 'eleven' }
            12 { $strNumberOfItems = 'twelve' }
            13 { $strNumberOfItems = 'thirteen' }
            14 { $strNumberOfItems = 'fourteen' }
            15 { $strNumberOfItems = 'fifteen' }
            Default { $strNumberOfItems = $intCountOfItems.ToString() }
        }

        $strSeparator = '///'
        $boolSeparatorConflict = $false
        for ($intCounterA = 0; $intCounterA -lt $intCountOfItems; $intCounterA++) {
            if (@($ReferenceToArrayOfUnstructuredTextData.Value)[$intCounterA].Contains($strSeparator) -eq $true) {
                $boolSeparatorConflict = $true
                break
            }
        }
        if ($boolSeparatorConflict -eq $true) {
            $strSeparator = '|||'
            $boolSeparatorConflict = $false
            for ($intCounterA = 0; $intCounterA -lt $intCountOfItems; $intCounterA++) {
                if (@($ReferenceToArrayOfUnstructuredTextData.Value)[$intCounterA].Contains($strSeparator) -eq $true) {
                    $boolSeparatorConflict = $true
                    break
                }
            }
            if ($boolSeparatorConflict -eq $true) {
                $strSeparator = '###'
                $boolSeparatorConflict = $false
                for ($intCounterA = 0; $intCounterA -lt $intCountOfItems; $intCounterA++) {
                    if (@($ReferenceToArrayOfUnstructuredTextData.Value)[$intCounterA].Contains($strSeparator) -eq $true) {
                        $boolSeparatorConflict = $true
                        break
                    }
                }
                if ($boolSeparatorConflict -eq $true) {
                    $strSeparator = '///'
                    do {
                        $strSeparator += '///'
                        $boolSeparatorConflict = $false
                        for ($intCounterA = 0; $intCounterA -lt $intCountOfItems; $intCounterA++) {
                            if (@($ReferenceToArrayOfUnstructuredTextData.Value)[$intCounterA].Contains($strSeparator) -eq $true) {
                                $boolSeparatorConflict = $true
                                break
                            }
                        }
                    } while ($boolSeparatorConflict -eq $true)
                }
            }
        }

        if ($strSeparator -eq '///') {
            $strSeparatorText = 'three forward slashes (///)'
        } elseif ($strSeparator -eq '|||') {
            $strSeparatorText = 'three vertical bars (|||)'
        } elseif ($strSeparator -eq '###') {
            $strSeparatorText = 'three hash symbols (###)'
        } else {
            $strSeparatorText = 'exactly ' + $strSeparator.Length.ToString() + ' forward slashes (' + $strSeparator + ')'
        }

        $strPrompt = 'In as few words as possible (certainly no more than 1-3 words), describe the topic, main idea, or theme of the following ' + $strNumberOfItems + ' texts, treated as a set. Each text is separated by ' + $strSeparatorText + ': '
        $strPrompt += @($ReferenceToArrayOfUnstructuredTextData.Value) -join $strSeparator
    }

    $strTopic = Get-SingleAzureOpenAIChatResponse -MaximumAttempts $MaximumAttempts -AzureOpenAIEndpoint $AzureOpenAIEndpoint -AzureOpenAIDeploymentName $AzureOpenAIDeploymentName -AzureOpenAIAPIKey $AzureOpenAIAPIKey -AzureOpenAIAPIVersion $AzureOpenAIAPIVersion -Prompt $strPrompt -PSVersion $PSVersion

    # Remove quotation marks (") from the beginning and end of the string, if they
    # exist. Also include the "fancy" quotation marks like [char]8220 (“) and
    # [char]8221 (”) and [char]8222 („)
    $strTopic = $strTopic.Trim('"', [char]8220, [char]8221, [char]8222)

    # Remove a trailing period from the end of the string
    if ($strTopic.EndsWith('.')) {
        $strTopic = $strTopic.Substring(0, $strTopic.Length - 1)
    }
    return $strTopic
}

function Copy-Object {
    #region FunctionHeader #########################################################
    # This function is needed because PowerShell does not simply let you copy an object
    # of any complexity to another using the equal sign operator. If you were to use
    # the equal sign operator, it copies the pointer to the object - meaning that the
    # two variables that you thought were copies of each other are actually the same
    # object.
    #
    # Four positional arguments are required:
    #
    # The first argument is a reference to an output object to which the copy will
    # occur.
    #
    # The second argument is a reference to a source object from which the copy will
    # occur.
    #
    # The third argument is optional. If specified, it is an integer between 1 and
    # [int]::MaxValue that indicates the depth of the copy operation (i.e., how many
    # levels deep to copy nested objects, recursively). If set to $null or not
    # specified, the default copy depth is 2.
    #
    # The fourth argument is optional. If specified, it is a boolean indicating whether
    # the source object should be considered "safe", i.e., generated from a trusted
    # process and not possible to contain malicious code (see notes). If set to $null
    # or not specified, the default is $false.
    #
    # The function returns an integer indicating the success/failure of the process:
    #
    # 0 indicates success, that the source object was marked as serializable, the
    # function call indicated it was generated from a trusted process and not possible
    # to contain malicious code, and that the object was successfully copied using
    # BinaryFormatter - meaning that we are pretty well guaranteed that the source
    # object was copied in its entirety.
    #
    # 1 indicates success, but that the source object was either not marked as
    # serializable, or that it was not indicated to be generated from a trusted
    # process/possible to contain malicious code. In this case, the copy depth was used
    # to determine how "deeply" to copy nested objects. The destination object is not
    # guaranteed to be an exact copy of the source.
    #
    # 2 indicates failure; the object was not able to be copied
    #
    # Example usage:
    #
    # # Example 1: Fast object copy; this method *might* miss nested object data on
    # # complex objects
    # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
    # $DestinationObject = $null
    # $intReturnCode = Copy-Object ([ref]$DestinationObject) ([ref]$SourceObject) 1
    # # Note 1: @(@($SourceObject)[0].Modules)[0].Assembly is not equal to $null
    # # Note 2: @(@($DestinationObject)[0].Modules)[0].Assembly is equal to $null
    # # because the copy depth is 1
    # # Note 3: On the plus side, this example takes 0.5 - 2 seconds to complete -
    # # pretty fast.
    #
    # # Example 2: More-robust copy; this method can still miss nested object data on
    # # complex objects but is less likely to do so
    # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
    # $DestinationObject = $null
    # $intReturnCode = Copy-Object ([ref]$DestinationObject) ([ref]$SourceObject) 3
    # # Note 1: @(@($SourceObject)[0].Modules)[0].Assembly is not equal to $null
    # # Note 2: @(@($DestinationObject)[0].Modules)[0].Assembly is also not equal to
    # # $null because we copied the object "deeply enough".
    # # Note 3: This command could take approximately 6-20 minutes to complete (!).
    #
    # # Example 3: Copy an object that is marked as serializable and generated from a
    # # trusted process:
    # $SourceObject = New-Object 'System.Collections.Generic.List[System.String]'
    # for ($intCounter = 1; $intCounter -le 10000; $intCounter++) {
    #     $SourceObject.Add('Item' + ([string]$intCounter))
    # }
    # $DestinationObject = $null
    # $intReturnCode = Copy-Object ([ref]$DestinationObject) ([ref]$SourceObject) 3 $true
    #
    # Note: Copying an object that is marked as serializable is possible in its
    # entirety is possible by setting the fourth parameter to $true, which will tell
    # this function to use BinaryFormatter to perform the copy. However,
    # BinaryFormatter has inherent security vulnerabilities and Microsoft has
    # discouraged its use. See:
    # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide
    #
    # Note: This function is compatible all the way back to PowerShell v1.0.
    #
    # Version: 1.0.20240127.0
    #endregion FunctionHeader #########################################################

    #region License ################################################################
    # Copyright (c) 2024 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################################

    #region DownloadLocationNotice #################################################
    # The most up-to-date version of this script can be found on the author's GitHub
    # repository at https://github.com/franklesniak/Copy-Object
    #endregion DownloadLocationNotice #################################################

    #region FunctionsToSupportErrorHandling ########################################
    function Get-ReferenceToLastError {
        #region FunctionHeader #####################################################
        # Function returns $null if no errors on on the $error stack;
        # Otherwise, function returns a reference (memory pointer) to the last error
        # that occurred.
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/PowerShell_Resources
        #endregion DownloadLocationNotice #############################################

        if ($error.Count -gt 0) {
            [ref]($error[0])
        } else {
            $null
        }
    }

    function Test-ErrorOccurred {
        #region FunctionHeader #####################################################
        # Function accepts two positional arguments:
        #
        # The first argument is a reference (memory pointer) to the last error that had
        # occurred prior to calling the command in question - that is, the command that
        # we want to test to see if an error occurred.
        #
        # The second argument is a reference to the last error that had occurred as-of
        # the completion of the command in question.
        #
        # Function returns $true if it appears that an error occurred; $false otherwise
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/PowerShell_Resources
        #endregion DownloadLocationNotice #############################################

        # TO-DO: Validate input

        $boolErrorOccurred = $false
        if (($null -ne ($args[0])) -and ($null -ne ($args[1]))) {
            # Both not $null
            if ((($args[0]).Value) -ne (($args[1]).Value)) {
                $boolErrorOccurred = $true
            }
        } else {
            # One is $null, or both are $null
            # NOTE: ($args[0]) could be non-null, while ($args[1])
            # could be null if $error was cleared; this does not indicate an error.
            # So:
            # If both are null, no error
            # If ($args[0]) is null and ($args[1]) is non-null, error
            # If ($args[0]) is non-null and ($args[1]) is null, no error
            if (($null -eq ($args[0])) -and ($null -ne ($args[1]))) {
                $boolErrorOccurred
            }
        }

        $boolErrorOccurred
    }
    #endregion FunctionsToSupportErrorHandling ########################################

    #region FunctionsForCopyingAnObject ############################################
    function Test-PSSerializerTypeAvailabilityWithoutEnumeratingDotNETAssemblies {
        #region FunctionHeader #####################################################
        # Function tests to determine whether the PSSerializer type is availabile by
        # attempting to create an object using PSSerializer and using it to perform a
        # simple operation. If the creation/operation fails, then PSSerializer is
        # presumed unavailable for use.
        #
        # This function does not accept any arguments or parameters.
        #
        # The function returns $true if System.Management.Automation.PSSerializer is
        # available, $false otherwise
        #
        # Example usage:
        # $boolResult = $null
        # $boolResult = Test-PSSerializerTypeAvailabilityWithoutEnumeratingDotNETAssemblies
        # if ($boolResult -eq $true) {
        #   # Do something with System.Management.Automation.PSSerializer
        # } else {
        #   # Don't do anything with System.Management.Automation.PSSerializer because
        #   # it does not exist
        # }
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        $strCliXMLDummyValue = $null

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Attempt to create and use a PSSerializer object:
        $strCliXMLDummyValue = [System.Management.Automation.PSSerializer]::Serialize(1, 1)

        # Restore the former error preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            $strCliXMLDummyValue = $null
            return $false # System.Management.Automation.PSSerializer did not exist
        } else {
            $strCliXMLDummyValue = $null
            return $true # System.Management.Automation.PSSerializer did exist
        }
    }

    function Invoke-BinaryFormatterSerializeOperation {
        #region FunctionHeader #####################################################
        # This function takes a serialize-able object and converts it to binary,
        # storing it into a memorystream. This is a useful part of cloning an object
        # marked as serialize-able. However, please note that the use of
        # BinaryFormatter is not recommended due to security vulnerabilities. See
        # notes.
        #
        # Three positional arguments are required:
        #
        # The first argument is a reference to a System.IO.MemoryStream that will be
        # used to store output.
        #
        # The second argument is a reference to a
        # System.Runtime.Serialization.Formatters.Binary.BinaryFormatter object, which
        # is used to perform the operation.
        #
        # The third argument is a reference to the source object that we are trying to
        # clone.
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        #
        # Example 1:
        # $memoryStream = New-Object System.IO.MemoryStream
        # $binaryFormatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        # $boolSuccess = Invoke-BinaryFormatterSerializeOperation ([ref]$memoryStream) ([ref]$binaryFormatter) ([ref]$objToClone)
        #
        # Example 2:
        # $SourceObject = New-Object 'System.Collections.Generic.List[System.String]'
        # for ($intCounter = 1; $intCounter -le 10000; $intCounter++) {
        #     $SourceObject.Add('Item' + ([string]$intCounter))
        # }
        # $DestinationObject = $null
        # $memoryStream = New-Object System.IO.MemoryStream
        # $binaryFormatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        # $boolSuccess = Invoke-BinaryFormatterSerializeOperation ([ref]$memoryStream) ([ref]$binaryFormatter) ([ref]$SourceObject)
        # if ($boolSuccess -eq $false) {
        #   Write-Warning -Message 'Failed to serialize object using BinaryFormatter.'
        # } else {
        #   $memoryStream.Position = 0
        #   $boolSuccess = Invoke-BinaryFormatterDeserializeOperation ([ref]$DestinationObject) ([ref]$binaryFormatter) ([ref]$memoryStream)
        #   if ($boolSuccess -eq $false) {
        #       Write-Warning -Message 'Failed to deserialize object using BinaryFormatter.'
        #   } else {
        #       Write-Host -Object 'Successfully used serialization (using BinaryFormatter) to clone an object.'
        #   }
        # }
        #
        # Note: The use of BinaryFormatter is not recommended due to security
        # vulnerabilities. See:
        # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide
        # The author recommends that BinaryFormatter not be used unless the input
        # object is marked as serializable and the input object is created from a
        # trusted process - not from external input or other untrusted processes/
        # sources.
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        # TODO: Validate input

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Use BinaryFormatter to perform the serialization
        (($args[1]).Value).Serialize((($args[0]).Value), (($args[2]).Value))

        # Restore the former error preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            return $false
        } else {
            return $true
        }
    }

    function Invoke-BinaryFormatterDeserializeOperation {
        #region FunctionHeader #####################################################
        # This function de-serializes an object, converting it from binary to a usable,
        # "rich" object, accessible within PowerShell or other system calls. This is a
        # useful part of cloning an object marked as serialize-able. However, please
        # note that the use of BinaryFormatter is not recommended due to security
        # vulnerabilities. See notes.
        #
        # Three positional arguments are required:
        #
        # The first argument is a reference to an object that will be used to store
        # output (i.e., store the copied object that we are creating as part of this
        # operation.
        #
        # The second argument is a reference to a
        # System.Runtime.Serialization.Formatters.Binary.BinaryFormatter object, which
        # is used to perform the operation.
        #
        # The third argument is a reference to a System.IO.MemoryStream that contains
        # the stream of serialized data that will be used to construct the object
        # specified by the first argument.
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        #
        # Example 1:
        # $objDeserializedObject = $null
        # $boolSuccess = Invoke-BinaryFormatterDeserializeOperation ([ref]$objDeserializedObject) ([ref]$binaryFormatter) ([ref]$memoryStream)
        #
        # Example 2:
        # $SourceObject = New-Object 'System.Collections.Generic.List[System.String]'
        # for ($intCounter = 1; $intCounter -le 10000; $intCounter++) {
        #     $SourceObject.Add('Item' + ([string]$intCounter))
        # }
        # $DestinationObject = $null
        # $memoryStream = New-Object System.IO.MemoryStream
        # $binaryFormatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        # $boolSuccess = Invoke-BinaryFormatterSerializeOperation ([ref]$memoryStream) ([ref]$binaryFormatter) ([ref]$SourceObject)
        # if ($boolSuccess -eq $false) {
        #   Write-Warning -Message 'Failed to serialize object using BinaryFormatter.'
        # } else {
        #   $memoryStream.Position = 0
        #   $boolSuccess = Invoke-BinaryFormatterDeserializeOperation ([ref]$DestinationObject) ([ref]$binaryFormatter) ([ref]$memoryStream)
        #   if ($boolSuccess -eq $false) {
        #       Write-Warning -Message 'Failed to deserialize object using BinaryFormatter.'
        #   } else {
        #       Write-Host -Object 'Successfully used serialization (using BinaryFormatter) to clone an object.'
        #   }
        # }
        #
        # Note: The use of BinaryFormatter is not recommended due to security
        # vulnerabilities. See:
        # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide
        # The author recommends that BinaryFormatter not be used unless the input
        # object is marked as serializable and the input object is created from a
        # trusted process - not from external input or other untrusted processes/
        # sources.
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        # TODO: Validate input

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Use BinaryFormatter to perform the de-serialization
        (($args[0]).Value) = (($args[1]).Value).Deserialize(($args[2]).Value)

        # Restore the former error preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            return $false
        } else {
            return $true
        }
    }

    function Invoke-PSSerializerSerializeOperation {
        #region FunctionHeader #####################################################
        # This function uses PSSerializer to serialize an object, generating an XML
        # represenation of the original object. This is useful when cloning an object
        # that is not marked as serializable because instead of straight-serializing
        # the object, the source object is converted to XML up to a specified copy
        # depth, i.e., number of nested objects. It's also one of the possible
        # recommended ways to serialize an object while avoiding security
        # vulnerabilities inherent in the use of BinaryFormatter. See notes.
        #
        # Three positional arguments are required:
        #
        # The first argument is a reference to a string object that will be used to
        # store output (i.e., the serialized XML description of the source object).
        #
        # The second argument is a reference to the source object that we are trying to
        # clone.
        #
        # The third argument is optional. If specified, it is an integer indicating the
        # copy depth (i.e., the depth of nested objects to recursively enumerate and
        # serialize before giving up). If not specified, the function defaults to a
        # copy depth of 2
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        #
        # Example 1:
        # $strCliXMLSourceObject = ''
        # $boolSuccess = Invoke-PSSerializerSerializeOperation ([ref]$strCliXMLSourceObject) ([ref]$objToClone)
        #
        # Example 2:
        # $boolResult = Test-PSSerializerTypeAvailabilityWithoutEnumeratingDotNETAssemblies
        # if ($boolResult) {
        #   $strCliXMLSourceObject = ''
        #   $boolResult = Invoke-PSSerializerSerializeOperation ([ref]$strCliXMLSourceObject) ([ref]$SourceObject)
        #   if ($boolResult -eq $false) {
        #       Write-Warning -Message 'Failed to serialize object using PSSerializer.'
        #   } else {
        #       $boolResult = Invoke-PSSerializerDeserializeOperation ([ref]$DestinationObject) ([ref]$strCliXMLSourceObject)
        #       if ($boolResult -eq $false) {
        #            Write-Warning -Message 'Failed to deserialize object using PSSerializer.'
        #       } else {
        #           Write-Host -Object 'Successfully used serialization (to and from XML using PSSerializer) to clone an object.'
        #       }
        #   }
        # } else {
        #   Write-Warning -Message 'System.Management.Automation.PSSerializer is not available.'
        # }
        #
        # Note: Serializing to XML/de-serializing from XML is a useful way to copy an
        # object while avoiding security vulnerabilities even when an object is marked
        # as serializable and therefore can be copied using BinaryFormatter. See:
        # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        # TODO: Additional input validation

        $intDefaultCopyDepth = 2

        if ($args.Count -ge 3) {
            if ($null -ne $args[2]) {
                if (($args[2].GetType().FullName -notlike 'System.Int*') -and ($args[2].GetType().FullName -notlike 'System.UInt*')) {
                    return $false # Indicate error
                } else {
                    if ($args[2] -lt 1) {
                        return $false # Indicate error
                    } else {
                        $intCopyDepth = $args[2]
                    }
                }
            } else {
                # Fourth argument was $null - default it to 2
                $intCopyDepth = $intDefaultCopyDepth
            }
        } else {
            # Fourth argument was not supplied - default it to 2
            $intCopyDepth = $intDefaultCopyDepth
        }

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Use PSSerializer to perform the serialization
        (($args[0]).Value) = [System.Management.Automation.PSSerializer]::Serialize((($args[1]).Value), $intCopyDepth)

        # Restore the former error preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            return $false
        } else {
            return $true
        }
    }

    function Invoke-PSSerializerDeserializeOperation {
        #region FunctionHeader #####################################################
        # This function uses PSSerializer to de-serialize an object, creating new
        # object from an XML represenation. This is useful when cloning an object that
        # is not marked as serializable because instead of straight-serializing the
        # object, the source object is converted to XML up to a specified copy depth,
        # i.e., number of nested objects. It's also one of the possible recommended
        # ways to serialize an object while avoiding security vulnerabilities inherent
        # in the use of BinaryFormatter. See notes.
        #
        # Two positional arguments are required:
        #
        # The first argument is a reference to an object that will be used to store
        # output (i.e., the deserialized object).
        #
        # The second argument is a reference to a string object that contains a
        # serialized XML description of an object.
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        # $objDeserializedObject = $null
        # $boolSuccess = Invoke-PSSerializerDeserializeOperation ([ref]$objDeserializedObject) ([ref]$strCliXMLSourceObject)
        #
        # Note: Serializing to XML/de-serializing from XML is a useful way to copy an
        # object while avoiding security vulnerabilities even when an object is marked
        # as serializable and therefore can be copied using BinaryFormatter. See:
        # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        # TODO: Validate input

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Use PSSerializer to perform the de-serialization
        (($args[0]).Value) = [System.Management.Automation.PSSerializer]::Deserialize(($args[1]).Value)

        # Restore the former error preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            return $false
        } else {
            return $true
        }
    }

    function Export-ClixmlSafely {
        #region FunctionHeader #####################################################
        # This function uses Export-Clixml to serialize an object, generating an XML
        # represenation of the original object, and storing it on disk. This is useful
        # when cloning an object that is not marked as serializable because instead of
        # straight-serializing the object, the source object is converted to XML up to
        # a specified copy depth, i.e., number of nested objects. It's also one of the
        # possible recommended ways to serialize an object while avoiding security
        # vulnerabilities inherent in the use of BinaryFormatter. See notes.
        #
        # Three positional arguments are required:
        #
        # The first argument is a string object that indicates the path to the file
        # that will be used to store the XML representation of the object specified by
        # the fourth argument.
        #
        # The second argument is a reference to a source object that is being exported
        # in an XML representation to the file path specified by the first argument.
        #
        # The third argument is optional. If specified, it is an integer indicating the
        # copy depth (i.e., the depth of nested objects to recursively enumerate and
        # serialize before giving up). If not specified, the function defaults to a
        # copy depth of 2.
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        #
        # Example 1:
        # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
        # $DestinationObject = $null
        # $strTempFilePath = [System.IO.Path]::GetTempFileName()
        # $boolSuccess = Export-ClixmlSafely $strTempFilePath ([ref]$SourceObject)
        # if ($boolSuccess -eq $true) {
        #   $DestinationObject = Import-Clixml -Path $strTempFilePath
        #   Remove-Item -Path $strTempFilePath -Force -ErrorAction SilentlyContinue
        # }
        #
        # Example 2:
        # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
        # $DestinationObject = $null
        # $strTempFilePath = [System.IO.Path]::GetTempFileName()
        # $boolSuccess = Export-ClixmlSafely $strTempFilePath ([ref]$SourceObject)
        # if ($boolSuccess -ne $true) {
        #     Write-Warning -Message 'Failed to export object to XML.'
        # } else {
        #     $boolSuccess = Import-ClixmlSafely ([ref]$DestinationObject) $strTempFilePath
        #     if ($boolSuccess -ne $true) {
        #         Write-Warning -Message 'Failed to import object from XML.'
        #     } else {
        #         Write-Host -Object 'Successfully used serialization (to and from XML using Import/Export-Clixml) to clone an object.'
        #     }
        # }
        #
        # Note: Serializing to XML/de-serializing from XML is a useful way to copy an
        # object while avoiding security vulnerabilities even when an object is marked
        # as serializable and therefore can be copied using BinaryFormatter. See:
        # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        # TODO: Additional input validation

        $intDefaultCopyDepth = 2

        if ($args.Count -ge 3) {
            if ($null -ne $args[2]) {
                if (($args[2].GetType().FullName -notlike 'System.Int*') -and ($args[2].GetType().FullName -notlike 'System.UInt*')) {
                    return $false # Indicate error
                } else {
                    if ($args[2] -lt 1) {
                        return $false # Indicate error
                    } else {
                        $intCopyDepth = $args[2]
                    }
                }
            } else {
                # Fourth argument was $null - default it to 2
                $intCopyDepth = $intDefaultCopyDepth
            }
        } else {
            # Fourth argument was not supplied - default it to 2
            $intCopyDepth = $intDefaultCopyDepth
        }

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Use Export-Clixml to perform the serialization, storing the result on disk
        (($args[1]).Value) | Export-Clixml -Path ($args[0]) -Depth $intCopyDepth -Force

        # Restore the former error preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            return $false
        } else {
            return $true
        }
    }

    function Import-ClixmlSafely {
        #region FunctionHeader #####################################################
        # This function uses Import-Clixml to de-serialize an object, creating new
        # object from an XML represenation. This is useful when cloning an object that
        # is not marked as serializable. It's also one of the possible recommended ways
        # to de-serialize an object while avoiding security vulnerabilities inherent in
        # the use of BinaryFormatter. See notes.
        #
        # Two positional arguments are required:
        #
        # The first argument is a reference to an object that will be used to store
        # output (i.e., the object constructed from the XML file representation).
        #
        # The second argument is a string indicating the path to the file that contains
        # the XML representation of the object to construct and store in the first
        # argument.
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        #
        # Example 1:
        # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
        # $DestinationObject = $null
        # $strTempFilePath = [System.IO.Path]::GetTempFileName()
        # $SourceObject | Export-Clixml -Path $strTempFilePath -Depth 2 -Force
        # if ((Test-Path -Path $strTempFilePath) -eq $true) {
        #   $boolSuccess = Import-ClixmlSafely ([ref]$DestinationObject) $strTempFilePath
        #   if ($boolSuccess -eq $true) {
        #       Remove-Item -Path $strTempFilePath -Force -ErrorAction SilentlyContinue
        #   }
        # }
        #
        # Example 2:
        # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
        # $DestinationObject = $null
        # $strTempFilePath = [System.IO.Path]::GetTempFileName()
        # $boolSuccess = Export-ClixmlSafely $strTempFilePath ([ref]$SourceObject)
        # if ($boolSuccess -ne $true) {
        #     Write-Warning -Message 'Failed to export object to XML.'
        # } else {
        #     $boolSuccess = Import-ClixmlSafely ([ref]$DestinationObject) $strTempFilePath
        #     if ($boolSuccess -ne $true) {
        #         Write-Warning -Message 'Failed to import object from XML.'
        #     } else {
        #         Write-Host -Object 'Successfully used serialization (to and from XML using Import/Export-Clixml) to clone an object.'
        #     }
        # }
        #
        # Note: Serializing to XML/de-serializing from XML is a useful way to copy an
        # object while avoiding security vulnerabilities even when an object is marked
        # as serializable and therefore can be copied using BinaryFormatter. See:
        # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        if ($args.Count -ge 2) {
            if ($args[1].GetType().FullName -eq 'System.String') {
                # Second argument is a string
                if ((Test-Path -Path ($args[1]) -ErrorAction SilentlyContinue) -eq $false) {
                    # Specified XML file does not exist
                    return $false # Indicate error
                }
            } else {
                # Second argument is not a string
                return $false # Indicate error
            }
        } else {
            # Required arguments were not provided
            return $false # Indicate error
        }

        # TODO: Perform additional input validation

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Use Import-Clixml to perform the de-serialization from disk
        (($args[0]).Value) = Import-Clixml -Path ($args[1])

        # Restore the former error preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            return $false
        } else {
            return $true
        }
    }

    function ConvertTo-JsonSafely {
        #region FunctionHeader #####################################################
        # This function uses ConvertTo-Json to serialize an object, generating a JSON
        # represenation of the original object. This is useful when cloning an object
        # that is not marked as serializable because instead of straight-serializing
        # the object, the source object is converted to JSON up to a specified copy
        # depth, i.e., number of nested objects. It's also one of the possible
        # recommended ways to serialize an object while avoiding security
        # vulnerabilities inherent in the use of BinaryFormatter. See notes.
        #
        # Three positional arguments are required:
        #
        # The first argument is a reference to a string object. At the completion of
        # this function, the string object will hold the JSON representation of the
        # object specified by the second argument.
        #
        # The second argument is a reference to a source object that is being
        # serialized into a JSON representation and stored in the string reference in
        # the first argument.
        #
        # The third argument is optional. If specified, it is an integer indicating the
        # copy depth (i.e., the depth of nested objects to recursively enumerate and
        # serialize before giving up). If not specified, the function defaults to a
        # copy depth of 2.
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        #
        # Example 1:
        # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
        # $DestinationObject = $null
        # $strJSONRepresenationOfSourceObject = ''
        # $boolSuccess = ConvertTo-JsonSafely ([ref]$strJSONRepresenationOfSourceObject) ([ref]$SourceObject) 2
        # if ($boolSuccess -eq $true) {
        #   $DestinationObject = ConvertFrom-Json -InputObject $strJSONRepresenationOfSourceObject
        # }
        #
        # Example 2:
        # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
        # $DestinationObject = $null
        # $strJSONRepresenationOfSourceObject = ''
        # $boolSuccess = ConvertTo-JsonSafely ([ref]$strJSONRepresenationOfSourceObject) ([ref]$SourceObject) 2
        # if ($boolSuccess -eq $true) {
        #   $boolSuccess = ConvertFrom-JsonSafely ([ref]$DestinationObject) ([ref]$strJSONRepresenationOfSourceObject)
        #   if ($boolSuccess -eq $false) {
        #       Write-Warning -Message 'Failed to convert JSON to object.'
        #   } else {
        #       Write-Host -Object 'Successfully converted JSON to object.'
        #   }
        # } else {
        #   Write-Warning -Message 'Failed to convert object to JSON.'
        # }
        #
        # Note: Exporting to JSON and importing from JSON is a useful way to avoid
        # security vulnerabilities even when an object is marked as serializable and
        # therefore can be copied using BinaryFormatter. See:
        # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        # TODO: Additional input validation
        # TODO: Change the return value to indicate if a warning occurred

        $intDefaultCopyDepth = 2

        if ($args.Count -ge 3) {
            if ($null -ne $args[2]) {
                if (($args[2].GetType().FullName -notlike 'System.Int*') -and ($args[2].GetType().FullName -notlike 'System.UInt*')) {
                    return $false # Indicate error
                } else {
                    if ($args[2] -lt 1) {
                        return $false # Indicate error
                    } else {
                        $intCopyDepth = $args[2]
                    }
                }
            } else {
                # Fourth argument was $null - default it to 2
                $intCopyDepth = $intDefaultCopyDepth
            }
        } else {
            # Fourth argument was not supplied - default it to 2
            $intCopyDepth = $intDefaultCopyDepth
        }

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Store current warning preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerWarningPreference = $global:WarningPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Set WarningPreference to SilentlyContinue; this will suppress warning output.
        $global:WarningPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Convert the object to JSON
        (($args[0]).Value) = ConvertTo-Json -InputObject (($args[1]).Value) -Depth $intCopyDepth

        # Restore the former warning preference
        $global:WarningPreference = $actionPreferenceFormerWarningPreference

        # Restore the former error action preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            return $false
        } else {
            return $true
        }
    }

    function ConvertFrom-JsonSafely {
        #region FunctionHeader #####################################################
        # This function uses ConvertFrom-Json to de-serialize an object, creating new
        # object from a JSON represenation. This is useful when cloning an object that
        # is not marked as serializable. It's also one of the possible recommended ways
        # to de-serialize an object while avoiding security vulnerabilities inherent in
        # the use of BinaryFormatter. See notes.
        #
        # Two positional arguments are required:
        #
        # The first argument is a reference to an object that will be used to store
        # output (i.e., the object constructed from the JSON file representation).
        #
        # The second argument is a reference to a string. The string must be a JSON-
        # formatted representation of an object to construct.
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        #
        # Example 1:
        # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
        # $DestinationObject = $null
        # $strJSONRepresenationOfSourceObject = ''
        # $intCopyDepth = 2
        # $strJSONRepresenationOfSourceObject = ConvertTo-Json -InputObject $SourceObject -Depth $intCopyDepth
        # if ($strJSONRepresenationOfSourceObject -ne '') {
        #   $boolSuccess = ConvertFrom-JsonSafely ([ref]$DestinationObject) ([ref]$strJSONRepresenationOfSourceObject)
        # }
        #
        # Example 2:
        # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
        # $DestinationObject = $null
        # $strJSONRepresenationOfSourceObject = ''
        # $boolSuccess = ConvertTo-JsonSafely ([ref]$strJSONRepresenationOfSourceObject) ([ref]$SourceObject) 2
        # if ($boolSuccess -eq $true) {
        #   $boolSuccess = ConvertFrom-JsonSafely ([ref]$DestinationObject) ([ref]$strJSONRepresenationOfSourceObject)
        #   if ($boolSuccess -eq $false) {
        #       Write-Warning -Message 'Failed to convert JSON to object.'
        #   } else {
        #       Write-Host -Object 'Successfully converted JSON to object.'
        #   }
        # } else {
        #   Write-Warning -Message 'Failed to convert object to JSON.'
        # }
        #
        # Note: Exporting to JSON and importing from JSON is a useful way to avoid
        # security vulnerabilities even when an object is marked as serializable and
        # therefore can be copied using BinaryFormatter. See:
        # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        # TODO: Validate input

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Do the work of this function...
        (($args[0]).Value) = ConvertFrom-Json -InputObject (($args[1]).Value)

        # Restore the former error preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            return $false
        } else {
            return $true
        }
    }

    function Copy-ObjectNotMarkedAsSerializable {
        #region FunctionHeader #####################################################
        # This function uses methods appropriate for objects not marked as serializable
        # to copy an object. For example, this function uses PSSerializer,
        # ConvertTo/ConvertFrom-Json, or Export/Import-Clixml. It's also one of the
        # possible recommended ways to de-serialize an object while avoiding security
        # vulnerabilities inherent in the use of BinaryFormatter. See notes.
        #
        # Three positional arguments are required:
        #
        # The first argument is a reference to an output object to which the copy will
        # occur.
        #
        # The second argument is a reference to a source object from which the copy
        # will occur.
        #
        # The third argument is optional. If specified, it is an integer between 1 and
        # [int]::MaxValue that indicates the depth of the copy operation (i.e., how
        # many levels deep to copy nested objects, recursively). If not specified, the
        # default copy depth is 2.
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        #
        # # Example 1: Faster object copy; this method will miss nested object data on
        # # complex objects:
        #
        # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
        # $DestinationObject = $null
        # $boolSuccess = Copy-ObjectNotMarkedAsSerializable ([ref]$DestinationObject) ([ref]$SourceObject) 1
        # # Note 1: @(@($SourceObject)[0].Modules)[0].Assembly is not equal to $null
        # # Note 2: @(@($DestinationObject)[0].Modules)[0].Assembly is equal to $null
        # # because the copy depth is 1
        # # Note 3: On the plus side, this example takes 0.5 - 2 seconds to complete -
        # # pretty fast.
        #
        # # Example 2: More-robust copy; this method can stil miss nested object data
        # # on complex objects but is less likely to do so
        # $SourceObject = @([AppDomain]::CurrentDomain.GetAssemblies())
        # $DestinationObject = $null
        # $boolSuccess = Copy-ObjectNotMarkedAsSerializable ([ref]$DestinationObject) ([ref]$SourceObject) 3
        # # Note 1: @(@($SourceObject)[0].Modules)[0].Assembly is not equal to $null
        # # Note 2: @(@($DestinationObject)[0].Modules)[0].Assembly is also not equal to
        # # $null because we copied the object "deeply enough".
        # # Note 3: This command could take approximately 6-20 minutes to complete (!).
        #
        # Note: Copying an object using the methods employed by this function is a
        # useful way to avoid security vulnerabilities even when an object is marked as
        # serializable and therefore can be copied using BinaryFormatter. See:
        # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        # Anything greater than 2 can be too slow to use as a default setting
        $intDefaultCopyDepth = 2

        if (($args.Count -lt 2) -or ($args.Count -gt 3)) {
            Write-Error 'Copy-ObjectNotMarkedAsSerializable was called with the incorrect number of arguments. The first positional argument is required, representing a reference to the destination (new) object. The second positional argument is also required, representing a reference to the source object to be copied. Optionally, a third positional argument may be specified as an integer between 1 and [int]::MaxValue that indicates the depth of the copy operation (i.e., how many levels deep to copy nested objects, recursively)'
            return $false
        }

        if ((($args[0]).GetType().FullName -notlike 'System.Management.Automation.PSReference*') -or (($args[1]).GetType().FullName -notlike 'System.Management.Automation.PSReference*')) {
            Write-Error 'Copy-ObjectNotMarkedAsSerializable must be called with at least two arguments. The first positional argument is required, representing a reference to the destination (new) object. The second positional argument is also required, representing a reference to the source object to be copied. For example: $boolResult = Copy-ObjectNotMarkedAsSerializable ([ref]$DestinationObject) ([ref]$SourceObject)  -  Optionally, a third positional argument may be specified as an integer between 1 and [int]::MaxValue that indicates the depth of the copy operation (i.e., how many levels deep to copy nested objects, recursively)'
            return $false
        }

        if ($null -eq ($args[1].Value)) {
            # If the input object is null, simply return null
            (($args[0]).Value) = $null
            return $true
        }

        if ($args.Count -eq 3) {
            if ($null -ne $args[2]) {
                if (($args[2].GetType().FullName -notlike 'System.Int*') -and ($args[2].GetType().FullName -notlike 'System.UInt*')) {
                    Write-Error 'Copy-ObjectNotMarkedAsSerializable must be called with at least two arguments. The first positional argument is required, representing a reference to the destination (new) object. The second positional argument is also required, representing a reference to the source object to be copied. For example: $boolResult = Copy-ObjectNotMarkedAsSerializable ([ref]$DestinationObject) ([ref]$SourceObject)  -  Optionally, a third positional argument may be specified as an integer between 1 and [int]::MaxValue that indicates the depth of the copy operation (i.e., how many levels deep to copy nested objects, recursively)'
                    return $false
                } else {
                    if ($args[2] -lt 1) {
                        Write-Error 'Copy-ObjectNotMarkedAsSerializable must be called with at least two arguments. The first positional argument is required, representing a reference to the destination (new) object. The second positional argument is also required, representing a reference to the source object to be copied. For example: $boolResult = Copy-ObjectNotMarkedAsSerializable ([ref]$DestinationObject) ([ref]$SourceObject)  -  Optionally, a third positional argument may be specified as an integer between 1 and [int]::MaxValue that indicates the depth of the copy operation (i.e., how many levels deep to copy nested objects, recursively)'
                        return $false
                    } else {
                        $intCopyDepth = $args[2]
                    }
                }
            } else {
                # Third argument was $null - default it
                $intCopyDepth = $intDefaultCopyDepth
            }
        } else {
            # Third argument was not supplied - default it
            $intCopyDepth = $intDefaultCopyDepth
        }

        $boolOperationCompleted = $false

        if ($boolOperationCompleted -ne $true) {
            # Try ConvertTo/ConvertFrom-Json first because it is much faster than
            # PSSerializer or Export/Import-Clixml

            $strJSONRepresenationOfSourceObject = ''
            $boolSuccess = ConvertTo-JsonSafely ([ref]$strJSONRepresenationOfSourceObject) ($args[1]) $intCopyDepth
            if ($boolSuccess -eq $true) {
                $boolSuccess = ConvertFrom-JsonSafely ($args[0]) ([ref]$strJSONRepresenationOfSourceObject)
                if ($boolSuccess -eq $true) {
                    $boolOperationCompleted = $true
                }
            }
        }

        if ($boolOperationCompleted -ne $true) {
            # If we get here, an error occurred while tring to convert to/from JSON

            # Safely check to see if System.Management.Automation.PSSerializer exists
            # If it does, we will want to use it for performance reasons
            $boolPSSerializerExists = Test-PSSerializerTypeAvailabilityWithoutEnumeratingDotNETAssemblies
            if ($boolPSSerializerExists -eq $true) {
                # System.Management.Automation.PSSerializer exists, which means we
                # can create an XML representation of the source object in memory
                $strCliXMLSourceObject = ''
                $boolSuccess = Invoke-PSSerializerSerializeOperation ([ref]$strCliXMLSourceObject) ($args[1]) $intCopyDepth
                if ($boolSuccess -eq $true) {
                    $boolSuccess = Invoke-PSSerializerDeserializeOperation ($args[0]) ([ref]$strCliXMLSourceObject)
                    if ($boolSuccess -eq $true) {
                        $boolOperationCompleted = $true
                    }
                }
            }
        }

        if ($boolOperationCompleted -ne $true) {
            # If we get here, one of two things happened:
            # 1. ConvertTo/ConvertFrom-Json failed and
            #    System.Management.Automation.PSSerializer does not exist, which
            #    means we need to try Export/Import-Clixml.
            # 2. ConvertTo/ConvertFrom-Json failed and
            #    System.Management.Automation.PSSerializer does exist, but an error
            #    occurred while using it, which means we need to try
            #    Export/Import-Clixml.

            # Either way, we will try Export/Import-Clixml next:

            $strTempFilePath = [System.IO.Path]::GetTempFileName()
            $boolSuccess = Export-ClixmlSafely $strTempFilePath ($args[1]) $intCopyDepth
            if ($boolSuccess -eq $true) {
                $boolSuccess = Import-ClixmlSafely ($args[0]) $strTempFilePath
                if ($boolSuccess -eq $true) {
                    if (Test-Path $strTempFilePath) {
                        Remove-Item $strTempFilePath -Force -ErrorAction SilentlyContinue
                    }
                    $boolOperationCompleted = $true
                }
            }
        }

        if ($boolOperationCompleted -ne $true) {
            if ($intCopyDepth -ne $intDefaultCopyDepth) {
                Write-Warning ('An error occurred running Copy-ObjectNotMarkedAsSerializable with copy depth ' + $intCopyDepth + '; trying again with copy depth ' + $intDefaultCopyDepth + '...')

                if ($boolOperationCompleted -ne $true) {
                    # Try ConvertTo/ConvertFrom-Json first because it is much faster than
                    # PSSerializer or Export/Import-Clixml

                    $strJSONRepresenationOfSourceObject = ''
                    $boolSuccess = ConvertTo-JsonSafely ([ref]$strJSONRepresenationOfSourceObject) ($args[1]) $intDefaultCopyDepth
                    if ($boolSuccess -eq $true) {
                        $boolSuccess = ConvertFrom-JsonSafely ($args[0]) ([ref]$strJSONRepresenationOfSourceObject)
                        if ($boolSuccess -eq $true) {
                            $boolOperationCompleted = $true
                        }
                    }
                }

                if ($boolOperationCompleted -ne $true) {
                    # If we get here, an error occurred while tring to convert to/from JSON

                    # Safely check to see if System.Management.Automation.PSSerializer exists
                    # If it does, we will want to use it for performance reasons
                    $boolPSSerializerExists = Test-PSSerializerTypeAvailabilityWithoutEnumeratingDotNETAssemblies
                    if ($boolPSSerializerExists -eq $true) {
                        # System.Management.Automation.PSSerializer exists, which means we
                        # can create an XML representation of the source object in memory
                        $strCliXMLSourceObject = ''
                        $boolSuccess = Invoke-PSSerializerSerializeOperation ([ref]$strCliXMLSourceObject) ($args[1]) $intDefaultCopyDepth
                        if ($boolSuccess -eq $true) {
                            $boolSuccess = Invoke-PSSerializerDeserializeOperation ($args[0]) ([ref]$strCliXMLSourceObject)
                            if ($boolSuccess -eq $true) {
                                $boolOperationCompleted = $true
                            }
                        }
                    }
                }

                if ($boolOperationCompleted -ne $true) {
                    # If we get here, one of two things happened:
                    # 1. ConvertTo/ConvertFrom-Json failed and
                    #    System.Management.Automation.PSSerializer does not exist, which
                    #    means we need to try Export/Import-Clixml.
                    # 2. ConvertTo/ConvertFrom-Json failed and
                    #    System.Management.Automation.PSSerializer does exist, but an error
                    #    occurred while using it, which means we need to try
                    #    Export/Import-Clixml.

                    # Either way, we will try Export/Import-Clixml next:

                    $strTempFilePath = [System.IO.Path]::GetTempFileName()
                    $boolSuccess = Export-ClixmlSafely $strTempFilePath ($args[1]) $intDefaultCopyDepth
                    if ($boolSuccess -eq $true) {
                        $boolSuccess = Import-ClixmlSafely ($args[0]) $strTempFilePath
                        if ($boolSuccess -eq $true) {
                            if (Test-Path $strTempFilePath) {
                                Remove-Item $strTempFilePath -Force -ErrorAction SilentlyContinue
                            }
                            $boolOperationCompleted = $true
                        }
                    }
                }
            }
        }

        if ($boolOperationCompleted -ne $true) {
            # Write-Error 'An error occurred while running Copy-ObjectNotMarkedAsSerializable. The object was not able to be copied.'
            return $false
        } else {
            return $true
        }
    }

    function New-MemoryStream {
        #region FunctionHeader #####################################################
        # This function creates a System.IO.MemoryStream object
        #
        # One positional arguments is required: a reference to an object that will
        # become the MemoryStream object.
        #
        # This function uses the following arguments:
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        # $objMemoryStream = $null
        # $boolSuccess = New-MemoryStream ([ref]$objMemoryStream)
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        $refOutput = $args[0]

        # TODO: Validate input

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Create the MemoryStream object
        ($refOutput.Value) = New-Object -TypeName 'System.IO.MemoryStream'

        # Restore the former error preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            # Error occurred
            return $false
        } else {
            return $true
        }
    }

    function New-BinaryFormatter {
        #region FunctionHeader #####################################################
        # This function creates a
        # System.Runtime.Serialization.Formatters.Binary.BinaryFormatter object
        #
        # One positional arguments is required: a reference to an object that will
        # become the BinaryFormatter object.
        #
        # This function uses the following arguments:
        #
        # The function returns $true if the process completed successfully; $false
        # otherwise
        #
        # Example usage:
        # $objBinaryFormatter = $null
        # $boolSuccess = New-BinaryFormatter ([ref]$objBinaryFormatter)
        #
        # Version: 1.0.20240127.0
        #endregion FunctionHeader #####################################################

        #region License ############################################################
        # Copyright (c) 2024 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a copy
        # of this software and associated documentation files (the "Software"), to deal
        # in the Software without restriction, including without limitation the rights
        # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        # copies of the Software, and to permit persons to whom the Software is
        # furnished to do so, subject to the following conditions:
        #
        # The above copyright notice and this permission notice shall be included in
        # all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ############################################################

        #region DownloadLocationNotice #############################################
        # The most up-to-date version of this script can be found on the author's
        # GitHub repository at https://github.com/franklesniak/Copy-Object
        #endregion DownloadLocationNotice #############################################

        trap {
            # Intentionally left empty to prevent terminating errors from halting
            # processing
        }

        $refOutput = $args[0]

        # TODO: Validate input

        # Retrieve the newest error on the stack prior to doing work
        $refLastKnownError = Get-ReferenceToLastError

        # Store current error preference; we will restore it after we do the work of
        # this function
        $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

        # Set ErrorActionPreference to SilentlyContinue; this will suppress error
        # output. Terminating errors will not output anything, kick to the empty trap
        # statement and then continue on. Likewise, non-terminating errors will also
        # not output anything, but they do not kick to the trap statement; they simply
        # continue on.
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

        # Create the BinaryFormatter object
        ($refOutput.Value) = New-Object -TypeName 'System.Runtime.Serialization.Formatters.Binary.BinaryFormatter'

        # Restore the former error preference
        $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

        # Retrieve the newest error on the error stack
        $refNewestCurrentError = Get-ReferenceToLastError

        if (Test-ErrorOccurred $refLastKnownError $refNewestCurrentError) {
            # Error occurred
            return $false
        } else {
            return $true
        }
    }
    #endregion FunctionsForCopyingAnObject ############################################

    # Anything greater is too slow to use as a default setting; see notes above
    $intDefaultCopyDepth = 2

    if (($args.Count -lt 2) -or ($args.Count -gt 4)) {
        Write-Error 'Copy-Object was called with the incorrect number of arguments. The first positional argument is required, representing a reference to the destination object. The second positional argument is also required, representing a reference to the source object to be copied. Optionally, a third positional argument may be specified as an integer between 1 and [int]::MaxValue that indicates the depth of the copy operation (i.e., how many levels deep to copy nested objects, recursively). Finally, optionally, a fourth positional argument may be specified as a boolean indicating whether the source object should be considered "safe", i.e., generated from a trusted process and not possible to contain malicious code.'
        return 2 # Error
    }

    if (($args[0].GetType().FullName -notlike 'System.Management.Automation.PSReference*') -or ($args[1].GetType().FullName -notlike 'System.Management.Automation.PSReference*')) {
        Write-Error 'Copy-Object must be called with at least two arguments: The first positional argument is required, representing a reference to the destination object. The second positional argument is also required, representing a reference to the source object to be copied. For example: $intReturnCode = Copy-Object ([ref]$DestinationObject) ([ref]$SourceObject)  -  Optionally, a third positional argument may be specified as an integer between 1 and [int]::MaxValue that indicates the depth of the copy operation (i.e., how many levels deep to copy nested objects, recursively). Finally, optionally, a fourth positional argument may be specified as a boolean indicating whether the source object should be considered "safe", i.e., generated from a trusted process and not possible to contain malicious code.'
        return 2 # Error
    }

    if ($null -eq ($args[1].Value)) {
        # If the input object is null, simply return null
        (($args[0]).Value) = $null
        return 0 # Success
    }

    if ($args.Count -ge 3) {
        if ($null -ne $args[2]) {
            if (($args[2].GetType().FullName -notlike 'System.Int*') -and ($args[2].GetType().FullName -notlike 'System.UInt*')) {
                Write-Error 'Copy-Object must be called with at least two arguments: The first positional argument is required, representing a reference to the destination object. The second positional argument is also required, representing a reference to the source object to be copied. For example: $intReturnCode = Copy-Object ([ref]$DestinationObject) ([ref]$SourceObject)  -  Optionally, a third positional argument may be specified as an integer between 1 and [int]::MaxValue that indicates the depth of the copy operation (i.e., how many levels deep to copy nested objects, recursively). Finally, optionally, a fourth positional argument may be specified as a boolean indicating whether the source object should be considered "safe", i.e., generated from a trusted process and not possible to contain malicious code.'
                return 2 # Error
            } else {
                if ($args[2] -lt 1) {
                    Write-Error 'Copy-Object must be called with at least two arguments: The first positional argument is required, representing a reference to the destination object. The second positional argument is also required, representing a reference to the source object to be copied. For example: $intReturnCode = Copy-Object ([ref]$DestinationObject) ([ref]$SourceObject)  -  Optionally, a third positional argument may be specified as an integer between 1 and [int]::MaxValue that indicates the depth of the copy operation (i.e., how many levels deep to copy nested objects, recursively). Finally, optionally, a fourth positional argument may be specified as a boolean indicating whether the source object should be considered "safe", i.e., generated from a trusted process and not possible to contain malicious code.'
                    return 2 # Error
                } else {
                    $intCopyDepth = $args[2]
                }
            }
        } else {
            # Third argument was $null - default it
            $intCopyDepth = $intDefaultCopyDepth
        }
    } else {
        # Third argument was not supplied - default it
        $intCopyDepth = $intDefaultCopyDepth
    }

    if ($args.Count -ge 4) {
        if ($null -ne $args[3]) {
            if ($args[3].GetType().FullName -ne 'System.Boolean') {
                Write-Error 'Copy-Object must be called with at least two arguments: The first positional argument is required, representing a reference to the destination object. The second positional argument is also required, representing a reference to the source object to be copied. For example: $intReturnCode = Copy-Object ([ref]$DestinationObject) ([ref]$SourceObject)  -  Optionally, a third positional argument may be specified as an integer between 1 and [int]::MaxValue that indicates the depth of the copy operation (i.e., how many levels deep to copy nested objects, recursively). Finally, optionally, a fourth positional argument may be specified as a boolean indicating whether the source object should be considered "safe", i.e., generated from a trusted process and not possible to contain malicious code.'
                return 2 # Error
            } else {
                $boolSourceObjectIsSafe = $args[3]
            }
        } else {
            # Fourth argument was $null - default it
            $boolSourceObjectIsSafe = $false
        }
    } else {
        # Fourth argument was not supplied - default it
        $boolSourceObjectIsSafe = $false
    }

    # If we are still here, we are dealing with a reference to a source object in
    # $args[1], we have a copy depth set in $intCopyDepth, and we have whether the
    # source object is marked as safe in $boolSourceObjectIsSafe

    if ($boolSourceObjectIsSafe -eq $true) {
        # If the source object is marked as safe, we can use BinaryFormatter to copy it
        # if the object is marked as serializable. This is the most reliable way to
        # copy an object, but it is also the most dangerous because BinaryFormatter
        # has inherent security vulnerabilities. See:
        # https://learn.microsoft.com/en-us/dotnet/standard/serialization/binaryformatter-security-guide

        # Check to see if the object is marked as serializable:
        if (($args[1].Value).GetType().FullName -match ([regex]::Escape('[]'))) {
            # Input object is an array. Check serializability of first object within the array
            if (($args[1].Value).Count -eq 0) {
                # Input object was an empty array.
                $boolSerializableObject = $true
            } else {
                $boolNonNullItemInArray = $false
                for ($intArrayCounter = 0; $intArrayCounter -lt ($args[1].Value).Count; $intArrayCounter++) {
                    if ($null -ne (($args[1].Value)[$intArrayCounter])) {
                        $boolNonNullItemInArray = $true
                        break # exit the for loop
                    }
                }

                if ($boolNonNullItemInArray) {
                    # check it out
                    if ((($args[1].Value)[$intArrayCounter]).GetType().IsSerializable -eq $true) {
                        $boolSerializableObject = $true
                    } else {
                        $boolSerializableObject = $false
                    }
                } else {
                    # an array full of $nulls is serializable
                    $boolSerializableObject = $true
                }
            }
        } else {
            # Source object is not an array
            if (($args[1].Value).GetType().IsSerializable -eq $true) {
                $boolSerializableObject = $true
            } else {
                $boolSerializableObject = $false
            }
        }
    } else {
        $boolSerializableObject = $false
    }

    $boolSerializedObjectCopyUsed = $true

    if ($boolSerializableObject) {
        # Object appears serializable and is marked as safe
        $memoryStream = $null
        $boolSuccess = New-MemoryStream ([ref]$memoryStream)
        if ($boolSuccess -ne $true) {
            # Error occurred; fall back to non-serializable methods
            $boolSerializedObjectCopyUsed = $false
            $boolResult = Copy-ObjectNotMarkedAsSerializable ($args[0]) ($args[1]) $intCopyDepth
            if ($boolResult -ne $true) {
                # Write-Error 'Because an error occurred in Copy-ObjectNotMarkedAsSerializable, the Copy-Object operation cannot continue.'
                return 2
            }
        } else {
            # No error occurred; keep going
            $binaryFormatter = $null
            $boolSuccess = New-BinaryFormatter ([ref]$binaryFormatter)
            if ($boolSuccess -ne $true) {
                # Error occurred; fall back to non-serializable methods
                $boolSerializedObjectCopyUsed = $false
                $boolResult = Copy-ObjectNotMarkedAsSerializable ($args[0]) ($args[1]) $intCopyDepth
                if ($boolResult -ne $true) {
                    # Write-Error 'Because an error occurred in Copy-ObjectNotMarkedAsSerializable, the Copy-Object operation cannot continue.'
                    return 2
                }
            } else {
                # No error occurred; keep going
                $boolSuccess = Invoke-BinaryFormatterSerializeOperation ([ref]$memoryStream) ([ref]$binaryFormatter) ($args[1])
                if ($boolSuccess -ne $true) {
                    # Error occurred; fall back to non-serializable methods
                    $boolSerializedObjectCopyUsed = $false
                    $boolResult = Copy-ObjectNotMarkedAsSerializable ($args[0]) ($args[1]) $intCopyDepth
                    if ($boolResult -ne $true) {
                        # Write-Error 'Because an error occurred in Copy-ObjectNotMarkedAsSerializable, the Copy-Object operation cannot continue.'
                        return 2
                    }
                } else {
                    # No error occurred; keep going
                    $memoryStream.Position = 0
                    $boolSuccess = Invoke-BinaryFormatterDeserializeOperation ($args[0]) ([ref]$binaryFormatter) ([ref]$memoryStream)
                    if ($boolSuccess -ne $true) {
                        # Error occurred; fall back to non-serializable methods
                        $boolSerializedObjectCopyUsed = $false
                        $boolResult = Copy-ObjectNotMarkedAsSerializable ($args[0]) ($args[1]) $intCopyDepth
                        if ($boolResult -ne $true) {
                            # Write-Error 'Because an error occurred in Copy-ObjectNotMarkedAsSerializable, the Copy-Object operation cannot continue.'
                            return 2
                        }
                    } else {
                        # No error occurred; keep going
                        $memoryStream.Close()
                    }
                }
            }
        }
    } else {
        # Object is not marked as serializable or not indicated as safe
        $boolSerializedObjectCopyUsed = $false
        $boolResult = Copy-ObjectNotMarkedAsSerializable ($args[0]) ($args[1]) $intCopyDepth
        if ($boolResult -ne $true) {
            # Write-Error 'Because an error occurred in Copy-ObjectNotMarkedAsSerializable, the Copy-Object operation cannot continue.'
            return 2
        }
    }

    if ($boolSerializedObjectCopyUsed) {
        return 0
    } else {
        return 1
    }
}

$versionPS = Get-PSVersion

#region Quit if PowerShell version is unsupported by Az Module #####################
if ($versionPS -lt [version]'5.1') {
    Write-Warning 'This script requires PowerShell v5.1 or higher. Please upgrade to PowerShell v5.1 or higher and try again.'
    return # Quit script
}
#endregion Quit if PowerShell version is unsupported by Az Module #####################

# Make sure the input files exist
if ((Test-Path -Path $ClusterMetadataInputCSVPath -PathType Leaf) -eq $false) {
    Write-Warning ('Cluster metadata CSV file not found at: "' + $ClusterMetadataInputCSVPath + '"')
    return # Quit script
}
if ((Test-Path -Path $UnstructuredTextDataInputCSVPath -PathType Leaf) -eq $false) {
    Write-Warning ('Unstructured text data CSV file not found at: "' + $UnstructuredTextDataInputCSVPath + '"')
    return # Quit script
}

#region Check for required PowerShell Modules ######################################
$hashtableModuleNameToInstalledModules = @{}
$hashtableModuleNameToInstalledModules.Add('Az.Accounts', @())
$hashtableModuleNameToInstalledModules.Add('Az.KeyVault', @())
$hashtableModuleNameToInstalledModules.Add('Microsoft.PowerShell.SecretManagement', @())
$hashtableModuleNameToInstalledModules.Add('Microsoft.PowerShell.SecretStore', @())
$refHashtableModuleNameToInstalledModules = [ref]$hashtableModuleNameToInstalledModules
Get-PowerShellModuleUsingHashtable -ReferenceToHashtable $refHashtableModuleNameToInstalledModules

$hashtableCustomNotInstalledMessageToModuleNames = @{}

$strAzNotInstalledMessage = 'Az.Accounts and/or Az.KeyVault modules were not found. Please install the full Az module and then try again.' + [System.Environment]::NewLine + 'You can install the Az PowerShell module from the PowerShell Gallery by running the following command:' + [System.Environment]::NewLine + 'Install-Module Az;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
$hashtableCustomNotInstalledMessageToModuleNames.Add($strAzNotInstalledMessage, @('Az.Accounts', 'Az.KeyVault'))

$refhashtableCustomNotInstalledMessageToModuleNames = [ref]$hashtableCustomNotInstalledMessageToModuleNames
$boolResult = Test-PowerShellModuleInstalledUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ThrowErrorIfModuleNotInstalled -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames

if ($boolResult -eq $false) {
    return # Quit script
}
#endregion Check for required PowerShell Modules ######################################

#region Check for PowerShell module updates ########################################
if ($DoNotCheckForModuleUpdates.IsPresent -eq $false) {
    Write-Verbose 'Checking for module updates...'
    $hashtableCustomNotUpToDateMessageToModuleNames = @{}

    $strAzNotUpToDateMessage = 'A newer version of the Az.Accounts and/or Az.KeyVault modules was found. Please consider updating it by running the following command:' + [System.Environment]::NewLine + 'Install-Module Az -Force;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    $hashtableCustomNotUpToDateMessageToModuleNames.Add($strAzNotUpToDateMessage, @('Az.Accounts', 'Az.KeyVault'))

    $refhashtableCustomNotUpToDateMessageToModuleNames = [ref]$hashtableCustomNotUpToDateMessageToModuleNames
    $boolResult = Test-PowerShellModuleUpdatesAvailableUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ThrowErrorIfModuleNotInstalled -ThrowWarningIfModuleNotUpToDate -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames -ReferenceToHashtableOfCustomNotUpToDateMessages $refhashtableCustomNotUpToDateMessageToModuleNames
}
#endregion Check for PowerShell module updates ########################################

$PSAzureContext = Get-AzContext
if ($null -eq $PSAzureContext) {
    # Connect to Azure without caching credentials to disk:
    [void](Connect-AzAccount -Tenant $EntraIdTenantId -Subscription $AzureSubscriptionId -Scope Process)
    $PSAzureContext = Get-AzContext
    if ($null -eq $PSAzureContext) {
        Write-Warning 'No Azure context found. Please connect to Azure and try again.'
        return # Quit script
    }
}

$arrSecretVaults = @(@(Get-SecretVault) | Where-Object { $_.Name -eq ($AzureKeyVaultName + '-AKV') })
if ($arrSecretVaults.Count -eq 0) {
    # Secret vault is not registered

    #TODO: store the connection results in a variable, and make sure the connection was successful?

    # Set the Azure Key Vault as the default secret store
    $parameters = @{
        Name = ($AzureKeyVaultName + '-AKV')
        ModuleName = 'Az.KeyVault'
        VaultParameters = @{
            AZKVaultName = $AzureKeyVaultName
            SubscriptionId = $AzureSubscriptionId
        }
        DefaultVault = $true
    }

    # Register the Azure Key Vault in the SecretManagement module
    Register-SecretVault @parameters
}

# Get the secret stored in AKV
$apiKey = Get-Secret -Name $SecretName -Vault ($AzureKeyVaultName + '-AKV')

# Convert the secure string to a plain text string
$strAPIKey = [System.Net.NetworkCredential]::new("", $apiKey).Password

if ([string]::IsNullOrEmpty($strAPIKey)) {
    Write-Warning 'Unable to retrieve key from Azure Key Vault. Please add the API Key to the Azure Key Vault or retry the connection.'
    return # Quit script
}

# Make sure the output file doesn't already exist and if it does, delete it and then
# verify that it's gone
if ((Test-Path -Path $OutputCSVPath -PathType Leaf) -eq $true) {
    Remove-Item -Path $OutputCSVPath -Force
    if ((Test-Path -Path $OutputCSVPath -PathType Leaf) -eq $true) {
        Write-Warning ('Output CSV file already exists and could not be deleted (the file may be in use): "' + $OutputCSVPath + '"')
        return
    }
}

# Import the cluster metadata CSV
$arrClusterMetadataCSV = @()
$arrClusterMetadataCSV = @(Import-Csv $ClusterMetadataInputCSVPath)
if ($arrClusterMetadataCSV.Count -eq 0) {
    Write-Warning ('Input CSV file is empty: "' + $ClusterMetadataInputCSVPath + '"')
    return
}

# Import the unstructured text data CSV
$arrUnstructuredTextDataCSV = @()
$arrUnstructuredTextDataCSV = @(Import-Csv $UnstructuredTextDataInputCSVPath)
if ($arrUnstructuredTextDataCSV.Count -eq 0) {
    Write-Warning ('Input CSV file is empty: "' + $UnstructuredTextDataInputCSVPath + '"')
    return
}

# Create the list to store output
if ($versionPS -ge ([version]'6.0')) {
    $listPSCustomObjectOutput = New-Object -TypeName 'System.Collections.Generic.List[PSCustomObject]'
} else {
    # On Windows PowerShell (versions older than 6.x), we use an ArrayList instead
    # of a generic list
    # TODO: Fill in rationale for this
    #
    # Technically, in older versions of PowerShell, the type in the ArrayList will
    # be a PSObject; but that does not matter for our purposes.
    $listPSCustomObjectOutput = New-Object -TypeName 'System.Collections.ArrayList'
}

#region Collect Stats/Objects Needed for Writing Progress ##########################
$intProgressReportingFrequency = 3
$intTotalItems = $arrClusterMetadataCSV.Count
$strProgressActivity = 'Getting the topic for each cluster'
$strProgressStatus = 'Processing'
$strProgressCurrentOperationPrefix = 'Processing item'
$timedateStartOfLoop = Get-Date
# Create a queue for storing lagging timestamps for ETA calculation
$queueLaggingTimestamps = New-Object System.Collections.Queue
$queueLaggingTimestamps.Enqueue($timedateStartOfLoop)
#endregion Collect Stats/Objects Needed for Writing Progress ##########################

Write-Verbose ($strProgressStatus + '...')

for ($intRowIndex = 0; $intRowIndex -lt $intTotalItems; $intRowIndex++) {
    #region Report Progress ########################################################
    $intCounterLoop = $intRowIndex
    $intCurrentItemNumber = $intCounterLoop + 1 # Forward direction for loop
    if ((($intCurrentItemNumber -ge ($intProgressReportingFrequency * 3)) -and ($intCurrentItemNumber % $intProgressReportingFrequency -eq 0)) -or ($intCurrentItemNumber -eq $intTotalItems)) {
        # Create a progress bar after the first (3 x $intProgressReportingFrequency) items have been processed
        $timeDateLagging = $queueLaggingTimestamps.Dequeue()
        $datetimeNow = Get-Date
        $timespanTimeDelta = $datetimeNow - $timeDateLagging
        $intNumberOfItemsProcessedInTimespan = $intProgressReportingFrequency * ($queueLaggingTimestamps.Count + 1)
        $doublePercentageComplete = ($intCurrentItemNumber - 1) / $intTotalItems
        $intItemsRemaining = $intTotalItems - $intCurrentItemNumber + 1
        Write-Progress -Activity $strProgressActivity -Status $strProgressStatus -PercentComplete ($doublePercentageComplete * 100) -CurrentOperation ($strProgressCurrentOperationPrefix + ' ' + $intCurrentItemNumber + ' of ' + $intTotalItems + ' (' + [string]::Format('{0:0.00}', ($doublePercentageComplete * 100)) + '%)') -SecondsRemaining (($timespanTimeDelta.TotalSeconds / $intNumberOfItemsProcessedInTimespan) * $intItemsRemaining)
    }
    #endregion Report Progress ########################################################

    # Create a copy of the source object and add the new data to it
    $psobjectUpdated = $null
    $boolResult = Copy-Object ([ref]$psobjectUpdated) ([ref]($arrClusterMetadataCSV[$intRowIndex]))

    if ($UseMostRepresentativeItemForTopicExtraction -eq $true) {
        $refIndexOfMostRepresentativeItemInCluster = [ref]((($arrClusterMetadataCSV[$intRowIndex]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $ClusterMetadataFieldNameIndexToMostRepresentativeItem }).Value)
        $refUnstructuredTextData = [ref]((($arrUnstructuredTextDataCSV[$refIndexOfMostRepresentativeItemInCluster.Value]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $UnstructuredTextDataFieldNameContainingTextData }).Value)
        # Write-Host $refUnstructuredTextData.Value
        $strTopicOfMostRepresentativeItem = Get-TopicFromDataSetUsingAzureOpenAI -AzureOpenAIEndpoint $AzureOpenAIEndpoint -AzureOpenAIDeploymentName $AzureOpenAIDeploymentName -AzureOpenAIAPIKey $strAPIKey -AzureOpenAIAPIVersion $AzureOpenAIAPIVersion -ReferenceToArrayOfUnstructuredTextData ([ref]@($refUnstructuredTextData.Value))

        $psobjectUpdated | Add-Member -MemberType NoteProperty -Name $OutputDataFieldNameForTopicFromMostRepresentativeItem -Value $strTopicOfMostRepresentativeItem
    }
    if ($OutputDataFromMostRepresentativeItem -eq $true) {
        $psobjectUpdated | Add-Member -MemberType NoteProperty -Name $OutputDataFieldNameForMostRepresentativeItem -Value ($refUnstructuredTextData.Value)
    }
    $arrTexts = @()
    if ($UseNMostRepresentativeItemsInClusterForTopicExtraction -eq $true) {
        $refNumberOfNMostRepresentativeDataPointsInCluster = [ref]((($arrClusterMetadataCSV[$intRowIndex]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $ClusterMetadataFieldNameNumberOfNMostRepresentativeDataPoints }).Value)
        $refIndicesOfNMostRepresentativeItemsInCluster = [ref]((($arrClusterMetadataCSV[$intRowIndex]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $ClusterMetadataFieldNameIndiciesOfNMostRepresentativeItems }).Value)

        $intNumberOfNMostRepresentativeDataPointsInCluster = [int]($refNumberOfNMostRepresentativeDataPointsInCluster.Value)
        $arrIndicesOfNMostRepresentativeItemsInCluster = Split-StringOnLiteralString ($refIndicesOfNMostRepresentativeItemsInCluster.Value) $ClusterMetadataIndicesSeparator

        for ($intCounterB = 0; $intCounterB -lt $intNumberOfNMostRepresentativeDataPointsInCluster; $intCounterB++) {
            $intIndex = [int]($arrIndicesOfNMostRepresentativeItemsInCluster[$intCounterB])
            $arrTexts += (($arrUnstructuredTextDataCSV[$intIndex]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $UnstructuredTextDataFieldNameContainingTextData }).Value
        }

        $strTopicOfNMostRepresentativeItems = Get-TopicFromDataSetUsingAzureOpenAI -AzureOpenAIEndpoint $AzureOpenAIEndpoint -AzureOpenAIDeploymentName $AzureOpenAIDeploymentName -AzureOpenAIAPIKey $strAPIKey -AzureOpenAIAPIVersion $AzureOpenAIAPIVersion -ReferenceToArrayOfUnstructuredTextData ([ref]$arrTexts)

        $psobjectUpdated | Add-Member -MemberType NoteProperty -Name $OutputDataFieldNameForTopicFromNMostRepresentativeItems -Value $strTopicOfNMostRepresentativeItems
    }
    if ($OutputDataFromNMostRepresentativeItems -eq $true) {
        if ($arrTexts.Count -eq 0) {
            for ($intCounterB = 0; $intCounterB -lt $intNumberOfNMostRepresentativeDataPointsInCluster; $intCounterB++) {
                $intIndex = [int]($arrIndicesOfNMostRepresentativeItemsInCluster[$intCounterB])
                $arrTexts += (($arrUnstructuredTextDataCSV[$intIndex]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $UnstructuredTextDataFieldNameContainingTextData }).Value
            }
        }
        $boolSeparatorConflict = $false
        for ($intCounterB = 0; $intCounterB -lt $arrTexts.Count; $intCounterB++) {
            if (($arrTexts[$intCounterB]).Contains($SeparatorForItemsInCluster) -eq $true) {
                $boolSeparatorConflict = $true
                break
            }
        }
        if ($boolSeparatorConflict -eq $true) {
            Write-Warning 'The separator for items in a cluster conflicts with the text data. Please choose a different separator.'
            return # Quit script
        }
        $strTexts = $arrTexts -join $SeparatorForItemsInCluster
        $psobjectUpdated | Add-Member -MemberType NoteProperty -Name $OutputDataFieldNameForNMostRepresentativeItems -Value $strTexts
    }
    $arrTexts = @()
    if ($UseAllItemsInClusterForTopicExtraction -eq $true) {
        $refCountOfItemsInCluster = [ref]((($arrClusterMetadataCSV[$intRowIndex]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $ClusterMetadataFieldNameCountOfItemsInCluster }).Value)
        $refIndicesOfAllItemsInCluster = [ref]((($arrClusterMetadataCSV[$intRowIndex]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $ClusterMetadataFieldNameIndiciesOfAllItemsInCluster }).Value)

        $intCountOfItemsInCluster = [int]($refCountOfItemsInCluster.Value)
        $arrIndicesOfAllItemsInCluster = Split-StringOnLiteralString ($refIndicesOfAllItemsInCluster.Value) $ClusterMetadataIndicesSeparator

        for ($intCounterB = 0; $intCounterB -lt $intCountOfItemsInCluster; $intCounterB++) {
            $intIndex = [int]($arrIndicesOfAllItemsInCluster[$intCounterB])
            $arrTexts += (($arrUnstructuredTextDataCSV[$intIndex]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $UnstructuredTextDataFieldNameContainingTextData }).Value
        }

        $strTopicOfAllItems = Get-TopicFromDataSetUsingAzureOpenAI -OpenAIAPIKey $strAPIKey -ReferenceToArrayOfUnstructuredTextData ([ref]$arrTexts)

        $psobjectUpdated | Add-Member -MemberType NoteProperty -Name $OutputDataFieldNameForTopicFromAllItemsInCluster -Value $strTopicOfAllItems
    }
    if ($OutputDataFromAllItemsInCluster -eq $true) {
        if ($arrTexts.Count -eq 0) {
            for ($intCounterB = 0; $intCounterB -lt $intCountOfItemsInCluster; $intCounterB++) {
                $intIndex = [int]($arrIndicesOfAllItemsInCluster[$intCounterB])
                $arrTexts += (($arrUnstructuredTextDataCSV[$intIndex]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $UnstructuredTextDataFieldNameContainingTextData }).Value
            }
        }
        $boolSeparatorConflict = $false
        for ($intCounterB = 0; $intCounterB -lt $arrTexts.Count; $intCounterB++) {
            if (($arrTexts[$intCounterB]).Contains($SeparatorForItemsInCluster) -eq $true) {
                $boolSeparatorConflict = $true
                break
            }
        }
        if ($boolSeparatorConflict -eq $true) {
            Write-Warning 'The separator for items in a cluster conflicts with the text data. Please choose a different separator.'
            return # Quit script
        }
        $strTexts = $arrTexts -join $SeparatorForItemsInCluster
        $psobjectUpdated | Add-Member -MemberType NoteProperty -Name $OutputDataFieldNameForAllItemsInCluster -Value $strTexts
    }

    # Add the updated object to the output list
    if ($versionPS -ge ([version]'6.0')) {
        $listPSCustomObjectOutput.Add($psobjectUpdated)
    } else {
        [void]($listPSCustomObjectOutput.Add($psobjectUpdated))
    }

    #region Post-Loop Progress Reporting ###########################################
    if ($intCurrentItemNumber -eq $intTotalItems) {
        Write-Progress -Activity $strProgressActivity -Status $strProgressStatus -Completed
    }
    if ($intCounterLoop % $intProgressReportingFrequency -eq 0) {
        # Add lagging timestamp to queue
        $queueLaggingTimestamps.Enqueue((Get-Date))
    }
    # Increment counter
    $intCounterLoop++
    #endregion Post-Loop Progress Reporting ###########################################
}

# Export the CSV
$listPSCustomObjectOutput | Export-Csv -Path $OutputCSVPath -NoTypeInformation
