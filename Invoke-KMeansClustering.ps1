# Invoke-KMeansClustering.ps1
# Version: 2.0.20240422.0

<#
.SYNOPSIS
Inputs a CSV containing "embeddings" and uses the K-Means clustering algorithm to group
the data into clusters. The script then exports the cluster metadata to a new CSV file.

.DESCRIPTION
This script reads in a CSV file containing "embeddings" (i.e., numerical
representations of data) and uses the K-Means clustering algorithm to group the data
into clusters. The script then creates a new CSV file where each row represents a
cluster. Each row contains a number representing the index of the item in the original
CSV closest to the centroid of the cluster, a number representing 'n' for the 'n'-most
representative items in the cluster (i.e., those closest to the centroid), a semicolon-
separated list of the indicies of the items in the original CSV that are the 'n'-most
representative items in the cluster, the number of total items in the cluster, and a
semicolon-separated list of the indices of the items in the original CSV file that
belong to the cluster.

.PARAMETER InputCSVPath
Specifies the path to the input CSV file containing the embeddings to be clustered.

.PARAMETER DataFieldNameContainingEmbeddings
Specifies the name of the field in the input CSV file containing the embeddings. The
embeddings must be stored as a semicolon-separated string of numbers.

.PARAMETER OutputCSVPath
Specifies the path to the output CSV file that will list the K-means cluster metadata.

.PARAMETER NSizeForMostRepresentativeDataPoints
Specifies the number of data points to include in the output CSV file that are closest
to the center of each cluster. The default value is 5.

.PARAMETER NumberOfClusters
Optional parameter; allows the user to specify the number of clusters they want the
script to use. When this parameter is not specified, the script will use the "elbow
method" to determine the optimal number of clusters.

.EXAMPLE
PS C:\> .\Invoke-KMeansClustering.ps1 -InputCSVPath 'C:\Users\jdoe\Documents\West Monroe Pulse Survey Comments Aug 2021 - With Embeddings.csv' -DataFieldNameContainingEmbeddings 'Embeddings' -OutputCSVPath 'C:\Users\jdoe\Documents\West Monroe Pulse Survey Comments Aug 2021 - Cluster Metadata.csv'

This example reads in a CSV file containing embeddings and uses the K-Means
clustering algorithm to group the data into clusters. The script then creates a new CSV
file containing the cluster metadata.

.OUTPUTS
None
#>

#region License ################################################################
# Copyright (c) 2024 Frank Lesniak and Daniel Stutz
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
    [Parameter(Mandatory = $true)][string]$InputCSVPath,
    [Parameter(Mandatory = $false)][string]$DataFieldNameContainingEmbeddings = 'Embeddings',
    [Parameter(Mandatory = $true)][string]$OutputCSVPath,
    [Parameter(Mandatory = $false)][string]$OutputCSVPathForAllClusterStatistics,
    [Parameter(Mandatory = $false)][string]$OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters,
    [Parameter(Mandatory = $false)][int]$NSizeForMostRepresentativeDataPoints = 5,
    [Parameter(Mandatory = $false)][ValidateScript({ ($_ -ge 2) -or ($_ -eq 0) })][int]$NumberOfClusters,
    [Parameter(Mandatory = $false)][switch]$DoNotCalculateExtendedStatistics,
    [Parameter(Mandatory = $false)][switch]$DoNotCheckForModuleUpdates,
    [Parameter(Mandatory = $false)][double]$WeightingFactorForClusterSize = 8,
    [Parameter(Mandatory = $false)][double]$WeightingFactorForWCSS = 3,
    [Parameter(Mandatory = $false)][double]$WeightingFactorForWCSSSecondDerivative = 42,
    [Parameter(Mandatory = $false)][double]$WeightingFactorForSihouetteScore = 23,
    [Parameter(Mandatory = $false)][double]$WeightingFactorForDaviesBouldinScore = 12,
    [Parameter(Mandatory = $false)][double]$WeightingFactorForCalinskiHarabaszScore = 12
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

function Test-NuGetDotOrgRegisteredAsPackageSource {
    <#
    .SYNOPSIS
    Tests to see nuget.org is registered as a package source. If it is not, the
    function can optionally throw an error or warning

    .DESCRIPTION
    The Test-NuGetDotOrgRegisteredAsPackageSource function tests to see if nuget.org is
    registered as a package source. If it is not, the function can optionally throw an
    error or warning that gives the user instructions to register nuget.org as a
    package source.

    .PARAMETER ThrowErrorIfNuGetDotOrgNotRegistered
    Is a switch parameter. If this parameter is specified, an error is thrown to tell
    the user that nuget.org is not registered as a package source, and the user is
    given instructions on how to register it. If this parameter is not specified, no
    error is thrown.

    .PARAMETER ThrowWarningIfNuGetDotOrgNotRegistered
    Is a switch parameter. If this parameter is specified, a warning is thrown to tell
    the user that nuget.org is not registered as a package source, and the user is
    given instructions on how to register it. If this parameter is not specified, or if
    the ThrowErrorIfNuGetDotOrgNotRegistered parameter was specified, no warning is
    thrown.

    .EXAMPLE
    $boolResult = Test-NuGetDotOrgRegisteredAsPackageSource -ThrowErrorIfNuGetDotOrgNotRegistered

    This example checks to see if nuget.org is registered as a package source. If it is
    not, an error is thrown to tell the user that nuget.org is not registered as a
    package source, and the user is given instructions on how to register it.

    .OUTPUTS
    [boolean] - Returns $true if nuget.org is registered as a package source; otherwise, returns $false.

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
        [Parameter(Mandatory = $false)][switch]$ThrowErrorIfNuGetDotOrgNotRegistered,
        [Parameter(Mandatory = $false)][switch]$ThrowWarningIfNuGetDotOrgNotRegistered
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
        Write-Warning 'Test-NuGetDotOrgRegisteredAsPackageSource requires PowerShell version 5.0 or newer.'
        return
    }

    $WarningPreferenceAtStartOfFunction = $WarningPreference
    $VerbosePreferenceAtStartOfFunction = $VerbosePreference
    $DebugPreferenceAtStartOfFunction = $DebugPreference

    $boolThrowErrorForMissingPackageSource = $false
    $boolThrowWarningForMissingPackageSource = $false

    if ($ThrowErrorIfNuGetDotOrgNotRegistered.IsPresent -eq $true) {
        $boolThrowErrorForMissingPackageSource = $true
    } elseif ($ThrowWarningIfNuGetDotOrgNotRegistered.IsPresent -eq $true) {
        $boolThrowWarningForMissingPackageSource = $true
    }

    $boolPackageSourceFound = $true
    Write-Information ('Checking for registered package sources (Get-PackageSource)...')
    $WarningPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $arrPackageSources = @(Get-PackageSource)
    $WarningPreference = $WarningPreferenceAtStartOfFunction
    $VerbosePreference = $VerbosePreferenceAtStartOfFunction
    $DebugPreference = $DebugPreferenceAtStartOfFunction
    if (@($arrPackageSources | Where-Object { $_.Location -eq 'https://api.nuget.org/v3/index.json' }).Count -eq 0) {
        $boolPackageSourceFound = $false
    }

    if ($boolPackageSourceFound -eq $false) {
        $strMessage = 'The nuget.org package source is not registered. Please register it and then try again.' + [System.Environment]::NewLine + 'You can register it by running the following command: ' + [System.Environment]::NewLine + '[void](Register-PackageSource -Name NuGetOrg -Location https://api.nuget.org/v3/index.json -ProviderName NuGet);'

        if ($boolThrowErrorForMissingPackageSource -eq $true) {
            Write-Error $strMessage
        } elseif ($boolThrowWarningForMissingPackageSource -eq $true) {
            Write-Warning $strMessage
        }
    }

    return $boolPackageSourceFound
}

function Get-PackagesUsingHashtable {
    <#
    .SYNOPSIS
    Gets a list of installed "software packages" (typically NuGet packages) for each
    entry in a hashtable.

    .DESCRIPTION
    The Get-PackagesUsingHashtable function steps through each entry in the supplied
    hashtable. If a corresponding package is installed, then the information about
    the newest version of that package is stored in the value of the hashtable entry
    corresponding to the software package.

    .PARAMETER ReferenceToHashtable
    Is a reference to a hashtable. The value of the reference should be a hashtable
    with keys that are the names software packages and values that are initialized
    to be $null.

    .EXAMPLE
    $hashtablePackageNameToInstalledPackageMetadata = @{}
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord.Math', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord.Statistics', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord.MachineLearning', $null)
    $refHashtablePackageNameToInstalledPackages = [ref]$hashtablePackageNameToInstalledPackageMetadata
    Get-PackagesUsingHashtable -ReferenceToHashtable $refHashtablePackageNameToInstalledPackages

    This example checks each of the four software packages specified. For each software
    package specified, if the software package is installed, the value of the hashtable
    entry will be set to the newest-installed version of the package. If the software
    package is not installed, the value of the hashtable entry remains $null.

    .OUTPUTS
    None

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

    # Version 1.0.20240401.0

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtable
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
        Write-Warning 'Get-PackagesUsingHashtable requires PowerShell version 5.0 or newer.'
        return
    }

    $WarningPreferenceAtStartOfFunction = $WarningPreference
    $VerbosePreferenceAtStartOfFunction = $VerbosePreference
    $DebugPreferenceAtStartOfFunction = $DebugPreference

    $arrPackagesToGet = @(($ReferenceToHashtable.Value).Keys)

    $WarningPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $arrPackagesInstalled = @(Get-Package)
    $WarningPreference = $WarningPreferenceAtStartOfFunction
    $VerbosePreference = $VerbosePreferenceAtStartOfFunction
    $DebugPreference = $DebugPreferenceAtStartOfFunction

    for ($intCounter = 0; $intCounter -lt $arrPackagesToGet.Count; $intCounter++) {
        Write-Information ('Checking for ' + $arrPackagesToGet[$intCounter] + ' software package...')
        $arrMatchingPackages = @($arrPackagesInstalled | Where-Object { $_.Name -eq $arrPackagesToGet[$intCounter] })
        if ($arrMatchingPackages.Count -eq 0) {
            ($ReferenceToHashtable.Value).Item($arrPackagesToGet[$intCounter]) = $null
        } else {
            ($ReferenceToHashtable.Value).Item($arrPackagesToGet[$intCounter]) = $arrMatchingPackages[0]
        }
    }
}

function Test-PackageInstalledUsingHashtable {
    <#
    .SYNOPSIS
    Tests to see if a software package (typically a NuGet package) is installed based
    on entries in a hashtable. If the software package is not installed, an error or
    warning message may optionally be displayed.

    .DESCRIPTION
    The Test-PackageInstalledUsingHashtable function steps through each entry in the
    supplied hashtable and, if there are any software packages not installed, it
    optionally throws an error or warning for each software package that is not
    installed. If all software packages are installed, the function returns $true;
    otherwise, if any software package is not installed, the function returns $false.

    .PARAMETER ReferenceToHashtableOfInstalledPackages
    Is a reference to a hashtable. The hashtable must have keys that are the names of
    software packages with each key's value populated with
    Microsoft.PackageManagement.Packaging.SoftwareIdentity objects (the result of
    Get-Package). If a software package is not installed, the value of the hashtable
    entry should be $null.

    .PARAMETER ReferenceToHashtableOfSkippingDependencies
    Is a reference to a hashtable. The hashtable must have keys that are the names of
    software packages with each key's value populated with a boolean value. The boolean
    indicates whether the software package should be installed without its
    dependencies. Generally, dependencies should not be skipped, so the default value
    for each key should be $false. However, sometimes the Install-Package command
    throws an erroneous dependency loop error, but in investigating its dependencies in
    the package's .nuspec file, you may find that the version of .NET that you will use
    has no dependencies. In this case, it's safe to use -SkipDependencies.

    This can also be verified here:
    https://www.nuget.org/packages/<PackageName>/#dependencies-body-tab

    If this parameter is not supplied, or if a key-value pair is not supplied in the
    hashtable for a given software package, the script will default to not skipping the
    software package's dependencies.

    .PARAMETER ThrowErrorIfPackageNotInstalled
    Is a switch parameter. If this parameter is specified, an error is thrown for each
    software package that is not installed. If this parameter is not specified, no
    error is thrown.

    .PARAMETER ThrowWarningIfPackageNotInstalled
    Is a switch parameter. If this parameter is specified, a warning is thrown for each
    software package that is not installed. If this parameter is not specified, or if
    the ThrowErrorIfPackageNotInstalled parameter was specified, no warning is thrown.

    .PARAMETER ReferenceToHashtableOfCustomNotInstalledMessages
    Is a reference to a hashtable. The hashtable must have keys that are custom error
    or warning messages (string) to be displayed if one or more software packages are
    not installed. The value for each key must be an array of software package names
    (strings) relevant to that error or warning message.

    If this parameter is not supplied, or if a custom error or warning message is not
    supplied in the hashtable for a given software package, the script will default to
    using the following message:

    <PACKAGENAME> software package not found. Please install it and then try again.
    You can install the <PACKAGENAME> software package by running the following
    command:
    Install-Package -ProviderName NuGet -Name <PACKAGENAME> -Force -Scope CurrentUser;

    .PARAMETER ReferenceToArrayOfMissingPackages
    Is a reference to an array. The array must be initialized to be empty. If any
    software packages are not installed, the names of those software packages are added
    to the array.

    .EXAMPLE
    $hashtablePackageNameToInstalledPackageMetadata = @{}
    $hashtablePackageNameToInstalledPackageMetadata.Add('Azure.Core', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Microsoft.Identity.Client', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Azure.Identity', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord', $null)
    $refHashtablePackageNameToInstalledPackages = [ref]$hashtablePackageNameToInstalledPackageMetadata
    Get-PackagesUsingHashtable -ReferenceToHashtable $refHashtablePackageNameToInstalledPackages
    $hashtableCustomNotInstalledMessageToPackageNames = @{}
    $strAzureIdentityNotInstalledMessage = 'Azure.Core, Microsoft.Identity.Client, and/or Azure.Identity packages were not found. Please install the Azure.Identity package and its dependencies and then try again.' + [System.Environment]::NewLine + 'You can install the Azure.Identity package and its dependencies by running the following command:' + [System.Environment]::NewLine + 'Install-Package -ProviderName NuGet -Name ''Azure.Identity'' -Force -Scope CurrentUser;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    $hashtableCustomNotInstalledMessageToPackageNames.Add($strAzureIdentityNotInstalledMessage, @('Azure.Core', 'Microsoft.Identity.Client', 'Azure.Identity'))
    $refhashtableCustomNotInstalledMessageToPackageNames = [ref]$hashtableCustomNotInstalledMessageToPackageNames
    $boolResult = Test-PackageInstalledUsingHashtable -ReferenceToHashtableOfInstalledPackages $refHashtablePackageNameToInstalledPackages -ThrowErrorIfPackageNotInstalled -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToPackageNames

    This example checks to see if the Azure.Core, Microsoft.Identity.Client,
    Azure.Identity, and Accord packages are installed. If any of these packages are not
    installed, an error is thrown and $boolResult is set to $false. Because a custom
    error message was specified for the Azure.Core, Microsoft.Identity.Client, and
    Azure.Identity packages, if any one of those is missing, the custom error message
    is thrown just once. However, if Accord is missing, a separate error message would
    be thrown. If all packages are installed, $boolResult is set to $true.

    .OUTPUTS
    [boolean] - Returns $true if all Packages are installed; otherwise, returns $false.
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
    [OutputType([Boolean])]
    param (
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfInstalledPackages,
        [Parameter(Mandatory = $false)][ref]$ReferenceToHashtableOfSkippingDependencies,
        [Parameter(Mandatory = $false)][switch]$ThrowErrorIfPackageNotInstalled,
        [Parameter(Mandatory = $false)][switch]$ThrowWarningIfPackageNotInstalled,
        [Parameter(Mandatory = $false)][ref]$ReferenceToHashtableOfCustomNotInstalledMessages,
        [Parameter(Mandatory = $false)][ref]$ReferenceToArrayOfMissingPackages
    )

    $boolThrowErrorForMissingPackage = $false
    $boolThrowWarningForMissingPackage = $false

    if ($ThrowErrorIfPackageNotInstalled.IsPresent -eq $true) {
        $boolThrowErrorForMissingPackage = $true
    } elseif ($ThrowWarningIfPackageNotInstalled.IsPresent -eq $true) {
        $boolThrowWarningForMissingPackage = $true
    }

    $boolResult = $true

    $hashtableMessagesToThrowForMissingPackage = @{}
    $hashtablePackageNameToCustomMessageToThrowForMissingPackage = @{}
    if ($null -ne $ReferenceToHashtableOfCustomNotInstalledMessages) {
        $arrMessages = @(($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Keys)
        foreach ($strMessage in $arrMessages) {
            $hashtableMessagesToThrowForMissingPackage.Add($strMessage, $false)

            ($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Item($strMessage) | ForEach-Object {
                $hashtablePackageNameToCustomMessageToThrowForMissingPackage.Add($_, $strMessage)
            }
        }
    }

    $arrPackageNames = @(($ReferenceToHashtableOfInstalledPackages.Value).Keys)
    foreach ($strPackageName in $arrPackageNames) {
        if ($null -eq ($ReferenceToHashtableOfInstalledPackages.Value).Item($strPackageName)) {
            # Package not installed
            $boolResult = $false

            if ($hashtablePackageNameToCustomMessageToThrowForMissingPackage.ContainsKey($strPackageName) -eq $true) {
                $strMessage = $hashtablePackageNameToCustomMessageToThrowForMissingPackage.Item($strPackageName)
                $hashtableMessagesToThrowForMissingPackage.Item($strMessage) = $true
            } else {
                if ($null -ne $ReferenceToHashtableOfSkippingDependencies) {
                    if (($ReferenceToHashtableOfSkippingDependencies.Value).ContainsKey($strPackageName) -eq $true) {
                        $boolSkipDependencies = ($ReferenceToHashtableOfSkippingDependencies.Value).Item($strPackageName)
                    } else {
                        $boolSkipDependencies = $false
                    }
                } else {
                    $boolSkipDependencies = $false
                }

                if ($boolSkipDependencies -eq $true) {
                    $strMessage = $strPackageName + ' software package not found. Please install it and then try again.' + [System.Environment]::NewLine + 'You can install the ' + $strPackageName + ' package by running the following command:' + [System.Environment]::NewLine + 'Install-Package -ProviderName NuGet -Name ''' + $strPackageName + ''' -Force -Scope CurrentUser -SkipDependencies;' + [System.Environment]::NewLine + [System.Environment]::NewLine
                } else {
                    $strMessage = $strPackageName + ' software package not found. Please install it and then try again.' + [System.Environment]::NewLine + 'You can install the ' + $strPackageName + ' package by running the following command:' + [System.Environment]::NewLine + 'Install-Package -ProviderName NuGet -Name ''' + $strPackageName + ''' -Force -Scope CurrentUser;' + [System.Environment]::NewLine + [System.Environment]::NewLine
                }

                $hashtableMessagesToThrowForMissingPackage.Add($strMessage, $true)
            }

            if ($null -ne $ReferenceToArrayOfMissingPackages) {
                ($ReferenceToArrayOfMissingPackages.Value) += $strPackageName
            }
        }
    }

    if ($boolThrowErrorForMissingPackage -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForMissingPackage.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingPackage.Item($strMessage) -eq $true) {
                Write-Error $strMessage
            }
        }
    } elseif ($boolThrowWarningForMissingPackage -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForMissingPackage.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingPackage.Item($strMessage) -eq $true) {
                Write-Warning $strMessage
            }
        }
    }

    return $boolResult
}

function Get-DLLPathsForPackagesUsingHashtable {
    <#
    .SYNOPSIS
    Using a hashtable of installed software package metadata, gets the path to the
    .dll file(s) within each software package that is most appropriate to use

    .DESCRIPTION
    Software packages contain .dll files for different .NET Framework versions. This
    function steps through each entry in the supplied hashtable. If a corresponding
    package is installed, then the path to the .dll file(s) within the package that is
    most appropriate to use is stored in the value of the hashtable entry corresponding
    to the software package.

    .PARAMETER ReferenceToHashtableOfInstalledPackages
    Is a reference to a hashtable. The hashtable must have keys that are the names of
    software packages with each key's value populated with
    Microsoft.PackageManagement.Packaging.SoftwareIdentity objects (the result of
    Get-Package). If a software package is not installed, the value of the hashtable
    entry should be $null.

    .PARAMETER ReferenceToHashtableOfSpecifiedDotNETVersions
    Is an optional parameter. If supplied, it must be a reference to a hashtable. The
    hashtable must have keys that are the names of software packages with each key's
    value populated with a string that is the version of .NET Framework that the
    software package is to be used with. If a key-value pair is not supplied in the
    hashtable for a given software package, the function will default to doing its best
    to select the most appropriate version of the software package given the current
    operating environment and PowerShell version.

    .PARAMETER ReferenceToHashtableOfEffectiveDotNETVersions
    Is initially a reference to an empty hashtable. When execution completes, the
    hashtable will be populated with keys that are the names of the software packages
    specified in the hashtable referenced by the
    ReferenceToHashtableOfInstalledPackages parameter. The value of each entry will be
    a string that is the folder corresponding to the version of .NET that makes the
    most sense given the current platform and .NET Framework version. If no suitable
    folder is found, the value of the hashtable entry remains an empty string.

    For example, reference the following folder name taxonomy at nuget.org:
    https://www.nuget.org/packages/System.Text.Json#supportedframeworks-body-tab

    .PARAMETER ReferenceToHashtableOfDLLPaths
    Is initially a reference to an empty hashtable. When execution completes, the
    hashtable will be populated with keys that are the names of the software packages
    specified in the hashtable referenced by the
    ReferenceToHashtableOfInstalledPackages parameter. The value of each entry will be
    an array populated with the path to the .dll file(s) within the package that are
    most appropriate to use, given the current platform and .NET Framework version.
    If no suitable DLL file is found, the array will be empty.

    .EXAMPLE
    $hashtablePackageNameToInstalledPackageMetadata = @{}
    $hashtablePackageNameToInstalledPackageMetadata.Add('Azure.Core', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Microsoft.Identity.Client', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Azure.Identity', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord', $null)
    $refHashtablePackageNameToInstalledPackages = [ref]$hashtablePackageNameToInstalledPackageMetadata
    Get-PackagesUsingHashtable -ReferenceToHashtable $refHashtablePackageNameToInstalledPackages
    $hashtableCustomNotInstalledMessageToPackageNames = @{}
    $strAzureIdentityNotInstalledMessage = 'Azure.Core, Microsoft.Identity.Client, and/or Azure.Identity packages were not found. Please install the Azure.Identity package and its dependencies and then try again.' + [System.Environment]::NewLine + 'You can install the Azure.Identity package and its dependencies by running the following command:' + [System.Environment]::NewLine + 'Install-Package -ProviderName NuGet -Name 'Azure.Identity' -Force -Scope CurrentUser;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    $hashtableCustomNotInstalledMessageToPackageNames.Add($strAzureIdentityNotInstalledMessage, @('Azure.Core', 'Microsoft.Identity.Client', 'Azure.Identity'))
    $refhashtableCustomNotInstalledMessageToPackageNames = [ref]$hashtableCustomNotInstalledMessageToPackageNames
    $boolResult = Test-PackageInstalledUsingHashtable -ReferenceToHashtableOfInstalledPackages $refHashtablePackageNameToInstalledPackages -ThrowErrorIfPackageNotInstalled -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToPackageNames
    if ($boolResult -eq $false) { return }
    $hashtablePackageNameToEffectiveDotNETVersions = @{}
    $refHashtablePackageNameToEffectiveDotNETVersions = [ref]$hashtablePackageNameToEffectiveDotNETVersions
    $hashtablePackageNameToDLLPaths = @{}
    $refHashtablePackageNameToDLLPaths = [ref]$hashtablePackageNameToDLLPaths
    Get-DLLPathsForPackagesUsingHashtable -ReferenceToHashtableOfInstalledPackages $refHashtablePackageNameToInstalledPackages -ReferenceToHashtableOfEffectiveDotNETVersions $refHashtablePackageNameToEffectiveDotNETVersions -ReferenceToHashtableOfDLLPaths $refHashtablePackageNameToDLLPaths

    This example checks each of the four software packages specified. For each software
    package specified, if the software package is installed, the value of the hashtable
    entry will be set to the path to the .dll file(s) within the package that are most
    appropriate to use, given the current platform and .NET Framework version. If no
    suitable DLL file is found, the value of the hashtable entry remains an empty array
    (@()).

    .OUTPUTS
    None

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
    param (
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfInstalledPackages,
        [Parameter(Mandatory = $false)][ref]$ReferenceToHashtableOfSpecifiedDotNETVersions,
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfEffectiveDotNETVersions,
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfDLLPaths
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
        Write-Warning 'Get-DLLPathsForPackagesUsingHashtable requires PowerShell version 5.0 or newer.'
        return
    }

    $arrPackageNames = @(($ReferenceToHashtableOfInstalledPackages.Value).Keys)
    foreach ($strPackageName in $arrPackageNames) {
        ($ReferenceToHashtableOfEffectiveDotNETVersions.Value).Add($strPackageName, '')
        ($ReferenceToHashtableOfDLLPaths.Value).Add($strPackageName, @())
    }

    # Get the base folder path for each package
    $hashtablePackageNameToBaseFolderPath = @{}
    foreach ($strPackageName in $arrPackageNames) {
        if ($null -ne ($ReferenceToHashtableOfInstalledPackages.Value).Item($strPackageName)) {
            $strPackageFilePath = ($ReferenceToHashtableOfInstalledPackages.Value).Item($strPackageName).Source
            $strPackageFilePath = $strPackageFilePath.Replace('file:///', '')
            $strPackageFileParentFolderPath = [System.IO.Path]::GetDirectoryName($strPackageFilePath)

            $hashtablePackageNameToBaseFolderPath.Add($strPackageName, $strPackageFileParentFolderPath)
        }
    }

    # Determine the current platform
    $boolIsLinux = $false
    if (Test-Path variable:\IsLinux) {
        if ($IsLinux -eq $true) {
            $boolIsLinux = $true
        }
    }

    $boolIsMacOS = $false
    if (Test-Path variable:\IsMacOS) {
        if ($IsMacOS -eq $true) {
            $boolIsMacOS = $true
        }
    }

    if ($boolIsLinux -eq $true) {
        if (($versionPS -ge [version]'7.5') -and ($versionPS -lt [version]'7.6')) {
            # .NET 9.0
            $arrDotNETVersionPreferenceOrder = @('net9.0', 'net8.0', 'net7.0', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.4') -and ($versionPS -lt [version]'7.5')) {
            # .NET 8.0
            $arrDotNETVersionPreferenceOrder = @('net8.0', 'net7.0', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.3') -and ($versionPS -lt [version]'7.4')) {
            # .NET 7.0
            $arrDotNETVersionPreferenceOrder = @('net7.0', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.2') -and ($versionPS -lt [version]'7.3')) {
            # .NET 6.0
            $arrDotNETVersionPreferenceOrder = @('net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.1') -and ($versionPS -lt [version]'7.2')) {
            # .NET 5.0
            $arrDotNETVersionPreferenceOrder = @('net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.0') -and ($versionPS -lt [version]'7.1')) {
            # .NET Core 3.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.2') -and ($versionPS -lt [version]'7.0')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.1') -and ($versionPS -lt [version]'6.2')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.0') -and ($versionPS -lt [version]'6.1')) {
            # .NET Core 2.0
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } else {
            # A future, undefined version of PowerShell
            $arrDotNETVersionPreferenceOrder = @('net15.0', 'net14.0', 'net13.0', 'net12.0', 'net11.0', 'net10.0', 'net9.0', 'net8.0', 'net7.0', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        }
    } elseif ($boolIsMacOS -eq $true) {
        if (($versionPS -ge [version]'7.5') -and ($versionPS -lt [version]'7.6')) {
            # .NET 9.0
            $arrDotNETVersionPreferenceOrder = @('net9.0-macos', 'net9.0', 'net8.0-macos', 'net8.0', 'net7.0-macos', 'net7.0', 'net6.0-macos', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.4') -and ($versionPS -lt [version]'7.5')) {
            # .NET 8.0
            $arrDotNETVersionPreferenceOrder = @('net8.0-macos', 'net8.0', 'net7.0-macos', 'net7.0', 'net6.0-macos', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.3') -and ($versionPS -lt [version]'7.4')) {
            # .NET 7.0
            $arrDotNETVersionPreferenceOrder = @('net7.0-macos', 'net7.0', 'net6.0-macos', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.2') -and ($versionPS -lt [version]'7.3')) {
            # .NET 6.0
            $arrDotNETVersionPreferenceOrder = @('net6.0-macos', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.1') -and ($versionPS -lt [version]'7.2')) {
            # .NET 5.0
            $arrDotNETVersionPreferenceOrder = @('net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.0') -and ($versionPS -lt [version]'7.1')) {
            # .NET Core 3.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.2') -and ($versionPS -lt [version]'7.0')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.1') -and ($versionPS -lt [version]'6.2')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.0') -and ($versionPS -lt [version]'6.1')) {
            # .NET Core 2.0
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } else {
            # A future, undefined version of PowerShell
            $arrDotNETVersionPreferenceOrder = @('net15.0-macos', 'net15.0', 'net14.0-macos', 'net14.0', 'net13.0-macos', 'net13.0', 'net12.0-macos', 'net12.0', 'net11.0-macos', 'net11.0', 'net10.0-macos', 'net10.0', 'net9.0-macos', 'net9.0', 'net8.0-macos', 'net8.0', 'net7.0-macos', 'net7.0', 'net6.0-macos', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        }
    } else {
        # Windows
        if (($versionPS -ge [version]'7.5') -and ($versionPS -lt [version]'7.6')) {
            # .NET 9.0
            $arrDotNETVersionPreferenceOrder = @('net9.0-windows', 'net9.0', 'net8.0-windows', 'net8.0', 'net7.0-windows', 'net7.0', 'net6.0-windows', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.4') -and ($versionPS -lt [version]'7.5')) {
            # .NET 8.0
            $arrDotNETVersionPreferenceOrder = @('net8.0-windows', 'net8.0', 'net7.0-windows', 'net7.0', 'net6.0-windows', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.3') -and ($versionPS -lt [version]'7.4')) {
            # .NET 7.0
            $arrDotNETVersionPreferenceOrder = @('net7.0-windows', 'net7.0', 'net6.0-windows', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.2') -and ($versionPS -lt [version]'7.3')) {
            # .NET 6.0
            $arrDotNETVersionPreferenceOrder = @('net6.0-windows', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.1') -and ($versionPS -lt [version]'7.2')) {
            # .NET 5.0
            $arrDotNETVersionPreferenceOrder = @('net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.0') -and ($versionPS -lt [version]'7.1')) {
            # .NET Core 3.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.2') -and ($versionPS -lt [version]'7.0')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.1') -and ($versionPS -lt [version]'6.2')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.0') -and ($versionPS -lt [version]'6.1')) {
            # .NET Core 2.0
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'5.0') -and ($versionPS -lt [version]'6.0')) {
            if ((Test-Path 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full') -eq $true) {
                $intDotNETFrameworkRelease = (Get-ItemPropertyValue -LiteralPath 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release)
                # if ($intDotNETFrameworkRelease -ge 533320) {
                #     # .NET Framework 4.8.1
                #     $arrDotNETVersionPreferenceOrder = @('net481', 'net48', 'net472', 'net471', 'netstandard2.0', 'net47', 'net463', 'net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 528040) {
                #     # .NET Framework 4.8
                #     $arrDotNETVersionPreferenceOrder = @('net48', 'net472', 'net471', 'netstandard2.0', 'net47', 'net463', 'net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 461808) {
                #     # .NET Framework 4.7.2
                #     $arrDotNETVersionPreferenceOrder = @('net472', 'net471', 'netstandard2.0', 'net47', 'net463', 'net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 461308) {
                #     # .NET Framework 4.7.1
                #     $arrDotNETVersionPreferenceOrder = @('net471', 'netstandard2.0', 'net47', 'net463', 'net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 460798) {
                #     # .NET Framework 4.7
                #     $arrDotNETVersionPreferenceOrder = @('net47', 'net463', 'net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 394802) {
                #     # .NET Framework 4.6.2
                #     $arrDotNETVersionPreferenceOrder = @('net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 394254) {
                #     # .NET Framework 4.6.1
                #     $arrDotNETVersionPreferenceOrder = @('net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 393295) {
                #     # .NET Framework 4.6
                #     $arrDotNETVersionPreferenceOrder = @('net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 379893) {
                #     # .NET Framework 4.5.2
                #     $arrDotNETVersionPreferenceOrder = @('net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 378675) {
                #     # .NET Framework 4.5.1
                #     $arrDotNETVersionPreferenceOrder = @('net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 378389) {
                #     # .NET Framework 4.5
                #     $arrDotNETVersionPreferenceOrder = @('net45', 'net40')
                # } else {
                #     # .NET Framework 4.5 or newer not found?
                #     # This should not be possible since this function requires
                #     # PowerShell 5.0 or newer, PowerShell 5.0 requires WMF 5.0, and
                #     # WMF 5.0 requires .NET Framework 4.5 or newer.
                #     Write-Warning 'The .NET Framework 4.5 or newer was not found. This should not be possible since this function requires PowerShell 5.0 or newer, PowerShell 5.0 requires WMF 5.0, and WMF 5.0 requires .NET Framework 4.5 or newer.'
                #     return
                # }
                if ($intDotNETFrameworkRelease -ge 533320) {
                    # .NET Framework 4.8.1
                    $arrDotNETVersionPreferenceOrder = @('net481', 'net48', 'net472', 'net471', 'net47', 'net463', 'net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard2.0', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 528040) {
                    # .NET Framework 4.8
                    $arrDotNETVersionPreferenceOrder = @('net48', 'net472', 'net471', 'net47', 'net463', 'net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard2.0', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 461808) {
                    # .NET Framework 4.7.2
                    $arrDotNETVersionPreferenceOrder = @('net472', 'net471', 'net47', 'net463', 'net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard2.0', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 461308) {
                    # .NET Framework 4.7.1
                    $arrDotNETVersionPreferenceOrder = @('net471', 'net47', 'net463', 'net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard2.0', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 460798) {
                    # .NET Framework 4.7
                    $arrDotNETVersionPreferenceOrder = @('net47', 'net463', 'net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 394802) {
                    # .NET Framework 4.6.2
                    $arrDotNETVersionPreferenceOrder = @('net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 394254) {
                    # .NET Framework 4.6.1
                    $arrDotNETVersionPreferenceOrder = @('net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 393295) {
                    # .NET Framework 4.6
                    $arrDotNETVersionPreferenceOrder = @('net46', 'net452', 'net451', 'net45', 'net40')
                } elseif ($intDotNETFrameworkRelease -ge 379893) {
                    # .NET Framework 4.5.2
                    $arrDotNETVersionPreferenceOrder = @('net452', 'net451', 'net45', 'net40')
                } elseif ($intDotNETFrameworkRelease -ge 378675) {
                    # .NET Framework 4.5.1
                    $arrDotNETVersionPreferenceOrder = @('net451', 'net45', 'net40')
                } elseif ($intDotNETFrameworkRelease -ge 378389) {
                    # .NET Framework 4.5
                    $arrDotNETVersionPreferenceOrder = @('net45', 'net40')
                } else {
                    # .NET Framework 4.5 or newer not found?
                    # This should not be possible since this function requires
                    # PowerShell 5.0 or newer, PowerShell 5.0 requires WMF 5.0, and
                    # WMF 5.0 requires .NET Framework 4.5 or newer.
                    Write-Warning 'The .NET Framework 4.5 or newer was not found. This should not be possible since this function requires PowerShell 5.0 or newer, PowerShell 5.0 requires WMF 5.0, and WMF 5.0 requires .NET Framework 4.5 or newer.'
                    return
                }
            }
        } else {
            # A future, undefined version of PowerShell
            $arrDotNETVersionPreferenceOrder = @('net15.0-windows', 'net15.0', 'net14.0-windows', 'net14.0', 'net13.0-windows', 'net13.0', 'net12.0-windows', 'net12.0', 'net11.0-windows', 'net11.0', 'net10.0-windows', 'net10.0', 'net9.0-windows', 'net9.0', 'net8.0-windows', 'net8.0', 'net7.0-windows', 'net7.0', 'net6.0-windows', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        }
    }

    foreach ($strPackageName in $arrPackageNames) {
        if ($null -ne $hashtablePackageNameToBaseFolderPath.Item($strPackageName)) {
            $strBaseFolderPath = ($hashtablePackageNameToBaseFolderPath.Item($strPackageName))

            $strDLLFolderPath = ''
            if ($null -ne $ReferenceToHashtableOfSpecifiedDotNETVersions) {
                if ($null -ne ($ReferenceToHashtableOfSpecifiedDotNETVersions.Value).Item($strPackageName)) {
                    $strDotNETVersion = ($ReferenceToHashtableOfSpecifiedDotNETVersions.Value).Item($strPackageName)

                    if ([string]::IsNullOrEmpty($strDotNETVersion) -eq $false) {
                        $strDLLFolderPath = Join-Path -Path $strBaseFolderPath -ChildPath 'lib'
                        $strDLLFolderPath = Join-Path -Path $strDLLFolderPath -ChildPath $strDotNETVersion

                        # Search this folder for .dll files and add them to the array
                        if (Test-Path -Path $strDLLFolderPath -PathType Container) {
                            $arrDLLFiles = @(Get-ChildItem -Path $strDLLFolderPath -Filter '*.dll' -File -Recurse)
                            if ($arrDLLFiles.Count -gt 0) {
                                # One or more DLL files found
                                ($ReferenceToHashtableOfEffectiveDotNETVersions.Value).Item($strPackageName) = $strDotNETVersion
                                ($ReferenceToHashtableOfDLLPaths.Value).Item($strPackageName) = @($arrDLLFiles | ForEach-Object {
                                        $_.FullName
                                    })
                            }
                        } else {
                            # The specified .NET version folder does not exist
                            # Set the DLL folder path to an empty string to then do a
                            # search for a usable folder
                            $strDLLFolderPath = ''
                        }
                    }
                }
            }

            if ([string]::IsNullOrEmpty($strDLLFolderPath)) {
                # Do a search for a usable folder

                foreach ($strDotNETVersion in $arrDotNETVersionPreferenceOrder) {
                    $strDLLFolderPath = Join-Path -Path $strBaseFolderPath -ChildPath 'lib'
                    $strDLLFolderPath = Join-Path -Path $strDLLFolderPath -ChildPath $strDotNETVersion

                    if (Test-Path -Path $strDLLFolderPath -PathType Container) {
                        $arrDLLFiles = @(Get-ChildItem -Path $strDLLFolderPath -Filter '*.dll' -File -Recurse)
                        if ($arrDLLFiles.Count -gt 0) {
                            # One or more DLL files found
                            ($ReferenceToHashtableOfEffectiveDotNETVersions.Value).Item($strPackageName) = $strDotNETVersion
                            ($ReferenceToHashtableOfDLLPaths.Value).Item($strPackageName) = @($arrDLLFiles | ForEach-Object {
                                    $_.FullName
                                })
                            break
                        }
                    }
                }
            }
        }
    }
}

function Measure-EuclideanDistance($Point1, $Point2) {
    $doubleSum = [double]0
    for ($i = 0; $i -lt $Point1.Length; $i++) {
        $doubleSum += [Math]::Pow($Point1[$i] - $Point2[$i], 2)
    }
    return [Math]::Sqrt($doubleSum)
}

function Invoke-KMeansClusteringForSpecifiedNumberOfClusters {
    param (
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfNumberOfClustersToArtifacts,
        [Parameter(Mandatory = $true)][ValidateScript({ $_ -ge 1 })][int]$NumberOfClusters,
        [Parameter(Mandatory = $true)][ref]$ReferenceToTwoDimensionalArrayOfEmbeddings,
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfEuclideanDistancesBetweenItems,
        [Parameter(Mandatory = $true)][ref]$ReferenceToCentroidOverWholeDataset,
        [Parameter(Mandatory = $false)][switch]$DoNotCalculateExtendedStatistics
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

    if (($ReferenceToHashtableOfNumberOfClustersToArtifacts.Value).ContainsKey($NumberOfClusters) -eq $true) {
        # The specified number of clusters has already been processed
        return
    }

    $boolCalculateExtendedStatistics = $true
    if ($null -ne $DoNotCalculateExtendedStatistics) {
        if ($DoNotCalculateExtendedStatistics.IsPresent -eq $true) {
            $boolCalculateExtendedStatistics = $false
        }
    }

    $PSCustomObjectKMeansArtifacts = New-Object -TypeName 'PSCustomObject'
    $PSCustomObjectKMeansArtifacts | Add-Member -MemberType NoteProperty -Name 'NumberOfClusters' -Value $NumberOfClusters
    $PSCustomObjectKMeansArtifacts | Add-Member -MemberType NoteProperty -Name 'HashtableClustersToItemsAndDistances' -Value $null
    $PSCustomObjectKMeansArtifacts | Add-Member -MemberType NoteProperty -Name 'WithinClusterSumOfSquares' -Value $null
    if ($boolCalculateExtendedStatistics -eq $true) {
        $PSCustomObjectKMeansArtifacts | Add-Member -MemberType NoteProperty -Name 'SilhouetteScore' -Value $null
        $PSCustomObjectKMeansArtifacts | Add-Member -MemberType NoteProperty -Name 'DaviesBouldinIndex' -Value $null
        $PSCustomObjectKMeansArtifacts | Add-Member -MemberType NoteProperty -Name 'CalinskiHarabaszIndex' -Value $null
    }

    $kmeans = New-Object -TypeName 'Accord.MachineLearning.KMeans' -ArgumentList @($NumberOfClusters)
    [void]($kmeans.Learn($ReferenceToTwoDimensionalArrayOfEmbeddings.Value))
    $arrClusterNumberAssignmentsForEachItem = $kmeans.Clusters.Decide($ReferenceToTwoDimensionalArrayOfEmbeddings.Value)

    #region Create Hashtable for Efficient Lookup of Cluster # to Associated Items #####
    if ($NumberOfClusters -eq 1) {
        Write-Information ('1 cluster: Creating hashtable for efficient lookup of cluster # to associated items...')
    } else {
        Write-Information ([string]$NumberOfClusters + ' clusters: Creating hashtable for efficient lookup of cluster # to associated items...')
    }
    # Create a hashtable for easier lookup of cluster number to comment index number
    $hashtableClustersToItems = @{}
    if ($versionPS -ge ([version]'6.0')) {
        for ($intCounterA = 0; $intCounterA -lt $NumberOfClusters; $intCounterA++) {
            $hashtableClustersToItems.Add($intCounterA, (New-Object -TypeName 'System.Collections.Generic.List[PSCustomObject]'))
        }
    } else {
        # On Windows PowerShell (versions older than 6.x), we use an ArrayList instead
        # of a generic list
        # TODO: Fill in rationale for this
        #
        # Technically, in older versions of PowerShell, the type in the ArrayList will
        # be a PSObject; but that does not matter for our purposes.
        for ($intCounterA = 0; $intCounterA -lt $NumberOfClusters; $intCounterA++) {
            $hashtableClustersToItems.Add($intCounterA, (New-Object -TypeName 'System.Collections.ArrayList'))
        }
    }

    # Populate the hashtable of cluster number -> associated items
    $intCounterMax = $arrClusterNumberAssignmentsForEachItem.Length
    if ($versionPS -ge ([version]'6.0')) {
        # PowerShell v6.0 or newer
        for ($intCounterA = 0; $intCounterA -lt $intCounterMax; $intCounterA++) {
            $intTopicNumber = $arrClusterNumberAssignmentsForEachItem[$intCounterA]

            # Add the updated object to the list
            ($hashtableClustersToItems.Item($intTopicNumber)).Add($intCounterA)
        }
    } else {
        # Windows PowerShell 5.0 or 5.1
        for ($intCounterA = 0; $intCounterA -lt $intCounterMax; $intCounterA++) {
            $intTopicNumber = $arrClusterNumberAssignmentsForEachItem[$intCounterA]

            # Add the updated object to the list
            [void](($hashtableClustersToItems.Item($intTopicNumber)).Add($intCounterA))
        }
    }
    #endregion Create Hashtable for Efficient Lookup of Cluster # to Associated Items #####

    #region Create Hashtable Including Euclidian Distance from Item to Centroid ########
    if ($NumberOfClusters -eq 1) {
        Write-Information ('1 cluster: Creating hashtable including Euclidian distance from each item to its cluster centroid...')
    } else {
        Write-Information ([string]$NumberOfClusters + ' clusters: Creating hashtable including Euclidian distance from each item to its cluster centroid...')
    }
    $PSCustomObjectKMeansArtifacts.HashtableClustersToItemsAndDistances = @{}
    if ($versionPS -ge ([version]'6.0')) {
        for ($intCounterA = 0; $intCounterA -lt $NumberOfClusters; $intCounterA++) {
            $PSCustomObjectKMeansArtifacts.HashtableClustersToItemsAndDistances.Add($intCounterA, (New-Object -TypeName 'System.Collections.Generic.List[PSCustomObject]'))
        }
    } else {
        # On Windows PowerShell (versions older than 6.x), we use an ArrayList instead
        # of a generic list
        # TODO: Fill in rationale for this
        #
        # Technically, in older versions of PowerShell, the type in the ArrayList will
        # be a PSObject; but that does not matter for our purposes.
        for ($intCounterA = 0; $intCounterA -lt $NumberOfClusters; $intCounterA++) {
            $PSCustomObjectKMeansArtifacts.HashtableClustersToItemsAndDistances.Add($intCounterA, (New-Object -TypeName 'System.Collections.ArrayList'))
        }
    }

    #region Collect Stats/Objects Needed for Writing Progress ##########################
    $intProgressReportingFrequency = 50
    $intTotalItems = $arrInputCSV.Count
    $strProgressActivity = 'Performing k-means clustering'
    if ($NumberOfClusters -eq 1) {
        $strProgressStatus = 'Calculating distances from items to cluster centroid (1 cluster)'
    } else {
        $strProgressStatus = 'Calculating distances from items to cluster centroids (' + [string]$NumberOfClusters + ' clusters)'
    }
    $strProgressCurrentOperationPrefix = 'Processing item'
    $timedateStartOfLoop = Get-Date
    # Create a queue for storing lagging timestamps for ETA calculation
    $queueLaggingTimestamps = New-Object System.Collections.Queue
    $queueLaggingTimestamps.Enqueue($timedateStartOfLoop)
    #endregion Collect Stats/Objects Needed for Writing Progress ##########################

    $intCounterLoop = 0
    if ($versionPS -ge ([version]'6.0')) {
        # PowerShell v6.0 or newer
        for ($intCounterA = 0; $intCounterA -lt $NumberOfClusters; $intCounterA++) {
            if ($NumberOfClusters -eq 1) {
                $ReferenceToCentroidOverWholeDataset.Value = ($kmeans.Clusters.Centroids)[$intCounterA]
            }
            $arrCentroid = ($kmeans.Clusters.Centroids)[$intCounterA]
            foreach ($intItemIndex in $hashtableClustersToItems.Item($intCounterA)) {
                #region Report Progress ########################################################
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

                # Centroid: $arrCentroid
                # Embeddings for this item: @($arrEmbeddings[$intItemIndex])
                # Distance: Measure-EuclideanDistance -Point1 $arrCentroid -Point2 @($arrEmbeddings[$intItemIndex])
                $doubleDistance = Measure-EuclideanDistance -Point1 $arrCentroid -Point2 @($arrEmbeddings[$intItemIndex])

                $psobject = New-Object PSCustomObject
                $psobject | Add-Member -MemberType NoteProperty -Name 'ItemNumber' -Value $intItemIndex
                $psobject | Add-Member -MemberType NoteProperty -Name 'DistanceFromCentroid' -Value $doubleDistance

                # Add the updated object to the list
                ($PSCustomObjectKMeansArtifacts.HashtableClustersToItemsAndDistances.Item($intCounterA)).Add($psobject)

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
        }
    } else {
        # PowerShell 5.0 or 5.1
        for ($intCounterA = 0; $intCounterA -lt $NumberOfClusters; $intCounterA++) {
            $arrCentroid = ($kmeans.Clusters.Centroids)[$intCounterA]
            foreach ($intItemIndex in $hashtableClustersToItems.Item($intCounterA)) {
                #region Report Progress ########################################################
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

                # Centroid: $arrCentroid
                # Embeddings for this item: @($arrEmbeddings[$intItemIndex])
                # Distance: Measure-EuclideanDistance -Point1 $arrCentroid -Point2 @($arrEmbeddings[$intItemIndex])
                $doubleDistance = Measure-EuclideanDistance -Point1 $arrCentroid -Point2 @($arrEmbeddings[$intItemIndex])

                $psobject = New-Object PSCustomObject
                $psobject | Add-Member -MemberType NoteProperty -Name 'ItemNumber' -Value $intItemIndex
                $psobject | Add-Member -MemberType NoteProperty -Name 'DistanceFromCentroid' -Value $doubleDistance

                # Add the updated object to the list
                [void](($PSCustomObjectKMeansArtifacts.HashtableClustersToItemsAndDistances.Item($intCounterA)).Add($psobject))

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
        }
    }
    #endregion Create Hashtable Including Euclidian Distance from Item to Centroid ########

    #region Create Hashtable of Silhouette Scores for Each Item #########################
    if ($boolCalculateExtendedStatistics -eq $true) {
        if ($NumberOfClusters -eq 1) {
            Write-Information ('1 cluster: Creating hashtable of silhouette scores for each item...')
        } else {
            Write-Information ([string]$NumberOfClusters + ' clusters: Creating hashtable of silhouette scores for each item...')
        }
        $hashtableItemsToSilhouetteScores = @{}

        $intCounterMax = $ReferenceToTwoDimensionalArrayOfEmbeddings.Value.Length

        #region Collect Stats/Objects Needed for Writing Progress ##########################
        $intProgressReportingFrequency = 1
        $intTotalItems = $intCounterMax
        $strProgressActivity = 'Performing k-means clustering'
        if ($NumberOfClusters -eq 1) {
            $strProgressStatus = 'Calculating silhouette scores for each item in the dataset (1 cluster)'
        } else {
            $strProgressStatus = 'Calculating silhouette scores for each item in the dataset (' + [string]$NumberOfClusters + ' clusters)'
        }
        $strProgressCurrentOperationPrefix = 'Processing item'
        $timedateStartOfLoop = Get-Date
        # Create a queue for storing lagging timestamps for ETA calculation
        $queueLaggingTimestamps = New-Object System.Collections.Queue
        $queueLaggingTimestamps.Enqueue($timedateStartOfLoop)
        #endregion Collect Stats/Objects Needed for Writing Progress ##########################

        for ($intCounterA = 0; $intCounterA -lt $intCounterMax; $intCounterA++) {
            $intClusterNumber = $arrClusterNumberAssignmentsForEachItem[$intCounterA]

            #region Report Progress ########################################################
            $intCurrentItemNumber = $intCounterA + 1 # Forward direction for loop
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

            # Get average distance to items in cluster (a)
            $doubleDistanceToItemsInCluster = [double]0
            $intCountOfItemsInCluster = 0
            $hashtableClustersToItems.Item($intClusterNumber) | Where-Object { $_ -ne $intCounterA } | ForEach-Object {
                $intItemIndex = $_
                $intCountOfItemsInCluster++

                # Calculate or look up Euclidean distance
                $intMinimumIndexNumber = [Math]::Min($intCounterA, $intItemIndex)
                $intMaximumIndexNumber = [Math]::Max($intCounterA, $intItemIndex)
                $boolCalculateEuclideanDistance = $true
                if ($ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.ContainsKey($intMinimumIndexNumber) -eq $true) {
                    if ($ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.Item($intMinimumIndexNumber).ContainsKey($intMaximumIndexNumber) -eq $true) {
                        $doubleEuclideanDistance = $ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.Item($intMinimumIndexNumber).Item($intMaximumIndexNumber)
                        $boolCalculateEuclideanDistance = $false
                    }
                }
                if ($boolCalculateEuclideanDistance -eq $true) {
                    $doubleEuclideanDistance = Measure-EuclideanDistance -Point1 @($arrEmbeddings[$intCounterA]) -Point2 @($arrEmbeddings[$intItemIndex])
                    if ($ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.ContainsKey($intMinimumIndexNumber) -eq $false) {
                        $ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.Add($intMinimumIndexNumber, @{})
                    }
                    $ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.Item($intMinimumIndexNumber).Add($intMaximumIndexNumber, $doubleEuclideanDistance)
                }

                # Add it to the total
                $doubleDistanceToItemsInCluster += $doubleEuclideanDistance
            }
            if ($intCountOfItemsInCluster -gt 0) {
                $doubleAverageDistanceToItemsInCluster = $doubleDistanceToItemsInCluster / $intCountOfItemsInCluster
            } else {
                $doubleAverageDistanceToItemsInCluster = [double]0
            }

            # Get minimum average distance to items in other clusters (b)
            $doubleMinimumAverageDistanceToItemsInOtherClusters = [double]::MaxValue
            for ($intCounterB = 0; $intCounterB -lt $NumberOfClusters; $intCounterB++) {
                if ($intCounterB -ne $intClusterNumber) {
                    $doubleDistanceToItemsInOtherCluster = [double]0
                    $intCountOfItemsInOtherCluster = 0
                    $hashtableClustersToItems.Item($intCounterB) | ForEach-Object {
                        $intItemIndex = $_
                        $intCountOfItemsInOtherCluster++

                        # Calculate or look up Euclidean distance
                        $intMinimumIndexNumber = [Math]::Min($intCounterA, $intItemIndex)
                        $intMaximumIndexNumber = [Math]::Max($intCounterA, $intItemIndex)
                        $boolCalculateEuclideanDistance = $true
                        if ($ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.ContainsKey($intMinimumIndexNumber) -eq $true) {
                            if ($ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.Item($intMinimumIndexNumber).ContainsKey($intMaximumIndexNumber) -eq $true) {
                                $doubleEuclideanDistance = $ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.Item($intMinimumIndexNumber).Item($intMaximumIndexNumber)
                                $boolCalculateEuclideanDistance = $false
                            }
                        }
                        if ($boolCalculateEuclideanDistance -eq $true) {
                            $doubleEuclideanDistance = Measure-EuclideanDistance -Point1 @($arrEmbeddings[$intCounterA]) -Point2 @($arrEmbeddings[$intItemIndex])
                            if ($ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.ContainsKey($intMinimumIndexNumber) -eq $false) {
                                $ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.Add($intMinimumIndexNumber, @{})
                            }
                            $ReferenceToHashtableOfEuclideanDistancesBetweenItems.Value.Item($intMinimumIndexNumber).Add($intMaximumIndexNumber, $doubleEuclideanDistance)
                        }

                        # Add it to the total
                        $doubleDistanceToItemsInOtherCluster += $doubleEuclideanDistance
                    }
                    if ($intCountOfItemsInOtherCluster -gt 0) {
                        $doubleAverageDistanceToItemsInOtherCluster = $doubleDistanceToItemsInOtherCluster / $intCountOfItemsInOtherCluster
                    } else {
                        $doubleAverageDistanceToItemsInOtherCluster = [double]0
                    }
                    if ($doubleAverageDistanceToItemsInOtherCluster -lt $doubleMinimumAverageDistanceToItemsInOtherClusters) {
                        $doubleMinimumAverageDistanceToItemsInOtherClusters = $doubleAverageDistanceToItemsInOtherCluster
                    }
                }
            }

            $doubleSilhouetteScore = ($doubleMinimumAverageDistanceToItemsInOtherClusters - $doubleAverageDistanceToItemsInCluster) / [Math]::Max($doubleMinimumAverageDistanceToItemsInOtherClusters, $doubleAverageDistanceToItemsInCluster)
            $hashtableItemsToSilhouetteScores.Add($intCounterA, $doubleSilhouetteScore)

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
    }
    #endregion Create Hashtable of Silhouette Scores for Each Item #########################

    #region Create Hashtable of Cluster Number to Averaged Distance from Each Item to its Centroid in that Cluster
    if ($boolCalculateExtendedStatistics -eq $true) {
        $hashtableClusterNumberToAverageDistanceFromEachItemToItsClusterCentroid = @{}
        for ($intCounterA = 0; $intCounterA -lt $NumberOfClusters; $intCounterA++) {
            $doubleTotalDistance = [double]0
            $intNumberOfItemsInCluster = $PSCustomObjectKMeansArtifacts.HashtableClustersToItemsAndDistances.Item($intCounterA).Count
            for ($intCounterB = 0; $intCounterB -lt $intNumberOfItemsInCluster; $intCounterB++) {
                $doubleTotalDistance += ($PSCustomObjectKMeansArtifacts.HashtableClustersToItemsAndDistances.Item($intCounterA))[$intCounterB].DistanceFromCentroid
            }
            $hashtableClusterNumberToAverageDistanceFromEachItemToItsClusterCentroid.Add($intCounterA, $doubleTotalDistance / $intNumberOfItemsInCluster)
        }
    }
    #endregion Create Hashtable of Cluster Number to Averaged Distance from Each Item to its Centroid in that Cluster

    #region Create Hashtable of Distances Between Each Centroid and All Other Centroids
    if ($boolCalculateExtendedStatistics -eq $true) {
        $hashtableClusterNumberToDistancesToOtherCentroids = @{}
        for ($intCounterA = 0; $intCounterA -lt ($NumberOfClusters - 1); $intCounterA++) {
            $hashtableClusterNumberToDistancesToOtherCentroids.Add($intCounterA, @{})
            for ($intCounterB = $intCounterA + 1; $intCounterB -lt $NumberOfClusters; $intCounterB++) {
                $doubleDistance = Measure-EuclideanDistance -Point1 ($kmeans.Clusters.Centroids)[$intCounterA] -Point2 ($kmeans.Clusters.Centroids)[$intCounterB]
                $hashtableClusterNumberToDistancesToOtherCentroids.Item($intCounterA).Add($intCounterB, $doubleDistance)
            }
        }
    }
    #endregion Create Hashtable of Distances Between Each Centroid and All Other Centroids

    #region Create Hashtable Of Average Distance for Each Cluster and Compute WCSS #
    if ($NumberOfClusters -eq 1) {
        Write-Information ('1 cluster: Creating hashtable of average distance for each cluster and computing within-cluster sum of squares...')
    } else {
        Write-Information ([string]$NumberOfClusters + ' clusters: Creating hashtable of average distance for each cluster and computing within-cluster sum of squares...')
    }
    $doubleWCSS = [double]0
    for ($intCounterA = 0; $intCounterA -lt $NumberOfClusters; $intCounterA++) {
        $doubleSum = [double]0
        foreach ($psobject in $PSCustomObjectKMeansArtifacts.HashtableClustersToItemsAndDistances.Item($intCounterA)) {
            $doubleWCSS += [math]::Pow($psobject.DistanceFromCentroid, 2)
            $doubleSum += $psobject.DistanceFromCentroid
        }
    }
    $PSCustomObjectKMeansArtifacts.WithinClusterSumOfSquares = $doubleWCSS
    #endregion Create Hashtable Of Average Distance for Each Cluster and Compute WCSS #

    #region Calculate Between Cluster Scatter
    if ($boolCalculateExtendedStatistics -eq $true) {
        $doubleBetweenClusterScatter = [double]0
        for ($intCounterA = 0; $intCounterA -lt $NumberOfClusters; $intCounterA++) {
            $doubleEuclideanDistance = Measure-EuclideanDistance -Point1 ($kmeans.Clusters.Centroids)[$intCounterA] -Point2 $ReferenceToCentroidOverWholeDataset.Value
            $intClusterSize = $PSCustomObjectKMeansArtifacts.HashtableClustersToItemsAndDistances.Item($intCounterA).Count
            $doubleBetweenClusterScatter += $intClusterSize * [math]::Pow($doubleEuclideanDistance, 2)
        }
    }
    #endregion Calculate Between Cluster Scatter

    #region Calculate Within Cluster Scatter
    if ($boolCalculateExtendedStatistics -eq $true) {
        $doubleWithinClusterScatter = $PSCustomObjectKMeansArtifacts.WithinClusterSumOfSquares
    }
    #endregion Calculate Within Cluster Scatter

    #region Compute Silhouette Score ###############################################
    if ($boolCalculateExtendedStatistics -eq $true) {
        if ($NumberOfClusters -eq 1) {
            Write-Information ('1 cluster: Computing silhouette score...')
        } else {
            Write-Information ([string]$NumberOfClusters + ' clusters: Computing silhouette score...')
        }
        if ($NumberOfClusters -gt 1) {
            $doubleSilhouetteScore = [double]0
            $intCounterMax = $ReferenceToTwoDimensionalArrayOfEmbeddings.Value.Length
            for ($intCounterA = 0; $intCounterA -lt $intCounterMax; $intCounterA++) {
                $doubleSilhouetteScore += $hashtableItemsToSilhouetteScores.Item($intCounterA)
            }
            $doubleSilhouetteScore /= $intCounterMax
            $PSCustomObjectKMeansArtifacts.SilhouetteScore = $doubleSilhouetteScore
        }
    }
    #endregion Compute Silhouette Score ###############################################

    #region Compute Davies-Bouldin Index ###########################################
    if ($boolCalculateExtendedStatistics -eq $true) {
        if ($NumberOfClusters -eq 1) {
            Write-Information ('1 cluster: Computing Davies-Bouldin index...')
        } else {
            Write-Information ([string]$NumberOfClusters + ' clusters: Computing Davies-Bouldin index...')
        }
        if ($NumberOfClusters -ge 2) {
            $doubleDBIndex = [double]0
            for ($intCounterA = 0; $intCounterA -lt $NumberOfClusters; $intCounterA++) {
                $doubleMax = [double]0
                for ($intCounterB = 0; $intCounterB -lt $NumberOfClusters; $intCounterB++) {
                    if ($intCounterA -ne $intCounterB) {
                        $intMinimumClusterNumber = [Math]::Min($intCounterA, $intCounterB)
                        $intMaximumClusterNumber = [Math]::Max($intCounterA, $intCounterB)
                        $doubleValue = ($hashtableClusterNumberToAverageDistanceFromEachItemToItsClusterCentroid.Item($intCounterA) + $hashtableClusterNumberToAverageDistanceFromEachItemToItsClusterCentroid.Item($intCounterB)) / $hashtableClusterNumberToDistancesToOtherCentroids.Item($intMinimumClusterNumber).Item($intMaximumClusterNumber)
                        if ($doubleValue -gt $doubleMax) {
                            $doubleMax = $doubleValue
                        }
                    }
                }
                $doubleDBIndex += $doubleMax
            }
            $doubleDBIndex /= $NumberOfClusters
            $PSCustomObjectKMeansArtifacts.DaviesBouldinIndex = $doubleDBIndex
        }
    }
    #endregion Compute Davies-Bouldin Index ###########################################

    #region Compute Calinski-Harabasz Index #########################################
    if ($boolCalculateExtendedStatistics -eq $true) {
        if ($NumberOfClusters -eq 1) {
            Write-Information ('1 cluster: Computing Calinski-Harabasz index...')
        } else {
            Write-Information ([string]$NumberOfClusters + ' clusters: Computing Calinski-Harabasz index...')
        }
        if ($NumberOfClusters -ge 2) {
            $doubleNumerator = $doubleBetweenClusterScatter / ($NumberOfClusters - 1)
            $doubleDenominator = $doubleWithinClusterScatter / (($ReferenceToTwoDimensionalArrayOfEmbeddings.Value.Length) - $NumberOfClusters)
            $doubleCHIndex = $doubleNumerator / $doubleDenominator
            $PSCustomObjectKMeansArtifacts.CalinskiHarabaszIndex = $doubleCHIndex
        }
    }
    #endregion Compute Calinski-Harabasz Index #########################################

    $ReferenceToHashtableOfNumberOfClustersToArtifacts.Value.Add($NumberOfClusters, $PSCustomObjectKMeansArtifacts)
}

# TODO: Figure out how to use Microsoft.ML instead of Accord, at least for PowerShell
# v6.0 and later. Microsoft.ML is actively maintained.

$versionPS = Get-PSVersion

#region Quit if PowerShell Version is Unsupported ##################################
if ($versionPS -lt [version]'5.0') {
    Write-Warning 'This script requires PowerShell v5.0 or higher. Please upgrade to PowerShell v5.0 or higher and try again.'
    return # Quit script
}
#endregion Quit if PowerShell Version is Unsupported ##################################

#region Check for required PowerShell Modules ######################################
if ([string]::IsNullOrEmpty($OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters) -eq $false) {
    # User has requested at least one Excel workbook; make sure ImportExcel is
    # installed
    $hashtableModuleNameToInstalledModules = @{}
    $hashtableModuleNameToInstalledModules.Add('ImportExcel', @())
    $refHashtableModuleNameToInstalledModules = [ref]$hashtableModuleNameToInstalledModules
    Get-PowerShellModuleUsingHashtable -ReferenceToHashtable $refHashtableModuleNameToInstalledModules

    $hashtableCustomNotInstalledMessageToModuleNames = @{}

    $refhashtableCustomNotInstalledMessageToModuleNames = [ref]$hashtableCustomNotInstalledMessageToModuleNames
    $boolResult = Test-PowerShellModuleInstalledUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ThrowErrorIfModuleNotInstalled -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames

    if ($boolResult -eq $false) {
        return # Quit script
    }
}
#endregion Check for required PowerShell Modules ######################################

#region Check for PowerShell module updates ########################################
if ([string]::IsNullOrEmpty($OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters) -eq $false) {
    # User has requested at least one Excel workbook; make sure ImportExcel is
    # up to date
    $boolDoNotCheckForModuleUpdates = $false
    if ($null -ne $DoNotCheckForModuleUpdates) {
        if ($DoNotCheckForModuleUpdates.IsPresent -eq $true) {
            $boolDoNotCheckForModuleUpdates = $true
        }
    }
    if ($boolDoNotCheckForModuleUpdates -eq $false) {
        Write-Verbose 'Checking for module updates...'
        $hashtableCustomNotUpToDateMessageToModuleNames = @{}

        $refhashtableCustomNotUpToDateMessageToModuleNames = [ref]$hashtableCustomNotUpToDateMessageToModuleNames
        $boolResult = Test-PowerShellModuleUpdatesAvailableUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ThrowErrorIfModuleNotInstalled -ThrowWarningIfModuleNotUpToDate -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames -ReferenceToHashtableOfCustomNotUpToDateMessages $refhashtableCustomNotUpToDateMessageToModuleNames
    }
}
#endregion Check for PowerShell module updates ########################################

#region Make sure the input file exists ############################################
if ((Test-Path -Path $InputCSVPath -PathType Leaf) -eq $false) {
    Write-Warning ('Input CSV file not found at: "' + $InputCSVPath + '"')
    return # Quit script
}
#endregion Make sure the input file exists ############################################

#region Make sure that nuget.org is registered as a package source #################
# (if not, throw a warning and quit)
$boolNuGetDotOrgRegisteredAsPackageSource = Test-NuGetDotOrgRegisteredAsPackageSource -ThrowWarningIfNuGetDotOrgNotRegistered
if ($boolNuGetDotOrgRegisteredAsPackageSource -eq $false) {
    return # Quit script
}
#endregion Make sure that nuget.org is registered as a package source #################

#region Make sure that required NuGet packages are installed #######################
$hashtablePackageNameToInstalledPackageMetadata = @{}

$hashtablePackageNameToInstalledPackageMetadata.Add('Accord', $null)
$hashtablePackageNameToInstalledPackageMetadata.Add('Accord.Math', $null)
$hashtablePackageNameToInstalledPackageMetadata.Add('Accord.Statistics', $null)
$hashtablePackageNameToInstalledPackageMetadata.Add('Accord.MachineLearning', $null)
$refHashtablePackageNameToInstalledPackages = [ref]$hashtablePackageNameToInstalledPackageMetadata
Get-PackagesUsingHashtable -ReferenceToHashtable $refHashtablePackageNameToInstalledPackages

$hashtablePackageNameToSkippingDependencies = @{}
$refHashtablePackageNameToSkippingDependencies = [ref]$hashtablePackageNameToSkippingDependencies

$hashtableCustomNotInstalledMessageToPackageNames = @{}
$strAccordMachineLearningNotInstalledMessage = 'Accord, Accord.Math, Accord.Statistics, and/or Accord.MachineLearning packages were not found. Please install the Accord.MachineLearning package and its dependencies and then try again.' + [System.Environment]::NewLine + 'You can install the Accord.MachineLearning package and its dependencies by running the following command:' + [System.Environment]::NewLine + 'Install-Package -ProviderName NuGet -Name ''Accord.MachineLearning'' -Force -Scope CurrentUser;' + [System.Environment]::NewLine + [System.Environment]::NewLine
$hashtableCustomNotInstalledMessageToPackageNames.Add($strAccordMachineLearningNotInstalledMessage, @('Accord', 'Accord.Math', 'Accord.Statistics', 'Accord.MachineLearning'))
$refhashtableCustomNotInstalledMessageToPackageNames = [ref]$hashtableCustomNotInstalledMessageToPackageNames

$boolResult = Test-PackageInstalledUsingHashtable -ReferenceToHashtableOfInstalledPackages $refHashtablePackageNameToInstalledPackages -ThrowWarningIfPackageNotInstalled -ReferenceToHashtableOfSkippingDependencies $refHashtablePackageNameToSkippingDependencies -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToPackageNames
if ($boolResult -eq $false) {
    return # Quit script
}
#endregion Make sure that required NuGet packages are installed #######################

#region Make sure the output files don't already exist #############################
# (and if it does, delete it and then verify that it's gone)
if ((Test-Path -Path $OutputCSVPath -PathType Leaf) -eq $true) {
    Remove-Item -Path $OutputCSVPath -Force -ErrorAction SilentlyContinue
    if ((Test-Path -Path $OutputCSVPath -PathType Leaf) -eq $true) {
        Write-Warning ('Output CSV file already exists and could not be deleted (the file may be in use): "' + $OutputCSVPath + '"')
        return
    }
}

if ([string]::IsNullOrEmpty($OutputCSVPathForAllClusterStatistics) -eq $false) {
    if ((Test-Path -Path $OutputCSVPathForAllClusterStatistics -PathType Leaf) -eq $true) {
        Remove-Item -Path $OutputCSVPathForAllClusterStatistics -Force -ErrorAction SilentlyContinue
        if ((Test-Path -Path $OutputCSVPathForAllClusterStatistics -PathType Leaf) -eq $true) {
            Write-Warning ('Output CSV file for cluster statistics already exists and could not be deleted (the file may be in use): "' + $OutputCSVPathForAllClusterStatistics + '"')
            return
        }
    }
}

if ([string]::IsNullOrEmpty($OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters) -eq $false) {
    if ((Test-Path -Path $OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters -PathType Leaf) -eq $true) {
        Remove-Item -Path $OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters -Force -ErrorAction SilentlyContinue
        if ((Test-Path -Path $OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters -PathType Leaf) -eq $true) {
            Write-Warning ('Output Excel workbook for all clusters already exists and could not be deleted (the file may be in use): "' + $OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters + '"')
            return
        }
    }
}
#endregion Make sure the output files don't already exist #############################

#region Import the CSV #############################################################
$arrInputCSV = @()
$arrInputCSV = @(Import-Csv $InputCSVPath)
if ($arrInputCSV.Count -eq 0) {
    Write-Warning ('Input CSV file is empty: "' + $InputCSVPath + '"')
    return
}
#endregion Import the CSV #############################################################

# Create a fixed-size array to store the embeddings
$arrEmbeddings = New-Object PSCustomObject[] $arrInputCSV.Count

#region Load Embeddings Into Arrays ################################################

#region Collect Stats/Objects Needed for Writing Progress ##########################
$intProgressReportingFrequency = 50
$intTotalItems = $arrInputCSV.Count
$strProgressActivity = 'Performing k-means clustering'
$strProgressStatus = 'Loading embeddings into arrays'
$strProgressCurrentOperationPrefix = 'Processing item'
$timedateStartOfLoop = Get-Date
# Create a queue for storing lagging timestamps for ETA calculation
$queueLaggingTimestamps = New-Object System.Collections.Queue
$queueLaggingTimestamps.Enqueue($timedateStartOfLoop)
#endregion Collect Stats/Objects Needed for Writing Progress ##########################

Write-Verbose ($strProgressStatus + '...')

for ($intRowIndex = 0; $intRowIndex -lt $arrInputCSV.Count; $intRowIndex++) {
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

    $refStringOfEmbeddings = [ref]((($arrInputCSV[$intRowIndex]).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $DataFieldNameContainingEmbeddings }).Value)
    $arrEmbeddings[$intRowIndex] = Split-StringOnLiteralString ($refStringOfEmbeddings.Value) ';'

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

#endregion Load Embeddings Into Arrays ################################################

#region Load the Accord.NET NuGet Package DLLs #####################################
$hashtablePackageNameToEffectiveDotNETVersions = @{}
$refHashtablePackageNameToEffectiveDotNETVersions = [ref]$hashtablePackageNameToEffectiveDotNETVersions
$hashtablePackageNameToDLLPaths = @{}
$refHashtablePackageNameToDLLPaths = [ref]$hashtablePackageNameToDLLPaths
Get-DLLPathsForPackagesUsingHashtable -ReferenceToHashtableOfInstalledPackages $refHashtablePackageNameToInstalledPackages -ReferenceToHashtableOfEffectiveDotNETVersions $refHashtablePackageNameToEffectiveDotNETVersions -ReferenceToHashtableOfDLLPaths $refHashtablePackageNameToDLLPaths

$arrNuGetPackages = @()
$arrNuGetPackages += 'Accord'
$arrNuGetPackages += 'Accord.Math'
$arrNuGetPackages += 'Accord.Statistics'
$arrNuGetPackages += 'Accord.MachineLearning'

$arrDLLPaths = @()

foreach ($strPackageName in $arrNuGetPackages) {
    if ($hashtablePackageNameToDLLPaths.Item($strPackageName).Count -gt 0) {
        foreach ($strDLLPath in $hashtablePackageNameToDLLPaths.Item($strPackageName)) {
            $arrDLLPaths += $strDLLPath

            # Load the .dll
            Write-Information ('Loading .dll: "' + $strDLLPath + '"')
            try {
                Add-Type -Path $strDLLPath
            } catch {
                $strMessage = 'Error loading .dll: "' + $strDLLPath + '"; the LoaderException(s) are: '
                $_.Exception.LoaderExceptions | ForEach-Object { $strMessage += $_.Message + '; ' }
                Write-Warning $strMessage
                return
            }
        }
    }
}
#endregion Load the Accord.NET NuGet Package DLLs #####################################

#region Determine upper and lower bounds for the number of clusters ################
if (($null -eq $NumberOfClusters) -or ($NumberOfClusters -le 0)) {
    $intMinNumberOfClusters = 1
    $intMaxNumberOfClusters = [int]([Math]::Ceiling([Math]::Sqrt($arrInputCSV.Count) * 1.2))
    if ($intMaxNumberOfClusters -gt $arrInputCSV.Count) {
        $intMaxNumberOfClusters = $arrInputCSV.Count
    }
} else {
    $intMinNumberOfClusters = $NumberOfClusters
    $intMaxNumberOfClusters = $NumberOfClusters
}
#endregion Determine upper and lower bounds for the number of clusters ################

#region Perform k-means clustering #################################################
$hashtableKToClusterArtifacts = @{}
$hashtableEuclideanDistancesBetweenItems = @{}
$arrCentroidOverWholeDataset = @()

$boolDoNotCalculateExtendedStatistics = $false
if ($null -ne $DoNotCalculateExtendedStatistics) {
    if ($DoNotCalculateExtendedStatistics.IsPresent -eq $true) {
        $boolDoNotCalculateExtendedStatistics = $true
    }
}

# Display warning/information messages if necessary
if ($boolDoNotCalculateExtendedStatistics -eq $true) {
    Write-Warning 'The script is running without calculating the silhouette score, Davies-Bouldin index, and Calinski-Harabasz index; this will significantly speed up execution, but may lead to a less-ideal selection for the number of clusters.'
    if (([string]::IsNullOrEmpty($OutputCSVPathForAllClusterStatistics) -eq $false) -and ([string]::IsNullOrEmpty($OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters) -eq $false)) {
        Write-Warning 'Additionally, the script is working in unsupervised mode; it will do its best to select the best number of clusters and output the results to a CSV file, but no statistics and no Excel workbook that contains information about alternative numbers of clusters will be created. Combined with the fact that the silhouette score, Davies-Bouldin index, and Calinski-Harabasz index are not being calculated, running unsupervised in this manner is not a recommended mode of execution.'
    }
} else {
    if (([string]::IsNullOrEmpty($OutputCSVPathForAllClusterStatistics) -eq $false) -and ([string]::IsNullOrEmpty($OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters) -eq $false)) {
        Write-Information 'The script is working in unsupervised mode; it will automatically select the best number of clusters and output the results to a CSV file, but no statistics and no Excel workbook that contains information about alternative numbers of clusters will be created.'
    }
}

if ($intMinNumberOfClusters -ne 1) {
    if ($boolDoNotCalculateExtendedStatistics -eq $false) {
        # We need to run it once for a single cluster to get the centroid for the whole
        # data set
        $hashtableKToClusterArtifactsTEMP = @{}
        Invoke-KMeansClusteringForSpecifiedNumberOfClusters -ReferenceToHashtableOfNumberOfClustersToArtifacts ([ref]$hashtableKToClusterArtifactsTEMP) -NumberOfClusters 1 -ReferenceToTwoDimensionalArrayOfEmbeddings ([ref]$arrEmbeddings) -ReferenceToHashtableOfEuclideanDistancesBetweenItems ([ref]$hashtableEuclideanDistancesBetweenItems) -ReferenceToCentroidOverWholeDataset ([ref]$arrCentroidOverWholeDataset)
    }

    # Now we can run it for the specified range of clusters
    if ($boolDoNotCalculateExtendedStatistics -eq $false) {
        for ($intK = $intMinNumberOfClusters; $intK -le $intMaxNumberOfClusters; $intK++) {
            Invoke-KMeansClusteringForSpecifiedNumberOfClusters -ReferenceToHashtableOfNumberOfClustersToArtifacts ([ref]$hashtableKToClusterArtifacts) -NumberOfClusters $intK -ReferenceToTwoDimensionalArrayOfEmbeddings ([ref]$arrEmbeddings) -ReferenceToHashtableOfEuclideanDistancesBetweenItems ([ref]$hashtableEuclideanDistancesBetweenItems) -ReferenceToCentroidOverWholeDataset ([ref]$arrCentroidOverWholeDataset)
        }
    } else {
        # $boolDoNotCalculateExtendedStatistics is $true
        for ($intK = $intMinNumberOfClusters; $intK -le $intMaxNumberOfClusters; $intK++) {
            Invoke-KMeansClusteringForSpecifiedNumberOfClusters -ReferenceToHashtableOfNumberOfClustersToArtifacts ([ref]$hashtableKToClusterArtifacts) -NumberOfClusters $intK -ReferenceToTwoDimensionalArrayOfEmbeddings ([ref]$arrEmbeddings) -ReferenceToHashtableOfEuclideanDistancesBetweenItems ([ref]$hashtableEuclideanDistancesBetweenItems) -ReferenceToCentroidOverWholeDataset ([ref]$arrCentroidOverWholeDataset) -DoNotCalculateExtendedStatistics
        }
    }
} else {
    if ($boolDoNotCalculateExtendedStatistics -eq $false) {
        for ($intK = $intMinNumberOfClusters; $intK -le $intMaxNumberOfClusters; $intK++) {
            Invoke-KMeansClusteringForSpecifiedNumberOfClusters -ReferenceToHashtableOfNumberOfClustersToArtifacts ([ref]$hashtableKToClusterArtifacts) -NumberOfClusters $intK -ReferenceToTwoDimensionalArrayOfEmbeddings ([ref]$arrEmbeddings) -ReferenceToHashtableOfEuclideanDistancesBetweenItems ([ref]$hashtableEuclideanDistancesBetweenItems) -ReferenceToCentroidOverWholeDataset ([ref]$arrCentroidOverWholeDataset)
        }
    } else {
        # $boolDoNotCalculateExtendedStatistics is $true
        for ($intK = $intMinNumberOfClusters; $intK -le $intMaxNumberOfClusters; $intK++) {
            Invoke-KMeansClusteringForSpecifiedNumberOfClusters -ReferenceToHashtableOfNumberOfClustersToArtifacts ([ref]$hashtableKToClusterArtifacts) -NumberOfClusters $intK -ReferenceToTwoDimensionalArrayOfEmbeddings ([ref]$arrEmbeddings) -ReferenceToHashtableOfEuclideanDistancesBetweenItems ([ref]$hashtableEuclideanDistancesBetweenItems) -ReferenceToCentroidOverWholeDataset ([ref]$arrCentroidOverWholeDataset) -DoNotCalculateExtendedStatistics
        }
    }
}
#endregion Perform k-means clustering #################################################

#region Calculate WCSS first and second derivative #################################
if ($boolDoNotCalculateExtendedStatistics -eq $false) {
    $arrKeyStatistics = $hashtableKToClusterArtifacts.Values | Select-Object -Property @('NumberOfClusters', 'WithinClusterSumOfSquares', 'SilhouetteScore', 'DaviesBouldinIndex', 'CalinskiHarabaszIndex') | Sort-Object -Property 'NumberOfClusters'
} else {
    $arrKeyStatistics = $hashtableKToClusterArtifacts.Values | Select-Object -Property @('NumberOfClusters', 'WithinClusterSumOfSquares') | Sort-Object -Property 'NumberOfClusters'
}
$intCounterAMax = $arrKeyStatistics.Count - 1
# Add WCSS first and second derivative properties to each object
for ($intCounterA = 0; $intCounterA -le $intCounterAMax; $intCounterA++) {
    $arrKeyStatistics[$intCounterA] | Add-Member -MemberType NoteProperty -Name 'WCSSFirstDerivative' -Value $null
    $arrKeyStatistics[$intCounterA] | Add-Member -MemberType NoteProperty -Name 'WCSSSecondDerivative' -Value $null
}

# Calculate the first derivative
$intCounterA = 0
if ($intCounterAMax -ge 1) {
    $doubleWCSSFirstDerivative = $arrKeyStatistics[$intCounterA + 1].WithinClusterSumOfSquares - $arrKeyStatistics[$intCounterA].WithinClusterSumOfSquares
    $arrKeyStatistics[$intCounterA].WCSSFirstDerivative = $doubleWCSSFirstDerivative
}
for ($intCounterA = 1; $intCounterA -le $intCounterAMax - 1; $intCounterA++) {
    $doubleWCSSFirstDerivative = ($arrKeyStatistics[$intCounterA + 1].WithinClusterSumOfSquares - $arrKeyStatistics[$intCounterA - 1].WithinClusterSumOfSquares) / 2
    $arrKeyStatistics[$intCounterA].WCSSFirstDerivative = $doubleWCSSFirstDerivative
}
$intCounterA = $intCounterAMax
if ($intCounterAMax -ge 1) {
    $doubleWCSSFirstDerivative = $arrKeyStatistics[$intCounterA].WithinClusterSumOfSquares - $arrKeyStatistics[$intCounterA - 1].WithinClusterSumOfSquares
    $arrKeyStatistics[$intCounterA].WCSSFirstDerivative = $doubleWCSSFirstDerivative
}

# Calculate the second derivative
$intCounterA = 0
if ($intCounterAMax -ge 1) {
    $doubleWCSSSecondDerivative = $arrKeyStatistics[$intCounterA + 1].WCSSFirstDerivative - $arrKeyStatistics[$intCounterA].WCSSFirstDerivative
    $arrKeyStatistics[$intCounterA].WCSSSecondDerivative = $doubleWCSSSecondDerivative
}
for ($intCounterA = 1; $intCounterA -le $intCounterAMax - 1; $intCounterA++) {
    $doubleWCSSSecondDerivative = ($arrKeyStatistics[$intCounterA + 1].WCSSFirstDerivative - $arrKeyStatistics[$intCounterA - 1].WCSSFirstDerivative) / 2
    $arrKeyStatistics[$intCounterA].WCSSSecondDerivative = $doubleWCSSSecondDerivative
}
$intCounterA = $intCounterAMax
if ($intCounterAMax -ge 1) {
    $doubleWCSSSecondDerivative = $arrKeyStatistics[$intCounterA].WCSSFirstDerivative - $arrKeyStatistics[$intCounterA - 1].WCSSFirstDerivative
    $arrKeyStatistics[$intCounterA].WCSSSecondDerivative = $doubleWCSSSecondDerivative
}
#endregion Calculate WCSS first and second derivative #################################

#region Rank the number of clusters to bias toward more clusters ###################
$boolBiasedNumberOfClusters = $false
$doubleMaxSihlouetteScore = $arrKeyStatistics | Sort-Object -Property 'SilhouetteScore' -Descending | Select-Object -First 1 -ExpandProperty 'SilhouetteScore'
if ($doubleMaxSihlouetteScore -lt 0.4) {
    Write-Warning ('The maximum silhouette score is less than 0.4 (' + [string]::Format('{0:0.00}', $doubleMaxSihlouetteScore) + '); the data may not be well-clustered.')
    # The data may not be well-clustered; rank the number of clusters to bias toward more clusters
    $boolBiasedNumberOfClusters = $true
    for ($intCounterA = 0; $intCounterA -le $intCounterAMax; $intCounterA++) {
        $arrKeyStatistics[$intCounterA] | Add-Member -MemberType NoteProperty -Name 'NumberOfClustersRank' -Value $null
    }

    $boolIdealPointFound = $false
    $intUpperExpectedNumberOfClusters = [int]([Math]::Ceiling([Math]::Sqrt($arrInputCSV.Count)))
    for ($intCounterA = $intCounterAMax; $intCounterA -ge 0; $intCounterA--) {
        if ($arrKeyStatistics[$intCounterA].NumberOfClusters -eq $intUpperExpectedNumberOfClusters) {
            $boolIdealPointFound = $true
            $intCounterB = $intCounterA
            break
        }
    }

    if ($boolIdealPointFound -eq $false) {
        Write-Warning 'Could not find the ideal maximum point for the number of clusters; using the last point instead.'
        $intCounterB = $intCounterAMax
    }

    $intCounterA = 1
    $arrKeyStatistics[$intCounterB].NumberOfClustersRank = $intCounterA
    $intLowestClusterProcessed = $intCounterB
    $intCounterA++

    for ($intCounterC = 0; ($intCounterB + $intCounterC + 1) -lt $arrKeyStatistics.Count; $intCounterC++) {
        if ($intCounterB - ($intCounterC * 3) - 1 -ge 0) {
            $arrKeyStatistics[$intCounterB - ($intCounterC * 3) - 1].NumberOfClustersRank = $intCounterA
            $intLowestClusterProcessed = $intCounterB - ($intCounterC * 3) - 1
            $intCounterA++
        }
        if ($intCounterB - ($intCounterC * 3) - 2 -ge 0) {
            $arrKeyStatistics[$intCounterB - ($intCounterC * 3) - 2].NumberOfClustersRank = $intCounterA
            $intLowestClusterProcessed = $intCounterB - ($intCounterC * 3) - 2
            $intCounterA++
        }
        if ($intCounterB - ($intCounterC * 3) - 3 -ge 0) {
            $arrKeyStatistics[$intCounterB - ($intCounterC * 3) - 3].NumberOfClustersRank = $intCounterA
            $intLowestClusterProcessed = $intCounterB - ($intCounterC * 3) - 3
            $intCounterA++
        }
        $arrKeyStatistics[$intCounterB + $intCounterC + 1].NumberOfClustersRank = $intCounterA
        $intCounterA++
    }

    for ($intCounterB = $intLowestClusterProcessed - 1; $intCounterB -ge 0; $intCounterB--) {
        $arrKeyStatistics[$intCounterB].NumberOfClustersRank = $intCounterA
        $intCounterA++
    }
}
#endregion Rank the number of clusters to bias toward more clusters ###################

#region Rank the WCSS to bias toward tighter clusters ##############################
for ($intCounterA = 0; $intCounterA -le $intCounterAMax; $intCounterA++) {
    $arrKeyStatistics[$intCounterA] | Add-Member -MemberType NoteProperty -Name 'WCSSRank' -Value $null
}

$intCounterA = 1

# Sort the array by WCSS
$arrKeyStatistics | Sort-Object -Property 'WithinClusterSumOfSquares' | ForEach-Object {
    $_.WCSSRank = $intCounterA
    $intCounterA++
}
#endregion Rank the WCSS to bias toward tighter clusters ##############################

#region Rank the WCSS second derivative to bias toward elbow point #################
for ($intCounterA = 0; $intCounterA -le $intCounterAMax; $intCounterA++) {
    $arrKeyStatistics[$intCounterA] | Add-Member -MemberType NoteProperty -Name 'WCSSSecondDerivativeRank' -Value $null
}

$intCounterA = 1

# Higher values of WCSSSecondDerivativeRank indicates it's more likely the elbow point
$arrKeyStatistics | Sort-Object -Property 'WCSSSecondDerivative' -Descending | ForEach-Object {
    if ($null -eq $_.WCSSSecondDerivative) {
        # First or second cluster - no second derivative
        $_.WCSSSecondDerivativeRank = $intCounterAMax
    } else {
        $_.WCSSSecondDerivativeRank = $intCounterA
        $intCounterA++
    }
}
#endregion Rank the WCSS second derivative to bias toward elbow point #################

#region Rank the Silhouette Score to bias toward higher values #####################
if ($boolDoNotCalculateExtendedStatistics -eq $false) {
    for ($intCounterA = 0; $intCounterA -le $intCounterAMax; $intCounterA++) {
        $arrKeyStatistics[$intCounterA] | Add-Member -MemberType NoteProperty -Name 'SilhouetteScoreRank' -Value $null
    }

    $intCounterA = 1

    # Higher values of SilhouetteScoreRank indicate better clustering
    $arrKeyStatistics | Sort-Object -Property 'SilhouetteScore' -Descending | ForEach-Object {
        if ($null -eq $_.SilhouetteScore) {
            $_.SilhouetteScoreRank = $intCounterAMax
        } else {
            $_.SilhouetteScoreRank = $intCounterA
            $intCounterA++
        }
    }
}
#endregion Rank the Silhouette Score to bias toward higher values #####################

#region Rank the Davies-Bouldin Index to bias toward lower values ##################
if ($boolDoNotCalculateExtendedStatistics -eq $false) {
    for ($intCounterA = 0; $intCounterA -le $intCounterAMax; $intCounterA++) {
        $arrKeyStatistics[$intCounterA] | Add-Member -MemberType NoteProperty -Name 'DaviesBouldinIndexRank' -Value $null
    }

    $intCounterA = 1

    # Lower values of DaviesBouldinIndexRank indicate better clustering
    $arrKeyStatistics | Sort-Object -Property 'DaviesBouldinIndex' | ForEach-Object {
        if ($null -eq $_.DaviesBouldinIndex) {
            $_.DaviesBouldinIndexRank = $intCounterAMax
        } else {
            $_.DaviesBouldinIndexRank = $intCounterA
            $intCounterA++
        }
    }
}
#endregion Rank the Davies-Bouldin Index to bias toward lower values ##################

#region Rank the Calinski-Harabasz Index to bias toward higher values ##############
if ($boolDoNotCalculateExtendedStatistics -eq $false) {
    for ($intCounterA = 0; $intCounterA -le $intCounterAMax; $intCounterA++) {
        $arrKeyStatistics[$intCounterA] | Add-Member -MemberType NoteProperty -Name 'CalinskiHarabaszIndexRank' -Value $null
    }

    $intCounterA = 1

    # Higher values of CalinskiHarabaszIndexRank indicate better clustering
    $arrKeyStatistics | Sort-Object -Property 'CalinskiHarabaszIndex' -Descending | ForEach-Object {
        if ($null -eq $_.CalinskiHarabaszIndex) {
            $_.CalinskiHarabaszIndexRank = $intCounterAMax
        } else {
            $_.CalinskiHarabaszIndexRank = $intCounterA
            $intCounterA++
        }
    }
}
#endregion Rank the Calinski-Harabasz Index to bias toward higher values ##############

#region Generate composite ranking #################################################
for ($intCounterA = 0; $intCounterA -le $intCounterAMax; $intCounterA++) {
    $doubleCompositeScore = [double]0

    $boolClusterSizeUsed = $false
    $boolWCSSUsed = $false
    $boolWCSSSecondDerivativeUsed = $false
    $boolSihouetteScoreUsed = $false
    $boolDaviesBouldinScoreUsed = $false
    $boolCalinskiHarabaszScoreUsed = $false

    if ($boolBiasedNumberOfClusters -eq $true) {
        if ($null -ne $arrKeyStatistics[$intCounterA].NumberOfClustersRank) {
            $boolClusterSizeUsed = $true
            $doubleCompositeScore += $WeightingFactorForClusterSize * $arrKeyStatistics[$intCounterA].NumberOfClustersRank
        }
    }

    if ($null -ne $arrKeyStatistics[$intCounterA].WCSSRank) {
        $boolWCSSUsed = $true
        $doubleCompositeScore += $WeightingFactorForWCSS * $arrKeyStatistics[$intCounterA].WCSSRank
    }

    if ($null -ne $arrKeyStatistics[$intCounterA].WCSSSecondDerivativeRank) {
        $boolWCSSSecondDerivativeUsed = $true
        $doubleCompositeScore += $WeightingFactorForWCSSSecondDerivative * $arrKeyStatistics[$intCounterA].WCSSSecondDerivativeRank
    }

    if ($boolDoNotCalculateExtendedStatistics -eq $false) {
        if ($null -ne $arrKeyStatistics[$intCounterA].SilhouetteScoreRank) {
            $boolSihouetteScoreUsed = $true
            $doubleCompositeScore += $WeightingFactorForSihouetteScore * $arrKeyStatistics[$intCounterA].SilhouetteScoreRank
        }

        if ($null -ne $arrKeyStatistics[$intCounterA].DaviesBouldinIndexRank) {
            $boolDaviesBouldinScoreUsed = $true
            $doubleCompositeScore += $WeightingFactorForDaviesBouldinScore * $arrKeyStatistics[$intCounterA].DaviesBouldinIndexRank
        }

        if ($null -ne $arrKeyStatistics[$intCounterA].CalinskiHarabaszIndexRank) {
            $boolCalinskiHarabaszScoreUsed = $true
            $doubleCompositeScore += $WeightingFactorForCalinskiHarabaszScore * $arrKeyStatistics[$intCounterA].CalinskiHarabaszIndexRank
        }
    }

    $doubleDivisor = [double]0

    if ($boolClusterSizeUsed -eq $true) {
        $doubleDivisor += $WeightingFactorForClusterSize
    }
    if ($boolWCSSUsed -eq $true) {
        $doubleDivisor += $WeightingFactorForWCSS
    }
    if ($boolWCSSSecondDerivativeUsed -eq $true) {
        $doubleDivisor += $WeightingFactorForWCSSSecondDerivative
    }
    if ($boolSihouetteScoreUsed -eq $true) {
        $doubleDivisor += $WeightingFactorForSihouetteScore
    }
    if ($boolDaviesBouldinScoreUsed -eq $true) {
        $doubleDivisor += $WeightingFactorForDaviesBouldinScore
    }
    if ($boolCalinskiHarabaszScoreUsed -eq $true) {
        $doubleDivisor += $WeightingFactorForCalinskiHarabaszScore
    }

    if ($doubleDivisor -gt 0) {
        $doubleCompositeScore /= $doubleDivisor
    } else {
        Write-Warning 'The divisor for the composite score was zero; the composite score will be set to zero.'
        $doubleCompositeScore = 0
    }

    $arrKeyStatistics[$intCounterA] | Add-Member -MemberType NoteProperty -Name 'CompositeRank' -Value $doubleCompositeScore
}
#endregion Generate composite ranking #################################################

#region Output statistics if requested #############################################
Write-Information 'Generating output CSV of statistics...'
if ([string]::IsNullOrEmpty($OutputCSVPathForAllClusterStatistics) -eq $false) {
    $arrKeyStatistics | Export-Csv -Path $OutputCSVPathForAllClusterStatistics -NoTypeInformation
}
#endregion Output statistics if requested #############################################

#region Output information about each k-means clustering operation (with varying numbers of clusters) if requested
Write-Information 'Generating output Excel workbook for cluster information...'
if ([string]::IsNullOrEmpty($OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters) -eq $false) {
    $arrNumberOfClusters = @($hashtableKToClusterArtifacts.Keys)
    $intCounterAMax = $arrNumberOfClusters.Count - 1
    for ($intCounterA = 0; $intCounterA -le $intCounterAMax; $intCounterA++) {
        $intNumberOfClusters = $arrNumberOfClusters[$intCounterA]

        if ($versionPS -ge ([version]'6.0')) {
            $listOutput = New-Object -TypeName 'System.Collections.Generic.List[PSCustomObject]'
        } else {
            # On Windows PowerShell (versions older than 6.x), we use an ArrayList instead
            # of a generic list
            # TODO: Fill in rationale for this
            #
            # Technically, in older versions of PowerShell, the type in the ArrayList will
            # be a PSObject; but that does not matter for our purposes.
            $listOutput = New-Object -TypeName 'System.Collections.ArrayList'
        }

        for ($intCounterB = 0; $intCounterB -lt $intNumberOfClusters; $intCounterB++) {
            # Cluster #: $intCounterB

            $arrSortedItems = $hashtableKToClusterArtifacts.Item($intNumberOfClusters).HashtableClustersToItemsAndDistances.Item($intCounterB) | Sort-Object -Property 'DistanceFromCentroid'
            $intMostRepresentativeItem = ($arrSortedItems[0]).ItemNumber
            $arrNMostRepresentativeItems = @($arrSortedItems | Select-Object -First $NSizeForMostRepresentativeDataPoints | ForEach-Object { $_.ItemNumber })
            $strNMostRepresentativeItems = $arrNMostRepresentativeItems -Join '; '
            $strItemsInCluster = @($arrSortedItems | ForEach-Object { $_.ItemNumber }) -Join '; '
            $psobject = New-Object -TypeName 'PSObject'
            $psobject | Add-Member -MemberType NoteProperty -Name 'MostRepresentativeItemIndex' -Value $intMostRepresentativeItem
            $psobject | Add-Member -MemberType NoteProperty -Name 'CountOfNMostRepresentativeItems' -Value ($arrNMostRepresentativeItems.Count)
            $psobject | Add-Member -MemberType NoteProperty -Name 'NMostRepresentativeItemIndices' -Value $strNMostRepresentativeItems
            $psobject | Add-Member -MemberType NoteProperty -Name 'CountOfItemsInCluster' -Value ($arrSortedItems.Count)
            $psobject | Add-Member -MemberType NoteProperty -Name 'ItemsInClusterIndices' -Value $strItemsInCluster

            # Add the cluster information to the list
            if ($versionPS -ge ([version]'6.0')) {
                $listOutput.Add($psobject)
            } else {
                # On Windows PowerShell (versions older than 6.x), we use an ArrayList instead
                # of a generic list
                # TODO: Fill in rationale for this
                #
                # Technically, in older versions of PowerShell, the type in the ArrayList will
                # be a PSObject; but that does not matter for our purposes.
                [void]($listOutput.Add($psobject))
            }
        }

        $listOutput |
            Sort-Object -Property @(@{ Expression = 'CountOfItemsInCluster'; Descending = $true }, @{ Expression = 'MostRepresentativeItem'; Descending = $false }) |
            Export-Excel -Path $OutputExcelWorkbookPathForClusterInformationForEachNumberOfClusters -WorksheetName ([string]$intNumberOfClusters + ' Clusters') -AutoFilter -FreezeTopRow -BoldTopRow
    }
}
#endregion Output information about each k-means clustering operation (with varying numbers of clusters) if requested

#region Select best number of clusters #############################################
$PSCustomObjectTopCluster = $arrKeyStatistics | Sort-Object -Property 'CompositeRank' | Select-Object -First 1
$intBestNumberOfClusters = $PSCustomObjectTopCluster.NumberOfClusters
Write-Information ('Best number of clusters: ' + $intBestNumberOfClusters)
#endregion Select best number of clusters #############################################

#region Generate output CSV for selected number of clusters ########################
Write-Information ('Generating output CSV for ' + $intBestNumberOfClusters + ' clusters')
$intNumberOfClusters = $intBestNumberOfClusters

if ($versionPS -ge ([version]'6.0')) {
    $listOutput = New-Object -TypeName 'System.Collections.Generic.List[PSCustomObject]'
} else {
    # On Windows PowerShell (versions older than 6.x), we use an ArrayList instead
    # of a generic list
    # TODO: Fill in rationale for this
    #
    # Technically, in older versions of PowerShell, the type in the ArrayList will
    # be a PSObject; but that does not matter for our purposes.
    $listOutput = New-Object -TypeName 'System.Collections.ArrayList'
}

for ($intCounterB = 0; $intCounterB -lt $intNumberOfClusters; $intCounterB++) {
    # Cluster #: $intCounterB

    $arrSortedItems = $hashtableKToClusterArtifacts.Item($intNumberOfClusters).HashtableClustersToItemsAndDistances.Item($intCounterB) | Sort-Object -Property 'DistanceFromCentroid'
    $intMostRepresentativeItem = ($arrSortedItems[0]).ItemNumber
    $arrNMostRepresentativeItems = @($arrSortedItems | Select-Object -First $NSizeForMostRepresentativeDataPoints | ForEach-Object { $_.ItemNumber })
    $strNMostRepresentativeItems = $arrNMostRepresentativeItems -Join '; '
    $strItemsInCluster = @($arrSortedItems | ForEach-Object { $_.ItemNumber }) -Join '; '
    $psobject = New-Object -TypeName 'PSObject'
    $psobject | Add-Member -MemberType NoteProperty -Name 'MostRepresentativeItemIndex' -Value $intMostRepresentativeItem
    $psobject | Add-Member -MemberType NoteProperty -Name 'CountOfNMostRepresentativeItems' -Value ($arrNMostRepresentativeItems.Count)
    $psobject | Add-Member -MemberType NoteProperty -Name 'NMostRepresentativeItemIndices' -Value $strNMostRepresentativeItems
    $psobject | Add-Member -MemberType NoteProperty -Name 'CountOfItemsInCluster' -Value ($arrSortedItems.Count)
    $psobject | Add-Member -MemberType NoteProperty -Name 'ItemsInClusterIndices' -Value $strItemsInCluster

    # Add the cluster information to the list
    if ($versionPS -ge ([version]'6.0')) {
        $listOutput.Add($psobject)
    } else {
        # On Windows PowerShell (versions older than 6.x), we use an ArrayList instead
        # of a generic list
        # TODO: Fill in rationale for this
        #
        # Technically, in older versions of PowerShell, the type in the ArrayList will
        # be a PSObject; but that does not matter for our purposes.
        [void]($listOutput.Add($psobject))
    }
}

$listOutput |
    Sort-Object -Property @(@{ Expression = 'CountOfItemsInCluster'; Descending = $true }, @{ Expression = 'MostRepresentativeItem'; Descending = $false }) |
    Export-Csv -Path $OutputCSVPath -NoTypeInformation
#endregion Generate output CSV for selected number of clusters ########################
