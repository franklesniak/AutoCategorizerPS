# Invoke-KMeansClustering.ps1
# Version: 1.0.20240409.0

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
Specifies the number of clusters that you want the function to use. The default value is
the square-root of $NumberOfDataPoints, rounded up, where $NumberOfDataPoints is the
number of rows in the input CSV file.

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
    [Parameter(Mandatory = $false)][int]$NSizeForMostRepresentativeDataPoints = 5,
    [Parameter(Mandatory = $false)][int]$NumberOfClusters
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
    Write-Debug ('Checking for registered package sources (Get-PackageSource)...')
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
        Write-Debug ('Checking for ' + $arrPackagesToGet[$intCounter] + ' software package...')
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

# TODO: Figure out how to use Microsoft.ML instead of Accord, at least for PowerShell
# v6.0 and later. Microsoft.ML is actively maintained.

$versionPS = Get-PSVersion

#region Quit if PowerShell Version is Unsupported ##################################
if ($versionPS -lt [version]'5.0') {
    Write-Warning 'This script requires PowerShell v5.0 or higher. Please upgrade to PowerShell v5.0 or higher and try again.'
    return # Quit script
}
#endregion Quit if PowerShell Version is Unsupported ##################################

# Make sure the input file exists
if ((Test-Path -Path $InputCSVPath -PathType Leaf) -eq $false) {
    Write-Warning ('Input CSV file not found at: "' + $InputCSVPath + '"')
    return # Quit script
}

# Make sure that nuget.org is registered as a package source; if not, throw a warning and quit
$boolNuGetDotOrgRegisteredAsPackageSource = Test-NuGetDotOrgRegisteredAsPackageSource -ThrowWarningIfNuGetDotOrgNotRegistered
if ($boolNuGetDotOrgRegisteredAsPackageSource -eq $false) {
    return # Quit script
}

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

# Make sure the output file doesn't already exist and if it does, delete it and then
# verify that it's gone
if ((Test-Path -Path $OutputCSVPath -PathType Leaf) -eq $true) {
    Remove-Item -Path $OutputCSVPath -Force
    if ((Test-Path -Path $OutputCSVPath -PathType Leaf) -eq $true) {
        Write-Warning ('Output CSV file already exists and could not be deleted (the file may be in use): "' + $OutputCSVPath + '"')
        return
    }
}

# Import the CSV
$arrInputCSV = @()
$arrInputCSV = @(Import-Csv $InputCSVPath)
if ($arrInputCSV.Count -eq 0) {
    Write-Warning ('Input CSV file is empty: "' + $InputCSVPath + '"')
    return
}

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
            Write-Debug ('Loading .dll: "' + $strDLLPath + '"')
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

# TODO: Dynamically set the number of clusters
if (($null -eq $NumberOfClusters) -or ($NumberOfClusters -le 0)) {
    $intNumberOfClusters = [int]([Math]::Ceiling([Math]::Sqrt($arrInputCSV.Count)))
} else {
    $intNumberOfClusters = $NumberOfClusters
}

$kmeans = New-Object -TypeName 'Accord.MachineLearning.KMeans' -ArgumentList @($intNumberOfClusters)
[void]($kmeans.Learn($arrEmbeddings))
$arrClusterNumberAssignmentsForEachItem = $kmeans.Clusters.Decide($arrEmbeddings)

#region Create Hashtable for Efficient Lookup of Cluster # to Associated Items #####
Write-Debug ('Creating hashtable for efficient lookup of cluster # to associated items...')
# Create a hashtable for easier lookup of cluster number to comment index number
$hashtableClustersToItems = @{}
if ($versionPS -ge ([version]'6.0')) {
    for ($intCounterA = 0; $intCounterA -lt $intNumberOfClusters; $intCounterA++) {
        $hashtableClustersToItems.Add($intCounterA, (New-Object -TypeName 'System.Collections.Generic.List[PSCustomObject]'))
    }
} else {
    # On Windows PowerShell (versions older than 6.x), we use an ArrayList instead
    # of a generic list
    # TODO: Fill in rationale for this
    #
    # Technically, in older versions of PowerShell, the type in the ArrayList will
    # be a PSObject; but that does not matter for our purposes.
    for ($intCounterA = 0; $intCounterA -lt $intNumberOfClusters; $intCounterA++) {
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
Write-Debug ('Creating hashtable including Euclidian distance from each item to its cluster centroid...')
$hashtableClustersToItemsAndDistances = @{}
if ($versionPS -ge ([version]'6.0')) {
    for ($intCounterA = 0; $intCounterA -lt $intNumberOfClusters; $intCounterA++) {
        $hashtableClustersToItemsAndDistances.Add($intCounterA, (New-Object -TypeName 'System.Collections.Generic.List[PSCustomObject]'))
    }
} else {
    # On Windows PowerShell (versions older than 6.x), we use an ArrayList instead
    # of a generic list
    # TODO: Fill in rationale for this
    #
    # Technically, in older versions of PowerShell, the type in the ArrayList will
    # be a PSObject; but that does not matter for our purposes.
    for ($intCounterA = 0; $intCounterA -lt $intNumberOfClusters; $intCounterA++) {
        $hashtableClustersToItemsAndDistances.Add($intCounterA, (New-Object -TypeName 'System.Collections.ArrayList'))
    }
}

#region Collect Stats/Objects Needed for Writing Progress ##########################
$intProgressReportingFrequency = 50
$intTotalItems = $arrInputCSV.Count
$strProgressActivity = 'Performing k-means clustering'
$strProgressStatus = 'Calculating distances from items to cluster centroids'
$strProgressCurrentOperationPrefix = 'Processing item'
$timedateStartOfLoop = Get-Date
# Create a queue for storing lagging timestamps for ETA calculation
$queueLaggingTimestamps = New-Object System.Collections.Queue
$queueLaggingTimestamps.Enqueue($timedateStartOfLoop)
#endregion Collect Stats/Objects Needed for Writing Progress ##########################

$intCounterLoop = 0
if ($versionPS -ge ([version]'6.0')) {
    # PowerShell v6.0 or newer
    for ($intCounterA = 0; $intCounterA -lt $intNumberOfClusters; $intCounterA++) {
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
            ($hashtableClustersToItemsAndDistances.Item($intCounterA)).Add($psobject)

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
    for ($intCounterA = 0; $intCounterA -lt $intNumberOfClusters; $intCounterA++) {
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
            [void](($hashtableClustersToItemsAndDistances.Item($intCounterA)).Add($psobject))

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

#region Generate Output CSV ########################################################
Write-Debug 'Generating output CSV...'
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

for ($intCounterA = 0; $intCounterA -lt $intNumberOfClusters; $intCounterA++) {
    # Cluster #: $intCounterA

    $arrSortedItems = $hashtableClustersToItemsAndDistances.Item($intCounterA) | Sort-Object -Property 'DistanceFromCentroid'
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
#endregion Generate Output CSV ########################################################
