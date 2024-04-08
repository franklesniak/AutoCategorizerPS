# Add-ContextToDataset.ps1
# Version: 1.0.20240407.0

<#
.SYNOPSIS
Takes input data related to a survey/Q&A and combines the question and answer fields to
give additional context.

.DESCRIPTION
The Add-ContextToDataset.ps1 script is useful for situations where multiple questions
are being asked, each of which has their own answer field. This script will combine the
question and answer fields into a single field, which can be useful for providing
additional context when analyzing the data.

.PARAMETER InputCSVPath
Specifies the path to the input CSV file containing the dataset to be updated.

.PARAMETER TextBeforeFieldName1
Specifies the text to be added before the first field name, e.g., "Question: ", or
"On an employee engagement survey, a question was asked: ". This text will be added to
the beginning of the first field's data to provide additional context.

.PARAMETER FieldName1
Specifies the name of the first field that will be combined to provide context, e.g.,
the "question" field.

.PARAMETER TextBeforeFieldName2
Specifies the text to be added after the first field name and before the second field
name, e.g., " #### Answer: ", or " #### The employee wrote the following comment: ".
This text will be added to the end of the first field's data and before the second
field's data to provide additional context.

.PARAMETER FieldName2
Specifies the name of the second field that will be combined to provide context, e.g.,
the "answer" field or "comment" field.

.PARAMETER TextAfterFieldName2
Specifies the text to be added after the second field name, e.g., " ####". This text will
be added to the end of the second field's data to provide additional context.

.PARAMETER AdditionalContextDataFieldName
Specifies the name of the field in the output CSV file that will contain dataset
including the additional context.

.PARAMETER OutputCSVPath
Specifies the path to the output CSV file that will contain the updated dataset.

.EXAMPLE
PS C:\> .\Add-ContextToDataset.ps1 -InputCSVPath 'C:\Users\jdoe\Documents\Contoso Employee Survey Comments Aug 2021.csv' -TextBeforeFieldName1 'On an employee engagement survey, a question was asked: ' -FieldName1 'Question-Scrubbed' -TextBeforeFieldName2 ' #### In response, the employee wrote the following comment: ' -FieldName2 'Comment-Scrubbed' -AdditionalContextDataFieldName 'AdditionalContext-Scrubbed' -OutputCSVPath 'C:\Users\jdoe\Documents\Contoso Employee Survey Comments Aug 2021 - With Additional Context.csv'

This example concatenates the "Question-Scrubbed" and "Comment-Scrubbed" fields from the
input CSV file 'C:\Users\jdoe\Documents\Contoso Employee Survey Comments Aug 2021.csv'
and prepends each of them with additional context. The resulting dataset is written to
the output CSV file 'C:\Users\jdoe\Documents\Contoso Employee Survey Comments Aug 2021
- With Additional Context.csv'.

In this specific example, the contents of the "Question-Scrubbed" field and the
"Comment-Scrubbed" field are concatenated into a new column
("AdditionalContext-Scrubbed"). Based on the supplied parameters, the resulting field
is structured like:

On an employee engagement survey, a question was asked: <Question-Scrubbed> #### In response, the employee wrote the following comment: <Comment-Scrubbed>

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
    [parameter(Mandatory = $true)][string]$InputCSVPath,
    [parameter(Mandatory = $false)][string]$TextBeforeFieldName1 = '',
    [parameter(Mandatory = $true)][string]$FieldName1,
    [parameter(Mandatory = $false)][string]$TextBeforeFieldName2 = '',
    [parameter(Mandatory = $true)][string]$FieldName2,
    [parameter(Mandatory = $false)][string]$TextAfterFieldName2 = '',
    [parameter(Mandatory = $true)][string]$AdditionalContextDataFieldName,
    [parameter(Mandatory = $true)][string]$OutputCSVPath
)

#region Functions ##################################################################
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
#endregion Functions ##################################################################

$versionPS = Get-PSVersion

# Make sure the input file exists
if ((Test-Path -Path $InputCSVPath -PathType Leaf) -eq $false) {
    Write-Warning ('Input CSV file not found at: "' + $InputCSVPath + '"')
    return
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

for ($intRowIndex = 0; $intRowIndex -lt $arrInputCSV.Count; $intRowIndex++) {
    $strNewData = ''
    $psobjectUpdated = $null

    $refPSObjectThis = [ref]($arrInputCSV[$intRowIndex])
    $strNewData = $TextBeforeFieldName1 + `
    (($refPSObjectThis.Value).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $FieldName1 }).Value + `
    $TextBeforeFieldName2 + `
    (($refPSObjectThis.Value).PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' -and $_.Name -eq $FieldName2 }).Value + `
    $TextAfterFieldName2

    # Create a copy of the source object and add the new data to it
    $psobjectUpdated = $null
    $boolResult = Copy-Object ([ref]$psobjectUpdated) $refPSObjectThis
    if ($boolResult -eq $true) {
        $psobjectUpdated | Add-Member -MemberType NoteProperty -Name $AdditionalContextDataFieldName -Value $strNewData

        # Add the updated object to the output list
        if ($versionPS -ge ([version]'6.0')) {
            $listPSCustomObjectOutput.Add($psobjectUpdated)
        } else {
            [void]($listPSCustomObjectOutput.Add($psobjectUpdated))
        }
    }
}

# Export the CSV
$listPSCustomObjectOutput | Export-Csv -Path $OutputCSVPath -NoTypeInformation
