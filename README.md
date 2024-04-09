# AutoCategorizerPS

A series of scripts that perform zero-shot (untrained) data classification using AI.

## Contents

- [Contents](#contents)
- [Motivation](#motivation)
- [Prerequisites \& Setup](#prerequisites--setup)
  - [Getting OpenAI API Key](#getting-an-openai-api-key)
  - [Setup Azure Key Vault](#setup-azure-key-vault)
  - [Install Required Software](#install-required-software)
- [Usage](#usage)

## Motivation

Unstructured text data is inherently difficult to analyze.
And, when we think about use cases involving a large volume of data (comments on a survey, product or restaurant reviews, etc.), manually analyzing each text entry becomes unsustainable.

Instead, it's helpful to understand the "key themes" of the data in these data sets.
So, we utilize OpenAI's embeddings API, K-means Clustering, and ChatGPT to group similar data and surface the topic/category of each "cluster."

At the time of creation, data science activities such as K-means clustering are reserved for scripting languages such as Python and Matlab.
On the other hand, PowerShell, a more rich, object-oriented scripting language, has very few data science use cases.
While solving the goal, this project also aims to change the programming language bias to show that PowerShell is a capable platform for data science activities.

## Prerequisites & Setup

### Getting an OpenAI API Key

1. To start, navigate to <https://platform.openai.com/api-keys> and login with your OpenAI account you have configured
2. Once you are on the page, click "+ Create new secret key
"
3. Name the key something memorable, set the permissions on the key to full access, then create the key. You will then be presented with the cleartext OpenAI key for you to save for later. Copy this to notepad or somewhere where you can copy the key later. We'll cover storing this key in an Azure Key Vault in the next section.

### Setup Azure Key Vault (AKV) to store OpenAI API key

1. To start, you'll need to provision an Azure Key Vault (AKV) in your Azure subscription. Once created, you need to have at least read permissions over the secrets stored in the vault, or you can use more permissive permissions such as Key Vault Administrator. For the scope of this document and readme, we won't include the full details on how to do this, but there are plenty of good guides online on how to setup an AKV and the proper permissions to access secrets in the AKV
2. Once you have an AKV created, sign in to the Azure Portal with your account that has access to the AKV. Then navigate to "Secrets" and click "Generate/Import" at the top. Name the secret for example "my-openai-api-key", enter the OpenAI API key you saved earlier in the "Secret Value" box, and finally ensure the "Enabled" checkbox is selected then click "Create".
3. After you have added the OpenAI API key to the AKV Secrets, you'll need to copy a few additional pieces of information to run the Get-TextEmbeddingsUsingOpenAI.ps1 script. Below is what you need to gather:
    - The Entra ID tenant ID that your Azure tenant is connected to
        - To retrieve this, search for "Tenant Properties" in the Azure portal and click on it when it appears. Then copy the "Tenant ID" that is listed as this is the Entra ID tenant ID.
    - Azure Subscription ID
        - To retrieve this, navigate to your AKV where your OpenAI API key is stored, then click "Overview". Copy the Subscription ID that is listed on this page.
    - AKV Name
        - The display name of your AKV
    - AKV Secret Name
        - This is the name of the secret you created in step #2, i.e. "my-openai-api-key"

### Install Required Software

#### Required PowerShell Version

To run all scripts, PowerShell 5.1 is required.
However, it is recommended to install and use the latest version of PowerShell.

#### Required PowerShell Modules

Additionally, in the PowerShell version that you are using, install the following PowerShell modules:

- Az.Accounts
- Az.KeyVault
- ImportExcel
- Microsoft.PowerShell.SecretManagement
- Microsoft.PowerShell.SecretStore

If any of the required PowerShell modules are missing or out of date, the script(s) that use that module will notify you and provide the required installation command.

#### Required NuGet Packages

The K-means Clustering script requires that you have nuget.org registered as a package source.
The script will notify you if you do not.

nuget.org can be registered using the following command:

```powershell
[void](Register-PackageSource -Name NuGetOrg -Location https://api.nuget.org/v3/index.json -ProviderName NuGet);
```

Next, you must have the required NuGet packages loaded.
The required packages are:

- Accord
- Accord.MachineLearning
- Accord.Math
- Accord.Statistics

The script will notify you if you are missing any of the required packages.

If needed, you can install the required packages using the following command:

```powershell
Install-Package -ProviderName NuGet -Name 'Accord.MachineLearning' -Force -Scope CurrentUser
```

## Usage

The exact usage of these scripts will vary depending on the dataset.
However, we provide guidance below on running each script in the process.
For more information on specific syntax, each function has comment-based help -- so you may use Get-Help to get more information:

```powershell
Get-Help NameOfScript
```

### Convert-MicrosoftFormsExcelExportToQuestionResponseFormat.ps1

It's common to use Microsoft Forms for survey data.
However, Microsoft Forms survey results are not stored in the required one-line-per-question/answer format.

Therefore, we include a script to convert Microsoft Forms output to the required CSV format.

Example usage:

```powershell
$arrQuestions = @('Without mentioning the Client name or specific people, what made your most challenging project challenging?', 'Without mentioning the Client name or specific people, what did you like the most about your favorite project?', 'Imagine the worst project possible. What is it about the project that would make it the worst?')
& '.\Convert-MicrosoftFormsExcelExportToQuestionResponseFormat.ps1' -InputExcelFilePath .\MicrosoftFormsSurveyExport.xlsx -ArrayOfQuestions $arrQuestions -OutputCSVPath .\ConvertedSurveyExport.csv
```

### Add-ContextToDataset.ps1

If your dataset would benefit from additional context being added to it, we include a script that can facilitate it.

Example usage:

```powershell
& '.\Add-ContextToDataset.ps1' -InputCSVPath .\ConvertedSurveyExport.csv -TextBeforeFieldName1 'On an employee engagement survey, a question was asked: ' -FieldName1 'Question' -TextBeforeFieldName2 ' #### In response, the employee wrote the following comment: ' -FieldName2 'Response' -AdditionalContextDataFieldName 'AdditionalContext' -OutputCSVPath .\ConvertedSurveyExport-WithContext.csv
```

### Convert-DataToAnonymizeAndRemoveJargon.ps1

Next, if your dataset contains confidential data and/or jargon that is not generally understood outside of the context of your organization, you should perform a "find and replace" operation to scrub the dataset.

To do this, you must prepare two CSVs:

- One for case-sensitive replacements
- One for case-insensitive replacements

Each CSV must have two columns:

- Find
- Replace

In both cases, when the text in the "find" column is found, it will be replaced with the text in the "Replace" column.

Example usage:

```powershell
& '.\Convert-DataToAnonymizeAndRemoveJargon.ps1' -InputCSVPath .\ConvertedSurveyExport-WithContext.csv -CaseSensitiveReplacementKeywordsInputCSVPath '.\ContosoCaseSensitiveReplacements.csv' -CaseInsensitiveReplacementKeywordsInputCSVPath '.\ContosoCaseInsensitiveReplacements.csv' -DataFieldName 'AdditionalContext' -AnonymizedAndDeJargonizedDataFieldName 'AdditionalContext_Scrubbed' -OutputCSVPath .\ConvertedSurveyExport-Scrubbed.csv'
```

### Get-TextEmbeddingsUsingOpenAI.ps1

Once the text data is prepared, the next step is to submit it to OpenAI's embeddings API to retrieve the embeddings (numerical representations of the text data).

Example usage:

```powershell
& '.\Get-TextEmbeddingsUsingOpenAI.ps1' -InputCSVPath .\ConvertedSurveyExport-Scrubbed.csv' -DataFieldNameToEmbed 'AdditionalContext_Scrubbed' -OutputCSVPath .\ConvertedSurveyExport-WithEmbeddings.csv -EntraIdTenantId '00bdb152-4d83-4056-9dce-a1a9f0210908' -AzureSubscriptionId 'a59e5b39-14b7-40dc-8052-52c7baca6f81' -AzureKeyVaultName 'PowerShellSecrets' -SecretName 'OpenAIAPIKey'
```

### Invoke-KMeansClustering.ps1

With the embeddings retrieved, we can perform K-means Clustering to group similar text data together.

Example usage:

```powershell
& '.\Invoke-KMeansClustering.ps1' -InputCSVPath .\ConvertedSurveyExport-WithEmbeddings.csv -DataFieldNameContainingEmbeddings 'Embeddings' -NumberOfClusters 10 -OutputCSVPath .\ConvertedSurveyExport-ClusterMetadata.csv
```

### Get-TopicForEachCluster.ps1

Finally, now that we have similar data clustered together, we can use the "most representative" item(s) from each cluster to determine the topic/category/main theme of each cluster.

Example usage:

```powershell
& '.\Get-TopicForEachCluster.ps1' -ClusterMetadataInputCSVPath .\ConvertedSurveyExport-ClusterMetadata.csv -UnstructuredTextDataInputCSVPath .\ConvertedSurveyExport-WithEmbeddings.csv -UnstructuredTextDataFieldNameContainingTextData 'AdditionalContext_Scrubbed' -OutputCSVPath .\ConvertedSurveyExport-ClusterMetadata-WithTopics.csv -EntraIdTenantId '00bdb152-4d83-4056-9dce-a1a9f0210908' -AzureSubscriptionId 'a59e5b39-14b7-40dc-8052-52c7baca6f81' -AzureKeyVaultName 'PowerShellSecrets' -SecretName 'OpenAIAPIKey'
```
