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

### Setup Azure Key Vault

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
