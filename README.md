# AutoCategorizerPS

A series of scripts that perform zero-shot (untrained) data classification using AI.

## Contents

- [Contents](#contents)
- [Motivation](#motivation)
- [Prerequisites \& Setup](#prerequisites--setup)
  - [Getting OpenAI API Key](#getting-an-openai-api-key)
  - [Setup Azure Key Vault](#setup-azure-key-vault)
  - [Install required software](#install-required-software)
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

### Install required software

## Usage
