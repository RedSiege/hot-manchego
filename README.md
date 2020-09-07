# Hot Manchego

Macro-Enabled Excel File Generator (.xlsm) using the EPPlus Library. 

## Usage

```
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /reference:EPPlus.dll hot-manchego.cs
hot-manchego.exe blank.xlsm vba.txt
```

Compile the C# program file along with the EPPlus DLL. Then call the hot-manchego exe file with two arguments: the first is a blank xlsm file and the second is a txt file with your macro in vba format.

## Introduction

In September 1, 2020, [NVISO published a blog post about Operation Epic Manchego](https://blog.nviso.eu/2020/09/01/epic-manchego-atypical-maldoc-delivery-brings-flurry-of-infostealers/#comments). A threat actor had been uploading Macro-Enabled Excel Files (xlsm) to VirusTotal with farily ordinary VBA macros. However, the  method they used to create the files helped them get past most A/V vendors. Instead of creating the malicious Excel files using Microsoft Office, like everyone does, they used a third-party library called EPPlus. When using EPPlus, the creation of the Excel document varied significantly enough that most A/V didn't catch a simple lolbas payload to get a beacon on a target machine. 

For more details about the Epic Manchego campaign and a detailed walkthrough of detection methods, please view [NVISO's post](https://blog.nviso.eu/2020/09/01/epic-manchego-atypical-maldoc-delivery-brings-flurry-of-infostealers/#comments). 

## About This Tool

Hot Manchego uses the EPPlus Library to create a Macro-Enabled Excel File. There are three files (plus the README) in this repository.

1. EPPlus.dll 
> This is the brains of the operation. The EPPlus library enables us to create the macro files.If you'd like to compile your own version of the EPPlus DLL provided in this repo, [the original source code repository is available here](https://github.com/JanKallman/EPPlus). We didn't make any modifications to the EPPlus Library for use in this tool. 

2. vba.txt
> This is just a sample vba file that pops calculator.

3. hot-manchego.cs
> The file was based off of Sample15.cs from the EPPlus project. This file drives the creation of the Macro-enabled Excel File. Once compiled, the exe takes two inputs: a blank xlsm file and a txt file with your vba.

## Detection

NVISO wrote some detection rules for these files. Please see their post.