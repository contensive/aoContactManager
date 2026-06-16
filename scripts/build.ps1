#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]   $LocalDeployTarget  = '',
    [hashtable]$RemoteDeployTarget = $null
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\..\Contensive5\scripts\build-addon-collection.psm1') -Force

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path

Invoke-ContensiveBuild `
    -CollectionName    'aoContactManager' `
    -CollectionPath    "$projectRoot\collections\aoContactManager" `
    -SolutionPath      "$projectRoot\source\ContactManager.sln" `
    -BinPath           "$projectRoot\source\ContactManager\bin\Release\netstandard2.0" `
    -DeploymentRoot    'C:\Deployments\aoContactManager' `
    -CleanFolders      @(
                           "$projectRoot\source\ContactManager\bin"
                           "$projectRoot\source\ContactManager\obj"
                       ) `
    -UiPath            "$projectRoot\ui" `
    -LocalDeployTarget  $LocalDeployTarget `
    -RemoteDeployTarget $RemoteDeployTarget
