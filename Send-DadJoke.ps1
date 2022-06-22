<#
.NOTES
    NAME: Send-DadJoke.ps1
    Type: PowerShell

        AUTHOR:  David Schulte
        DATE:    2022-06-20
        EMAIL:   celerium@celerium.org
        Updated:
        Date:

    VERSION HISTORY:
    0.1 - 2022-06-20 - Initial Release

    TODO:
    N\A

.SYNOPSIS
    Sends a dad joke to a Teams channel.

.DESCRIPTION
    The Send-DadJoke script sends a dad joke to a Teams channel using a Teams webhook connector URI.

    Dad jokes are randomly selected from icanhazdadjoke.com

    Unless the -Verbose parameter is used, no output is displayed.

.PARAMETER TeamsURI
    A string that defines where the Microsoft Teams connector URI sends information to.

.EXAMPLE
    .\Send-DadJoke.ps1 -TeamsURI 'https://outlook.office.com/webhook/123/123/123/.....'

    Using the defined webhooks connector URI a random dad joke is sent to the webhooks Teams channel.

    No output is displayed to the console.

.EXAMPLE
    .\Send-DadJoke.ps1 -TeamsURI 'https://outlook.office.com/webhook/123/123/123/.....' -Verbose

    Using the defined webhooks connector URI a random dad joke is sent to the webhooks Teams channel.

    Output is displayed to the console.

.INPUTS
    TeamsURI

.OUTPUTS
    Console, TXT

.LINK
    Celerium - https://www.celerium.org/
    Dad Jokes - icanhazdadjoke.com

#>

<############################################################################################
                                        Code
############################################################################################>
#Requires -Version 5.0

#Region  [ Parameters ]

[CmdletBinding()]
param(
        [Parameter(ValueFromPipeline = $true, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$TeamsURI
    )

#EndRegion  [ Parameters ]

Write-Verbose ''
Write-Verbose "START - $(Get-Date -Format yyyy-MM-dd-HH:mm)"
Write-Verbose ''
Write-Verbose " - (1/3) - $(Get-Date -Format MM-dd-HH:mm) - Gathering Dad Data"

#Region     [ Prerequisites ]

    $Log = "C:\Celerium\Logs\Send-DadJoke-Report"
    $TXTReport = "$Log\Send-DadJokeLog.txt"

#EndRegion  [ Prerequisites ]

#Region  [ Main Code ]

try {

    $DadSources = @(    'https://www.gamespot.com/a/uploads/original/1578/15789737/3367450-hulkhogan.jpg' ,
                        'https://www.incimages.com/uploaded_files/image/1920x1080/052312_Tools_of_Sales2_575x270-panoramic_17063.jpg' ,
                        'https://images.cdn.circlesix.co/image/1/640/0/uploads/posts/2017/01/8be3b3212fe4106bb6f381e838ac72ee.jpg' ,
                        'https://pyxis.nymag.com/v1/imgs/513/eea/21c0e7552a37c9fb9f0597652fa8f3b724-22-shaving-cream.rsquare.w700.jpg' ,
                        'https://images.jpost.com/image/upload/f_auto,fl_lossy/t_JM_ArticleMainImageFaceDetect/414771'
                    )
        $DadImage = Get-Random -InputObject $DadSources

    $DadFact = (Invoke-RestMethod -Uri 'https://icanhazdadjoke.com' -Headers @{ accept="application/json" } -ErrorAction Stop).joke

}
catch {
    Write-Error $_

    if ( (Test-Path -Path $Log -PathType Container) -eq $false ){
        New-Item -Path $Log -ItemType Directory > $null
    }

    (Get-Date -Format yyyy-MM-dd-HH:mm) + " - " + "[ Step (1/2) ]" + " - " + $_.Exception.Message | Out-File $TXTReport -Append -Encoding utf8

    exit
}

#EndRegion  [ Main Code ]

Write-Verbose " - (2/2) - $(Get-Date -Format MM-dd-HH:mm) - Sending Dad Data"

#Region     [ Teams Code ]

$JSONBody = @"
{
    "type":"message",
    "attachments":[
    {
        "contentType":"application/vnd.microsoft.card.adaptive",
        "contentUrl":null,
        "content":{
            "$('$schema')":"http://adaptivecards.io/schemas/adaptive-card.json",
            "type":"AdaptiveCard",
            "version":"1.4",
            "body": [
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Dad Jokes & Puns",
                            "weight": "bolder",
                            "size": "Large"
                        },
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": "auto",
                                    "items": [
                                        {
                                            "type": "Image",
                                            "url": "$DadImage",
                                            "altText": "Dad Stuff",
                                            "size": "medium",
                                            "style": "default"
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": "stretch",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Hi Teams I'm Dad!",
                                            "weight": "bolder",
                                            "wrap": true
                                        },
                                        {
                                            "type": "TextBlock",
                                            "spacing": "none",
                                            "text": "No your other left",
                                            "size": "Small",
                                            "isSubtle": true,
                                            "wrap": true
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "$DadFact",
                            "wrap": true
                        }
                    ]
                }
            ],
                "msTeams": {
                    "width": "Full"
                }
            }
        }
    ]
}
"@

try {

    Invoke-RestMethod -Uri $TeamsURI -Method Post -ContentType 'application/json' -Body $JsonBody -ErrorAction Stop > $null

}
catch {
    Write-Error $_

    if ( (Test-Path -Path $Log -PathType Container) -eq $false ){
        New-Item -Path $Log -ItemType Directory > $null
    }

    (Get-Date -Format yyyy-MM-dd-HH:mm) + " - " + "[ Step (2/2) ]" + " - " + $_.Exception.Message | Out-File $TXTReport -Append -Encoding utf8

    exit
}

#EndRegion  [ Teams Code ]

Write-Verbose ''
Write-Verbose "End - $(Get-Date -Format yyyy-MM-dd-HH:mm)"
Write-Verbose ''