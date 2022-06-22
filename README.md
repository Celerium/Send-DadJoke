# Send-DadJoke

The Send-DadJoke script sends a dad joke to a Teams channel using a Teams webhook connector URI.

Dad jokes are randomly selected from icanhazdadjoke.com

---

## Send-DadJoke

![Send-DadJoke](https://raw.githubusercontent.com/Celerium/Send-DadJoke/main/.github/Celerium-Send-DadJoke-Example001.png)

## Initial Setup & Running

1. Teams Channel > Connectors > Incoming Webhook
2. Give the Webhook a name & logo
    - Create the Webhook
4. Copy the URI
    - The URI is how you tell the script what teams channel to send posts to.

---

```posh
    .\Send-DadJoke.ps1 -TeamsURI 'https://outlook.office.com/webhook/123/123/123/.....'
```

Using the defined webhooks connector URI a random dad joke is sent to the webhooks Teams channel.

No output is displayed to the console.
Using the -Verbose option will give you a basic display output


## Help :blue_book:

  - Help info and a list of parameters can be found by running `Get-Help .\Send-DadJoke.ps1`, such as:

```posh
Get-Help .\Send-DadJoke.ps1
Get-Help .\Send-DadJoke.ps1 -Full
```

---