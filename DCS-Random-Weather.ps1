param (
    [Parameter(Position=0,mandatory=$true)]
    [string]$MizPath,
    [switch]$Discord
)

#region Fonctions

    function Unzip_File ($ZipFile,$DestFolder) {
        Set-Alias sz $ScriptHT.SevenZip
        $command = "sz e -y `"$ZipFile`" `"-o$DestFolder`" mission theatre"
        $resultUnzip = Invoke-Expression $command
        if ($resultUnzip -contains "Everything is Ok") { $resultUnzip = $true } else { $resultUnzip = $false }
        return $resultUnzip
    }

    function Zip_File ($ZipFile,$DestFolder) {
        Set-Alias sz $ScriptHT.SevenZip
        $command = "sz u `"$ZipFile`" `"$DestFolder\mission`""
        $resultUnzip = Invoke-Expression $command
        if ($resultUnzip -contains "Everything is Ok") { $resultUnzip = $true } else { $resultUnzip = $false }
        return $resultUnzip
    }

    function WindsInverter ([int]$Wind) {
        $test = $Wind + 180
        if ($test -ge 360) { return $Wind - 180 }
        else { return $test }
    }

    function SpeedKts ([int]$Speed) {
        return [math]::Round($Speed*1.944)
    }

    function Random_Date {
        $dateMin = get-date -year 2020 -month 1 -day 1
        $dateMax = get-date -year 2021 -month 1 -day 1 
        return New-Object DateTime (Get-Random -min $dateMin.ticks -max $dateMax.ticks)
    }

    function Random_Temp ($Temp) {
        [int]$Temp = $Temp
        $rand = Get-Random -Minimum -5 -Maximum 6
        return ($Temp + $rand).ToString()
    }

    function Get_CloudFormation ($Month) {
        $ProbaArray = $ScriptHT.Config.Probabilities.$MonthConfig.$($ScriptHT.Theatre)
        $rand = Get-Random -Maximum 101
        # Clear
        if (($rand -ge 0) -and ($rand -le $ProbaArray[0])) {
            $FormationType = $ScriptHT.Config.CloudFormations.Clear
        }
        # Light
        elseif (($rand -gt $ProbaArray[0]) -and ($rand -le $ProbaArray[1])) {
            $FormationType = $ScriptHT.Config.CloudFormations.Light
        }
        # Scattered
        elseif (($rand -gt $ProbaArray[1]) -and ($rand -le $ProbaArray[2])) {
            $FormationType = $ScriptHT.Config.CloudFormations.Scattered
        }
        # Broken
        elseif (($rand -gt $ProbaArray[2]) -and ($rand -le $ProbaArray[3])) {
            $FormationType = $ScriptHT.Config.CloudFormations.Broken
        }
        # Overcast ou Rain
        else {
            $rand = Get-Random -Maximum 101
            if (($rand -ge 0) -and ($rand -le $ProbaArray[4])) {
                $FormationType = $ScriptHT.Config.CloudFormations.Rain
            }
            else {
                $FormationType = $ScriptHT.Config.CloudFormations.Overcast
            }
        }
        # Random sur le Preset correspondant à la CloudFormation
        $rand = Get-Random -Maximum $FormationType.Count
        return $FormationType[$rand].ToString()
    }

    function SendDiscord ([switch]$Fail) {
        <# References
            https://www.gngrninja.com/script-ninja/2018/3/10/using-discord-webhooks-with-powershell
            https://www.gngrninja.com/script-ninja/2018/3/17/using-discord-webhooks-and-embeds-with-powershell-part-2
            https://discord.com/developers/docs/resources/webhook#execute-webhook
        #>

        $embedArray = New-Object System.Collections.ArrayList
        $fields = New-Object System.Collections.ArrayList

        if ($Fail) {       
            $payload = [PSCustomObject]@{
                content = "Problème dans la génération de la mission $($ScriptHT.MizFile)"
                username = $ScriptHT.Config.Discord.ServerName
            }
            Invoke-RestMethod -Uri $ScriptHT.DiscordLink -Method Post -Body ($payload | ConvertTo-Json) -ContentType 'Application/Json; Charset=utf-8'
        }
        else {
            [void]$fields.Add([PSCustomObject]@{
                name = "Vent à 0m"
                value = "$(SpeedKts $RandomHT.WindSpeedGround) kts du $($RandomHT.WindDirGround)"
                inline = $true
            })
            [void]$fields.Add([PSCustomObject]@{
                name = "2000m"
                value = "$(SpeedKts $RandomHT.WindSpeed2K) kts du $($RandomHT.WindDir2000)"
                inline = $true
            })
            [void]$fields.Add([PSCustomObject]@{
                name = "8000m"
                value = "$(SpeedKts $RandomHT.WindSpeed8K) kts du $($RandomHT.WindDir8000)"
                inline = $true
            })
            [void]$fields.Add([PSCustomObject]@{
                name = "Turbulences"
                value = "$($(SpeedKts $RandomHT.Turbulence)/10) kts"
                inline = $false
            })
            [void]$fields.Add([PSCustomObject]@{
                name = "Nuages"
                value = $ScriptHT.Config.CloudPresets.$($RandomHT.Preset).Description
                inline = $false
            })
            if ($RandomHT.Preset -ne "Preset0") {
                [void]$fields.Add([PSCustomObject]@{
                    name = "METAR"
                    value = $ScriptHT.Config.CloudPresets.$($RandomHT.Preset).Metar
                    inline = $false
                })
            }
            $image = [PSCustomObject]@{url = $ScriptHT.Config.CloudPresets.$($RandomHT.Preset).Image}
            [void]$embedArray.Add([PSCustomObject]@{
                color = $ScriptHT.Config.Discord.$($ScriptHT.Theatre).ThumbColor
                title = "Météo du jour à $($ScriptHT.Config.Discord.$($ScriptHT.Theatre).City)"
                description = "$($ScriptHT.Config.Discord.$($ScriptHT.Theatre).Flag) Nous sommes le $($RandomHT.Day) $($ScriptHT.Config.CultureMonths.$($RandomHT.Month)) $($RandomHT.Year), il fait $($($RandomHT.Temperature))°C"
                fields = $fields
                image = $image
            })
            $payload = [PSCustomObject]@{
                embeds = $embedArray
                username = $ScriptHT.Config.Discord.BotName
                avatar_url = $ScriptHT.Config.Discord.BotAvatar
            }
            Invoke-RestMethod -Uri $ScriptHT.DiscordLink -Body ($payload | ConvertTo-Json -Depth 4) -Method Post -ContentType 'Application/Json; Charset=utf-8'
        }

    }

#endregion

#region Initialisations

    # Définition des variables
    $RandomHT = [hashtable]::Synchronized(@{})
    $ScriptHT = [hashtable]::Synchronized(@{})
    $ScriptHT.Config = Get-Content "$PSScriptRoot\DCS-Random-Weather.json" | ConvertFrom-Json
    $ScriptHT.MizPath = $MizPath
    $ScriptHT.MizFile = Split-Path $MizPath -Leaf
    $ScriptHT.extractFolder = "$PSScriptRoot\extractFolder"
    $ScriptHT.MissionNew = New-Object System.Collections.ArrayList
    $ScriptHT.SevenZip = "$($env:ProgramFiles)\7-Zip\7z.exe"
    $ScriptHT.DiscordLink = Get-Content "$PSScriptRoot\Discord.api"

    # Création du répertoire d'extraction si inexistant
    if (!(Test-Path $ScriptHT.extractFolder)) {
        New-Item -Path $PSScriptRoot -Name $(Split-Path $ScriptHT.extractFolder -Leaf) -ItemType "directory"
    }
    # S'il existe déjà, on le vide
    else {
        Remove-Item "$($ScriptHT.extractFolder)\*" -Recurse
    }

    # Unzip Miz
    $unzipFile = Unzip_File -ZipFile $ScriptHT.MizPath -DestFolder $ScriptHT.extractFolder

    # Récupération des informations du Miz
    $ScriptHT.Mission = Get-Content "$($ScriptHT.extractFolder)\mission"
    $ScriptHT.Theatre = Get-Content "$($ScriptHT.extractFolder)\theatre"

    Write-Host "Traitement de la mission $($ScriptHT.MizFile)" -ForegroundColor "Yellow"
    Write-Host "Map $($ScriptHT.Theatre)" -ForegroundColor "Blue"

#endregion

if ($unzipFile) {

    #region Randomization

        # Date
        $RandomHT.Date  = Random_Date
        $RandomHT.Year  = $RandomHT.Date.Year.ToString()
        $RandomHT.Month = $RandomHT.Date.Month.ToString()
        $RandomHT.Day   = $RandomHT.Date.Day.ToString()
        $MonthConfig = "Month$($RandomHT.Month)"

        # Winds Speed & Direction
        $RandomHT.WindSpeedGround   = (Get-Random -Minimum 2 -Maximum 9).ToString() # Beaufort 2-4
        $RandomHT.WindDirGround     = (Get-Random -Maximum 360).ToString()
        $RandomHT.WindSpeed2K       = (Get-Random -Minimum 2 -Maximum 15).ToString() # Beaufort 2-6
        $RandomHT.WindDir2000       = (Get-Random -Maximum 360).ToString()
        $RandomHT.WindSpeed8K       = (Get-Random -Minimum 2 -Maximum 15).ToString() # Beaufort 2-6
        $RandomHT.WindDir8000       = (Get-Random -Maximum 360).ToString()

        # Turbulences
        $RandomHT.Turbulence = (Get-Random -Minimum 5 -Maximum 21).ToString()

        # Temperature
        $RandomHT.Temperature = Random_Temp $ScriptHT.Config.Temperatures.$MonthConfig.$($ScriptHT.Theatre)

        # Preset Cloud
        $Preset = "Preset$(Get_CloudFormation $MonthConfig)"
        $RandomHT.Preset = $Preset
        $CloudsPreset = $ScriptHT.Config.CloudPresets.$Preset.Name
        $Cloudsbase = $ScriptHT.Config.CloudPresets.$Preset.Base

        # Informations
        Write-Host "Date : $($RandomHT.Date)"
        Write-Host "Temperature : $($RandomHT.Temperature)"
        Write-Host "Winds : $($RandomHT.WindSpeedGround) m/s @ $($RandomHT.WindDirGround) - $($RandomHT.WindSpeed2K) m/s @ $($RandomHT.WindDir2000) - $($RandomHT.WindSpeed8K) m/s @ $($RandomHT.WindDir8000)"
        Write-Host "Turbulences : $($RandomHT.Turbulence)"
        Write-Host "Cloud : $CloudsPreset"

    #endregion

    #region Traitement

        for ($i = 0; $i -lt $ScriptHT.Mission.Count; $i++) {
            # Section Date
            if ($ScriptHT.Mission[$i] -match [regex]::escape('["date"]')) {
                [void]$ScriptHT.MissionNew.Add("    [`"date`"] = ")
                [void]$ScriptHT.MissionNew.Add("    {")
                [void]$ScriptHT.MissionNew.Add("        [`"Day`"] = $($RandomHT.Day),")
                [void]$ScriptHT.MissionNew.Add("        [`"Year`"] = $($RandomHT.Year),")
                [void]$ScriptHT.MissionNew.Add("        [`"Month`"] = $($RandomHT.Month),")
                [void]$ScriptHT.MissionNew.Add("    }, -- end of [`"date`"]")
                $i = $i+5
            }
            # Winds
            elseif ($ScriptHT.Mission[$i] -match [regex]::escape('["wind"]')) {
                [void]$ScriptHT.MissionNew.Add("        [`"wind`"] = ")
                [void]$ScriptHT.MissionNew.Add("        {")
                [void]$ScriptHT.MissionNew.Add("            [`"at8000`"] = ")
                [void]$ScriptHT.MissionNew.Add("            {")
                [void]$ScriptHT.MissionNew.Add("                [`"speed`"] = $($RandomHT.WindSpeed8K),")
                [void]$ScriptHT.MissionNew.Add("                [`"dir`"] = $($RandomHT.WindDir8000),")
                [void]$ScriptHT.MissionNew.Add("            }, -- end of [`"at8000`"]")
                [void]$ScriptHT.MissionNew.Add("            [`"at2000`"] = ")
                [void]$ScriptHT.MissionNew.Add("            {")
                [void]$ScriptHT.MissionNew.Add("                [`"speed`"] = $($RandomHT.WindSpeed2K),")
                [void]$ScriptHT.MissionNew.Add("                [`"dir`"] = $($RandomHT.WindDir2000),")
                [void]$ScriptHT.MissionNew.Add("            }, -- end of [`"at2000`"]")
                [void]$ScriptHT.MissionNew.Add("            [`"atGround`"] = ")
                [void]$ScriptHT.MissionNew.Add("            {")
                [void]$ScriptHT.MissionNew.Add("                [`"speed`"] = $($RandomHT.WindSpeedGround),")
                [void]$ScriptHT.MissionNew.Add("                [`"dir`"] = $($RandomHT.WindDirGround),")
                [void]$ScriptHT.MissionNew.Add("            }, -- end of [`"atGround`"]")
                [void]$ScriptHT.MissionNew.Add("        }, -- end of [`"wind`"]")
                $i = $i+17
            }
            # Turbulence
            elseif ($ScriptHT.Mission[$i] -match [regex]::escape('"groundTurbulence"')){
                [void]$ScriptHT.MissionNew.Add("        [`"groundTurbulence`"] = $($RandomHT.Turbulence),")
            }
            # Temperature
            elseif ($ScriptHT.Mission[$i] -match [regex]::escape('"temperature"')){
                [void]$ScriptHT.MissionNew.Add("            [`"temperature`"] = $($RandomHT.Temperature),")
            }
            # Clouds
            elseif ($ScriptHT.Mission[$i] -match [regex]::escape('["clouds"]')) {
                # Cas des fichiers mission sans Preset Nuage
                $j = $i+4
                if ($ScriptHT.Mission[$j] -notmatch [regex]::escape('["preset"]')) {
                    $i--
                }
                [void]$ScriptHT.MissionNew.Add("        [`"clouds`"] = ")
                [void]$ScriptHT.MissionNew.Add("        {")
                [void]$ScriptHT.MissionNew.Add("            [`"thickness`"] = 200,")
                [void]$ScriptHT.MissionNew.Add("            [`"density`"] = 0,")
                if ($CloudsPreset) {
                    [void]$ScriptHT.MissionNew.Add("            [`"preset`"] = `"$CloudsPreset`",")
                }
                [void]$ScriptHT.MissionNew.Add("            [`"base`"] = $Cloudsbase,")
                [void]$ScriptHT.MissionNew.Add("            [`"iprecptns`"] = 0,")
                [void]$ScriptHT.MissionNew.Add("        }, -- end of [`"clouds`"]")
                $i = $i+7
            }
            # No Change
            else {
                [void]$ScriptHT.MissionNew.Add($ScriptHT.Mission[$i])
            }
        }

    #endregion

    #region Exports

        # Export du nouveau fichier mission
        ($ScriptHT.MissionNew -join "`n") + "`n" | Set-Content -NoNewline "$($ScriptHT.extractFolder)\mission"

        # Update du fichier miz
        $zipFile = Zip_File -ZipFile $ScriptHT.MizPath -DestFolder $ScriptHT.extractFolder
        if ($zipFile) {
            Write-Host "Fichier Miz mis a jour" -ForegroundColor "Green"
            # Discord Message
            if ($Discord) {
                SendDiscord
                Write-Host "Message Discord Envoyé" -ForegroundColor "Yellow"
            }
        }
        else {
            Write-Host "Echec de mise a jour du Miz" -ForegroundColor "Red"
            # Discord Message
            if ($Discord) {
                SendDiscord -Fail
                Write-Host "Message Discord Envoyé" -ForegroundColor "Yellow"
            }
        }

    #endregion

}
else {
    Write-Host "Echec de l'extraction des fichiers du Miz" -ForegroundColor "Red"
}