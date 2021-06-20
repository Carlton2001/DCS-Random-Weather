param (
    [Parameter(Position=0,mandatory=$true)]
    [string]$MizPath
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

#endregion

#region Initialisations

    # Définition des variables
    $ScriptHT = [hashtable]::Synchronized(@{})
    $ScriptHT.Config = Get-Content "$PSScriptRoot\DCS-Random-Weather.json" | ConvertFrom-Json
    $ScriptHT.MizPath = $MizPath
    $ScriptHT.MizFile = Split-Path $MizPath -Leaf
    $ScriptHT.extractFolder = "$PSScriptRoot\extractFolder"
    $ScriptHT.MissionNew = New-Object System.Collections.ArrayList
    $ScriptHT.SevenZip = "$($env:ProgramFiles)\7-Zip\7z.exe"
    $ScriptHT.DiscordWebhook = "$PSScriptRoot\DiscordSendWebhook.exe"
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
        $randomDate = Random_Date
        $year   = $randomDate.Year.ToString()
        $month  = $randomDate.Month.ToString()
        $day    = $randomDate.Day.ToString()
        $MonthConfig = "Month$month"

        # Winds Speed & Direction
        $WindSpeedGround    = (Get-Random -Minimum 2 -Maximum 15).ToString() # Beaufort 2-6
        $WindDirGround      = (Get-Random -Maximum 360).ToString()
        $WindSpeed2000      = (Get-Random -Minimum 14 -Maximum 30).ToString() # Beaufort 7-10
        $WindDir2000        = (Get-Random -Maximum 360).ToString()
        $WindSpeed8000      = (Get-Random -Minimum 14 -Maximum 26).ToString() # Beaufort 7-9
        $WindDir8000        = (Get-Random -Maximum 360).ToString()

        # Turbulences
        $groundTurbulence = (Get-Random -Minimum 5 -Maximum 21).ToString()

        # Temperature
        $Temperature = Random_Temp $ScriptHT.Config.Temperatures.$MonthConfig.$($ScriptHT.Theatre)

        # Preset Cloud
        $Preset = "Preset$(Get_CloudFormation $MonthConfig)"
        $CloudsPreset = $ScriptHT.Config.CloudPresets.$Preset.Name
        $Cloudsbase = $ScriptHT.Config.CloudPresets.$Preset.Base

        # Informations
        Write-Host "Date : $randomDate"
        Write-Host "Temperature : $Temperature"
        Write-Host "Winds : $WindSpeedGround m/s @ $WindDirGround - $WindSpeed2000 m/s @ $WindDir2000 - $WindSpeed8000 m/s @ $WindDir8000"
        Write-Host "Turbulences : $groundTurbulence"
        Write-Host "Cloud : $CloudsPreset"

        # Discord
        if ($CloudsPreset) {
            $CloudsInformations = ("{0}[]({1})" -f $ScriptHT.Config.CloudPresets.$Preset.Metar, $ScriptHT.Config.CloudPresets.$Preset.Image)
        }
        else {
            $CloudsInformations = "Temps Clair, pas de nuage"
        }
        $DiscordMessage = ("{0} **Votre meteo de la journee pour {1}**\nNous sommes le {2} {3} {4}, il fait {5} degres\n__Vents__ :\nAu sol : {6} kts du {7}\n2000m : {8} kts du {9}\n8000m : {10} kts du {11}\nTurbulences : {12} kts\n__Nuages__ :\n" -f $ScriptHT.Config.Discord.$($ScriptHT.Theatre).Flag, $ScriptHT.Config.Discord.$($ScriptHT.Theatre).City, $day, $ScriptHT.Config.FrenchDates.$month, $year, $Temperature, $(SpeedKts $WindSpeedGround), $(WindsInverter $WindDirGround), $(SpeedKts $WindSpeed2000), $(WindsInverter $WindDir2000), $(SpeedKts $WindSpeed8000), $(WindsInverter $WindDir8000), $($(SpeedKts $groundTurbulence)/10))
        $DiscordMessage = $DiscordMessage + $CloudsInformations

    #endregion

    #region Traitement

        for ($i = 0; $i -lt $ScriptHT.Mission.Count; $i++) {
            # Section Date
            if ($ScriptHT.Mission[$i] -match [regex]::escape('["date"]')) {
                [void]$ScriptHT.MissionNew.Add("    [`"date`"] = ")
                [void]$ScriptHT.MissionNew.Add("    {")
                [void]$ScriptHT.MissionNew.Add("        [`"Day`"] = $day,")
                [void]$ScriptHT.MissionNew.Add("        [`"Year`"] = $year,")
                [void]$ScriptHT.MissionNew.Add("        [`"Month`"] = $month,")
                [void]$ScriptHT.MissionNew.Add("    }, -- end of [`"date`"]")
                $i = $i+5
            }
            # Winds
            elseif ($ScriptHT.Mission[$i] -match [regex]::escape('["wind"]')) {
                [void]$ScriptHT.MissionNew.Add("        [`"wind`"] = ")
                [void]$ScriptHT.MissionNew.Add("        {")
                [void]$ScriptHT.MissionNew.Add("            [`"at8000`"] = ")
                [void]$ScriptHT.MissionNew.Add("            {")
                [void]$ScriptHT.MissionNew.Add("                [`"speed`"] = $WindSpeed8000,")
                [void]$ScriptHT.MissionNew.Add("                [`"dir`"] = $WindDir8000,")
                [void]$ScriptHT.MissionNew.Add("            }, -- end of [`"at8000`"]")
                [void]$ScriptHT.MissionNew.Add("            [`"at2000`"] = ")
                [void]$ScriptHT.MissionNew.Add("            {")
                [void]$ScriptHT.MissionNew.Add("                [`"speed`"] = $WindSpeed2000,")
                [void]$ScriptHT.MissionNew.Add("                [`"dir`"] = $WindDir2000,")
                [void]$ScriptHT.MissionNew.Add("            }, -- end of [`"at2000`"]")
                [void]$ScriptHT.MissionNew.Add("            [`"atGround`"] = ")
                [void]$ScriptHT.MissionNew.Add("            {")
                [void]$ScriptHT.MissionNew.Add("                [`"speed`"] = $WindSpeedGround,")
                [void]$ScriptHT.MissionNew.Add("                [`"dir`"] = $WindDirGround,")
                [void]$ScriptHT.MissionNew.Add("            }, -- end of [`"atGround`"]")
                [void]$ScriptHT.MissionNew.Add("        }, -- end of [`"wind`"]")
                $i = $i+17
            }
            # Turbulence
            elseif ($ScriptHT.Mission[$i] -match [regex]::escape('"groundTurbulence"')){
                [void]$ScriptHT.MissionNew.Add("        [`"groundTurbulence`"] = $groundTurbulence,")
            }
            # Temperature
            elseif ($ScriptHT.Mission[$i] -match [regex]::escape('"temperature"')){
                [void]$ScriptHT.MissionNew.Add("            [`"temperature`"] = $Temperature,")
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
            $DiscordSent = Invoke-Expression -Command ".`"$($ScriptHT.DiscordWebhook)`" -m `"$DiscordMessage`" -n `"$($ScriptHT.Config.Discord.BotName)`" -a $($ScriptHT.Config.Discord.BotAvatar) -w $($ScriptHT.DiscordLink)"
        }
        else {
            Write-Host "Echec de mise a jour du Miz" -ForegroundColor "Red"
            # Discord Message
            $DiscordSent = Invoke-Expression -Command ".`"$($ScriptHT.DiscordWebhook)`" -m `"Probleme de generation de la mission $($ScriptHT.MizFile)`" -n `"IRRE_Serveur`" -w $($ScriptHT.DiscordLink)"
        }
        Write-Host "Message Discord Envoyé" -ForegroundColor "Yellow"

    #endregion

}
else {
    Write-Host "Echec de l'extraction des fichiers du Miz" -ForegroundColor "Red"
}