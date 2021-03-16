## Find or Create output folder based on parameters
## Start MKV and send files to output folder
## Loop through output folder
## Rename each file based on parameters

Param (
    [Parameter (Mandatory=$true)]
    [string]$OutputPath,
    [Parameter (Mandatory=$true)]
    [string]$Title,
    [Parameter (ParameterSetName="ShowType")]
    [switch]$IsTVShow,
    [Parameter (ParameterSetName="ShowType")]
    [int]$Season,
    [Parameter (ParameterSetName="ShowType")]
    [int]$BeginningEpisode
)

$ErrorActionPreference = "Stop"

Function Find-OutputFolder
{
    param(
        [string] $Path
    )

    Write-Host "Looking for Output folder.."
    if (!(Test-Path $Path))
    {
        Write-Host "Output folder not found, creating it..."
        New-Item -ItemType Directory -Path $Path
        Write-Host "$Path directory created."
    }
    Write-Host "$Path directory found."
}

Function Invoke-MKVCommand()
{
    if ($IsTVShow) {
        $OutputPath = Join-Path -Path $OutputPath -ChildPath "$Title\Season $Season"
    }

    Find-OutputFolder $OutputPath

    $makeMKVCommand = "makemkvcon";
    $mkvArgs = ( "mkv","-r","--minlength=1200 --messages=makemkv.out --progress=makemkv.out","disc:0","all","`"$OutputPath`"" )

    try 
    {
        $makeMKVSource = (Get-Command -Name $makeMKVCommand).Source
        $makeMkvProcess = Start-Process $makeMKVSource -ArgumentList $mkvArgs -NoNewWindow -PassThru
    }
    catch
    {
        Write-Host "Unable to execute MakeMKV: $_.message"
    }

    do
    {
        $stdOut = Get-Content "makemkv.out" -Tail 1 | ConvertFrom-Csv -Header "Code", "var1", "var2", "var3"

        if ($stdOut.Code -and $stdOut.Code.Contains(":"))
        {
            $code = ($stdOut.Code -split ":")[0]
            
            switch ($code)
            {
                "PRGV" {
                    $currentProgress = ([int]($stdOut.Code -split ":")[1] / [int]$stdOut.var2)
                    $totalProgress = ([int]$stdOut.var1 / [int]$stdOut.var2) 
                    $currentProgressPct = [int]($currentProgress * 100)
                    $totalProgressPct = [int]($totalProgress * 100)
                    $progress = "Title: $currentProgressPct %.  Total: $totalProgressPct % "
                    Write-Progress -Activity "Running $makeMKVSource..." -PercentComplete $currentProgressPct -CurrentOperation $progress -Status "Ripping..."
                }
                "MSG" {
                    $message = $stdOut.var3
                    Write-Progress -Activity "Running $makeMKVSource..." -CurrentOperation $message -Status "Reading..."
                }
                default {
                    $message = $stdOut.var3
                    Write-Progress -Activity "Running $makeMKVSource..." -CurrentOperation $message -Status "Processing..."
                }
            }
        }
        else 
        {
            Write-Progress -Activity "Running $makeMKVSource..." -Status "Reading..."
        }
        
        Start-Sleep -Seconds 5

    } until ($makeMkvProcess.HasExited)

    Rename-MKVFiles $OutputPath
}

Function Rename-MKVFiles
{
    param(
        [string] $Path
    )

    if ($IsTVShow){
        Rename-MultipleMKVFiles $Path
    }
    else {
        Rename-SingleMKVFile $Path
    }
}

Function Rename-MultipleMKVFiles()
{
    param(
        [string] $Path
    )

    if (Test-Path($Path))
    {
        $currentEpisode = $BeginningEpisode
        $fileNames = Get-ChildItem $Path -Filter "*t0*.mkv" | Select-Object FullName

        foreach($fileName in $fileNames)
        {
            Write-Host "Renaming $($fileName.FullName)..."

            if ($currentEpisode -lt 10)
            {
                $episode = "0$currentEpisode"
            }
            else 
            {
                $episode = "$currentEpisode"
            }

            if ($Season -lt 10)
            {
                $formattedSeason = "0$Season"
            }
            else {
                $formattedSeason = "$Season"
            }

            $newFileName = "$Title s$formattedSeason.e$episode.mkv"
            Rename-Item -Path $fileName.FullName -NewName $newFileName
            Write-Host "... renamed to $newFileName"

            $currentEpisode++
        }
    }
}

Function Rename-SingleMKVFile
{
    param(
        [string] $Path
    )

    if (Test-Path($Path))
    {
        $fileName = Get-ChildItem $Path -Filter "*t0*.mkv" | Select-Object FullName
        Write-Host "Renaming $($fileName.FullName)..."
        $newFileName = "$($Title).mkv"
        Rename-Item -Path $fileName.FullName -NewName $newFileName
        Write-Host "... renamed to $newFileName"
    }
}

Invoke-MKVCommand
