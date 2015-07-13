# Contributors: Eric Dixon, Keenan Newton, Brent Groom, Cem Aykan, Dan Benson
param([string]$name="", [switch]$list,[switch]$allgroups)

Add-PSSnapin Microsoft.FASTSearch.PowerShell -erroraction SilentlyContinue 


function save-KeywordContext($ssg, $folder) 
{
    #Save Contexts
    saveContext -ssg $ssg -contextCSVFile "$folder\Context.csv"
    
    #Save Keywords
    saveKeyWord -ssg $ssg -keywordCSVFile "$folder\Keyword.csv"
    
    #Save BestBets
    saveBestBet -ssg $ssg -bestBetCSVFile "$folder\BestBet.csv"

    #Save Visual BestBets
    saveVisualBestBet -ssg $ssg -visualBestBetCSVFile "$folder\VisualBestBet.csv"
    
    #Save Document Promotions
    saveDocumentPromotion -ssg $ssg -documentPromotionCSVFile "$folder\DocumentPromotion.csv"
}

function saveContext($ssg, $contextCSVFile) 
{
    Write-Host "Creating file $contextCSVFile"
    $csvFile = New-item -itemtype file $contextCSVFile  -force
    "User Context, Ask Me About, Office Location" | Out-File $contextCSVFile
     
    foreach($context in $ssg.Contexts)
    {
        if(!($context.Name))
        {
            continue
        }
        Write-Host "Saving Context: " $context.Name
        
        $askMeAbout = ""
        $officLocations = ""
        
        foreach($exp in $context.ContextExpression)
        {
            if($exp.Name -eq "SPS-Responsibility" -and $exp.Value)
            {
                $askMeAbout += $exp.Value + ";"
            }
            elseif ($exp.Name -eq "SPS-Location" -and $exp.Value)
            {
                $officLocations += $exp.Value + ";"
            }
        }
        $askMeAbout = $askMeAbout.Trim(";")
        $officLocations = $officLocations.Trim(";")
        
        $line = '"' + $context.Name +'","'+ $askMeAbout +'","'+ $officLocations +'"'
        $line | Out-File $contextCSVFile -append
           
    }
}

function saveKeyWord($ssg, $keywordCSVFile) 
{
    $foundkeywords = $false

    foreach($keyword in $ssg.Keywords) 
    {
        if(!($keyword.Term))
        {
            continue
        }
        if($foundkeywords -eq $false)
        {
          Write-Host "Creating file $keywordCSVFile"
          $csvFile = New-item -itemtype file $keywordCSVFile  -force
          "Keyword, Definition, Two-Way Synonym, One-Way Synonym" | Out-File $keywordCSVFile
          $foundkeywords = $true
        }
        Write-Host "Saving Keyword: " $keyword.Term
        
        $synonyms_oneway = ""
        $synonyms_twoway = ""
        foreach($synonym in $keyword.synonyms)
        {
            if($synonym.ExpansionType -eq "TwoWay")
            {
                $synonyms_twoway += $synonym.Term + ";"
            }    
            elseif($synonym.ExpansionType -eq "OneWay")
            {
                $synonyms_oneway += $synonym.Term + ";"
            }
        }
        $synonyms_oneway = $synonyms_oneway.Trim(";")
        $synonyms_twoway = $synonyms_twoway.Trim(";")
        
        #$line = '"'+$keyword.Term+'","'+$keyword.Definition+'","'+$synonyms_twoway+'","'+$synonyms_oneway+'"'
        $line = '"'+$keyword.Term+'"," ","'+$synonyms_twoway+'","'+$synonyms_oneway+'"'
        $line | Out-File $keywordCSVFile -append
    }
    
}

function saveBestBet($ssg, $bestBetCSVFile) 
{
    $foundbb = $false
    foreach($bb in $ssg.BestBets)
    {
        if(!($bb.Name))
        {
            continue
        }
        if($foundbb -eq $false)
        {
            Write-Host "Creating file $bestBetCSVFile"
            $csvFile = New-item -itemtype file $bestBetCSVFile  -force
            "BestBet, User Context, Keyword, Description, Url, Start Date, End Date, Position" | Out-File $bestBetCSVFile
            $foundbb = $true
        }
        Write-Host "Saving BestBet: " $bb.Name
        
        $contexts = ""
        foreach($context in $bb.contexts)
        {
            $contexts += $context.Name + ";"
        }
        $contexts = $contexts.Trim(";")
        $line = '"'+$bb.Name+'","'+$contexts+'","'+$bb.Keyword.Term+'","'+$bb.Description+'","'+$bb.Uri+'","'+$bb.StartDate+'","'+$bb.EndDate+'","'+$bb.Position+'"'
        $line | Out-File $bestBetCSVFile -append    
    }
}

function saveVisualBestBet($ssg, $visualBestBetCSVFile) 
{
    $foundvbb = $false
    foreach($vbb in $ssg.FeaturedContent)
    {
        if(!($vbb.Name))
        {
            continue
        }
        if($foundvbb -eq $false)
        {
            Write-Host "Creating file $visualBestBetCSVFile"
            $csvFile = New-item -itemtype file $visualBestBetCSVFile  -force
            "Visual BestBet, User Context, Keyword, Url, Start Date, End Date, Position"  | Out-File $visualBestBetCSVFile
            $foundvbb = $true
        }
        Write-Host "Saving Visual BestBet: " $vbb.Name
        
        $contexts = ""
        foreach($context in $vbb.contexts)
        {
            $contexts += $context.Name + ";"
        }
        $contexts = $contexts.Trim(";")
        $line = '"'+$vbb.Name+'","'+$contexts+'","'+$vbb.Keyword.Term+'","'+$vbb.Uri+'","'+$vbb.StartDate+'","'+$vbb.EndDate+'","'+$vbb.Position+'"'
        $line | Out-File $visualBestBetCSVFile -append
    }
}

function saveDocumentPromotion($ssg, $documentPromotionCSVFile) 
{
    $foundpromo = $true

    foreach($promo in $ssg.Promotions)
    {
        if(!($promo.Name))
        {
            continue
        }
        if($foundpromo -eq $false)
        {
            Write-Host "Creating file $documentPromotionCSVFile"
            $csvFile = New-item -itemtype file $documentPromotionCSVFile  -force
            "Title, User Context, Keyword, Url, Start Date, End Date, Boost Value" | Out-File $documentPromotionCSVFile
            $foundpromo = $true
        }
        Write-Host "Saving Promotion: " $promo.Name
        
        $contexts = ""
        foreach($context in $promo.contexts)
        {
            $contexts += $context.Name + ";"
        }
        $contexts = $contexts.Trim(";")
        
        $urls = ""
        foreach($url in $promo.PromotedItems)
        {
            $urls += $url.DocumentId + "|"
        }
        $urls = $urls.Trim("|")
        
        $line = '"'+$promo.Name+'","'+$contexts+'","'+$promo.Keyword.Term+'","'+$urls+'","'+$promo.StartDate+'","'+$promo.EndDate+'","'+$promo.BoostValue+'"'
        $line | Out-File $documentPromotionCSVFile -append
    }
}

function main($name, $list, $allgroups)
{
    if($list -or $allgroups)
    {
        $ssg = Get-FASTSearchSearchSettingGroup
        if($ssg)
        {
            foreach($group in $ssg)
            {
                Write-Host "Search Setting Group Name: " $group.Name
                if($allgroups)
                {
                    $ssg = Get-FASTSearchSearchSettingGroup -name $group.Name
                    if($ssg)
                    {
                        save-KeywordContext -ssg $ssg -folder "csv-$($group.Name)"
                    }
                    else
                    {
                        Write-Host -foregroundcolor 'red' "Unable to get Search Setting Group for name=" $name
                    }
                }
            }
        }
        else
        {
            Write-Host -foregroundcolor 'red' "Unable to find a Search Setting Group."
        }
        return
    }
    
    if($name)
    {
        $ssg = Get-FASTSearchSearchSettingGroup -name $name
        if($ssg)
        {
            save-KeywordContext -ssg $ssg -folder "csv-$($name)"
        }
        else
        {
            Write-Host -foregroundcolor 'red' "Unable to get Search Setting Group for name=" $name
        }
        return
    }
    
    $ssg = Get-FASTSearchSearchSettingGroup
    if($ssg -and $ssg.Length -eq 1)
    {    
        save-KeywordContext -ssg $ssg -folder "csv-$($name)"
    }
    else
    {
        Write-Host -foregroundcolor 'red' "Found more than one Search Setting Group."
        Write-Host -foregroundcolor 'red' "Use the -list switch to list all group names."
        Write-Host -foregroundcolor 'red' "Use the -name option to export a specific Search Setting Group."
        Write-Host -foregroundcolor 'red' "Example: <script_file>.ps1 -list"
        Write-Host -foregroundcolor 'red' "Example: <script_file>.ps1 -name <NameOfSearchSettingGroup>"
    }
}


main $name $list $allgroups




# SIG # Begin signature block
# MIINGAYJKoZIhvcNAQcCoIINCTCCDQUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/Ah3QCFzTkCk9Y+utC24CUYv
# 4s6gggpaMIIFIjCCBAqgAwIBAgIQBg4i3l65iHFvsYhyMrxXATANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE0MDcxNzAwMDAwMFoXDTE1MDcy
# MjEyMDAwMFowaTELMAkGA1UEBhMCQ0ExCzAJBgNVBAgTAk9OMREwDwYDVQQHEwhI
# YW1pbHRvbjEcMBoGA1UEChMTRGF2aWQgV2F5bmUgSm9obnNvbjEcMBoGA1UEAxMT
# RGF2aWQgV2F5bmUgSm9obnNvbjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAN0ZOWCIOEyhtxA/koB0azqKK40Pw3fa8GLif/ZM0cXJWGawkVgxOMbejeJW
# YCqXgEHF2MX/cJY8svCmLlX8M7AdjXYgtAS+C+cEHxrGAMMzj3/9EOu6DjzxIcwL
# l1GKoUwy8X3/GRGk3sBWT5CwKYRJdh9goWy74ltZN+sTKKeDHqpfuvxye6c++PC7
# 86wzm4MwfOIuPE9StFeo/0nKheEukfK9cpthlE5dUHpW0OjmJdX/g0mEdIjm2/Q2
# 50fzQyLQXOuMVIJ4Qk2comMDNRvZZvSPOBwWZ6fAR4tXfZwlpUcLv3wBbIjslhT7
# XasCm73TdBj+ZFDx2tUtpWguP/0CAwEAAaOCAbswggG3MB8GA1UdIwQYMBaAFFrE
# uXsqCqOl6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBS+FASXsrRle2tLXdkVyoT1Dbw7
# QDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAw
# bjA1oDOgMYYvaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1j
# cy1nMS5jcmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtY3MtZzEuY3JsMEIGA1UdIAQ7MDkwNwYJYIZIAYb9bAMBMCowKAYIKwYB
# BQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwgYQGCCsGAQUFBwEB
# BHgwdjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsG
# AQUFBzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEy
# QXNzdXJlZElEQ29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG
# 9w0BAQsFAAOCAQEAbhjcmv+WCZwWCIYQwiEsH94SesBr0cPqWjEtJrBefqU9zFdB
# u5oc/WytxdCkEj5bxkoN9aJmuDAZnHNHBwIYeUz0vNByZRz6HsPzNPxLxThajJTe
# YOHlSTMI/XzBhJ7VzCb3YFhkD5f9gcJ5n+Z94ebd/1SoIvc9iwC3tTf5x2O7aHPN
# iyoWLTV4+PgDntCy/YDj11+uviI9sUUjAajYPEDvoiWinyT+7RlbStlcEuBwqvqT
# nLaiRsK17rjawW87Nkq/jB8rymUR/fzluIpHmPA4P0NazH4v5f62mpMFqdk0osMU
# QJ/qqACQ+2+/eAw7Gr6igNvlsxQpPfxsPNtUkTCCBTAwggQYoAMCAQICEAQJGBtf
# 1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIG
# A1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTEzMTAyMjEyMDAw
# MFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGln
# aUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsxSRnP0PtFmbE620T1
# f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawOeSg6funRZ9PG+ykn
# x9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJRdQtoaPpiCwgla4c
# SocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEcz+ryCuRXu0q16XTm
# K/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whkPlKWwfIPEvTFjg/B
# ougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8lk9ECAwEAAaOCAc0w
# ggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQM
# MAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDov
# L29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8E
# ejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARIMEYwOAYKYIZIAYb9
# bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BT
# MAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAfBgNV
# HSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQsFAAOCAQEA
# PuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/Er4v97yrfIFU3sOH2
# 0ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3nEZOXP+QsRsHDpEV
# +7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpoaK+bp1wgXNlxsQyP
# u6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW6Fkd6fp0ZGuy62ZD
# 2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ92JuoVP6EpQYhS6S
# kepobEQysmah5xikmmRR7zGCAigwggIkAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUw
# EwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20x
# MTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcg
# Q0ECEAYOIt5euYhxb7GIcjK8VwEwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwx
# CjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGC
# NwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFCE4bME2+b29y5xu
# AdcuHMNR4yjoMA0GCSqGSIb3DQEBAQUABIIBAItJBaLrtcQEQ8+7/5vCG8sRAGEN
# th5j2gHo0SBXbHGBUBcMavUZF6J4cs5lUKwXmauO7Mk854k9XxvnG0Fc2WczFTjU
# 9i4nHYQnSWSbCW4zulyitftyplfoZx7LcUnNFZB/7bY4TZHmyAxFN4eTzeYZ/KQH
# mHZLikYahluOVu3lTqVxfIkJEuvuvrQOT3klE3Lg1sdoXgbj2Hf/xmCi965LyHDZ
# HZXmhgkXPlEEKGkSghzi13YwN3JYwuCIhDeVEfcpYx6Zqix6zKL/EFuVtlVagTuV
# QNqAMHhNQXvYNcEiRLLGuvgbDS/W4XRtHOFoxtHyHaZhERn67X4XQVMHEpo=
# SIG # End signature block
