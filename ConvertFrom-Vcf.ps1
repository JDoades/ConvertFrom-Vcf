

function New-Card {
    return [PSCustomObject][ordered]@{'First Name'="";'Categories'='';'Organization'='';'Address'='';'Mobile Phone'='';'Business Phone'='';'Business Phone 2'='';'Home Phone'='';'E-mail Address'='';'E-mail 2 Address'='';'E-mail 3 Address'='';'E-mail 4 Address'='';'Note'=''}
}

function Generate-CSVFromContactFile($filename){
    $content = Get-Content $fileName
    $cardCount = 0;
    $currentCard = New-Card
    #$state = 'out'
    foreach ($line in $content) {
        if ($line -match "^PRODID.*") {
    #        $state = 'in'
            $cardCount++;
            $telephones = 0;
            $emails = 0;
            $telephonecolumnname = ""
            $emailcolumnname = ""
            $currentCard = New-Card
            continue;
        }
        if ($line -match "^END:VCARD.*") {
    #        $state = 'out'
            $currentCard
            continue;
        }

        if ($line -match "^PHOTO;.*") {
    #        $state = 'photo'
            continue;
        }

        if ($line -match "FN:.*") {
            $tokens = $line -split ":"
            $currentCard."First Name" = $tokens[1] -replace "\\,", ","
        }

        if ($line -match "ORG:.*") {
            $tokens = $line -split ":"
            $currentCard."Organization" = (($tokens[1] -split ';') -join "`n") -replace "\\,", ","
        }

        if ($line -match "^ADR.*:.*" -or $line -match "^item1.ADR.*:.*") {
            $tokens = $line -split ":"
            $currentCard."Address" = (($tokens[1] -split ';').Where({-not [string]::IsNullOrWhitespace($_)}) -join "`n") -replace "\\,", "," -replace "\\n", "`n"
        }

        if ($line -match "TEL;") {
            $telephones++;
            $tokens = $line -split ":"
            $tokens[1] = ($tokens[1] -replace "\+44" , "0")
			$tokens[1] = (((($tokens[1] -split ';') -join "`n") -replace "\\,", "," ) -replace "\+44" , "0")
			$tokens[1] = AddSpaceAfterFiveChars $tokens[1]
            switch ($telephones) {
                1 {$telephonecolumnname = "Mobile Phone"}
                2 {$telephonecolumnname = "Business Phone"}
                3 {$telephonecolumnname = "Business Phone 2"}
                4 {$telephonecolumnname = "Home Phone"}
                default{}
                }
                $currentCard.$telephonecolumnname = $tokens[1]
                
            
        }
        if ($line -match "Email;") {
            $emails++;
            $tokens = $line -split ":"
             switch ($emails) {
                1 {$emailcolumnname = "E-mail Address"}
                2 {$emailcolumnname = "E-mail 2 Address"}
                3 {$emailcolumnname = "E-mail 3 Address"}
                4 {$emailcolumnname = "E-mail 4 Address"}
                default{}
                }
            $currentCard.$emailcolumnname = (($tokens[1] -split ';') -join "`n") -replace "\\,", ","
        }

        if ($line -match "NOTE:") {
            $telephones++;
            $tokens = $line -split ":"
            $currentCard."Note" = (($tokens[1] -split ';') -join "`n") -replace "\\,", ","
        }

        if ($line -match "CATEGORIES.*") {
            $tokens = $line -split ":"
            $currentCard."Categories" = $tokens[1]
        }  
    }
}

function AddSpaceAfterFiveChars {
    param(
        [string]$inputString
    )

    # Check if the string length is greater than 5 and if the 6th character is not a space
    if ($inputString.Length -gt 5 -and $inputString[5] -ne ' ') {
        # Insert a space after the 5th character
        $outputString = $inputString.Insert(5, ' ')
    } else {
        $outputString = $inputString
    }

    return $outputString
}

do{

$ErrorActionPreference = "Stop"

$files = Get-ChildItem "C:\contact-conversion\inbound\*" -include *.vcf


foreach ($file in $files) {
     
     $outboundfile =  "C:\contact-conversion\outbound\" + $file.BaseName + ".csv"
     $renamedfile = $file.BaseName + ".vcf.bak"
     Generate-CSVFromContactFile($file.FullName) | Export-Csv $outboundfile -NoTypeInformation
    
     Rename-Item -Path $file.FullName -NewName $renamedfile
     write-host "completed $file"
}

Start-Sleep -Seconds 5
}
while ($true)