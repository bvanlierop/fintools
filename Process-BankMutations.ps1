# ABN AMRO:
# 1. In "Rekeningoverzicht", select account
# 2. Choose "Rekeningopties"
# 3. Choose "Mutaties downloaden"
# 4. Choose "Selecteer periode" : Begin "01-xx-20xx" Einde "31-xx-20xx" 
# 5. Choose "Formaat" : "TXT"
# 6. Click button "download"
# 7. Replace this file name with the downloaded file name (e.g. TXT200504102506.TAB)
$abnFileName = "TXT200504102506.TAB"
$downloadFolder = "C:\Temp"
Clear-Host

# Create total row collection
$rows = @()
$unknowns = @()

# RABOBANK uses CSV (with headers)
# First replace / by _ & " " by _
#Import-Csv "C:\Users\Bart\Downloads\CSV_A_20200326_151100(1).csv" |`
# ABN AMRO uses CSV with TABs (no header)
$abnFilePath = Join-Path -Path $downloadFolder -ChildPath $abnFileName
Import-Csv $abnFilePath -Header "IBAN_BBAN", "Munt", "Datum", "SaldoVoor", "SaldoNa", "Rentedatum", "Bedrag", "Omschrijving" -Delimiter "`t" |`
ForEach-Object {
  
    # Create custom bank info object
    $obj = New-Object -TypeName psobject
    #$obj | Add-Member -MemberType NoteProperty -Name Tegenpartij -Value "$($_.Naam_tegenpartij)"
    #$obj | Add-Member -MemberType NoteProperty -Name Bedrag -Value "$($_.Bedrag)"
    $obj | Add-Member -MemberType NoteProperty -Name Omschrijving -Value "$($_.Omschrijving)"
    $obj | Add-Member -MemberType NoteProperty -Name Bedrag -Value "$($_.Bedrag)"
    
    # add a method to an object
    #$obj | Add-Member -MemberType ScriptMethod -Name "PrintInfo" -Value {$this.Tegenpartij +' '+$this.Bedrag}
    $obj | Add-Member -MemberType ScriptMethod -Name "PrintInfo" -Value {$this.Omschrijving +' '+$this.Bedrag}
    # Add row to collection
    $rows += $obj
}

# Remember totals
[decimal]$totalAmountDebit = '0.0'
[decimal]$totalAmountCredit = '0.0'

[decimal]$totalLoonB = '0.0'
[decimal]$totalLoonA = '0.0'

[decimal]$totalSparenB = '0.0'
[decimal]$totalSparenA = '0.0'

[decimal]$totalKosten = '0.0'

[decimal]$totalBoodschappen = '0.0'
[decimal]$totalKleding = '0.0'
[decimal]$totalEntertainment = '0.0'
[decimal]$totalHypotheek = '0.0'
[decimal]$totalNietLevenVerzekeringen = '0.0'
[decimal]$totalBankkostenAbn = '0.0'
[decimal]$totalverzekeringsPakketGezin = '0.0'
[decimal]$totalCashGepind = '0.0'
[decimal]$totalBrandstof = '0.0'
[decimal]$totalRioolAfval = '0.0'
[decimal]$totalInternetTv = '0.0'
[decimal]$totalOzb = '0.0'
[decimal]$totalWater = '0.0'
[decimal]$totalCadeaus = '0.0'
[decimal]$totalStroomGas = '0.0'
[decimal]$totalAutoBelasting = '169.0' # tijdelijk hardcoded
[decimal]$totalOnbekend = '0.0'

# Process each row
Write-Host "Verwerk rijen ..." -ForegroundColor DarkYellow
$rows | ForEach-Object {
    
    # Convert text amount to number
    [decimal]$amount = [decimal]($_.Bedrag.Replace(',', '.'))
    
    # Check polarity
    If($_.Bedrag.StartsWith("-")) {
        # We have a debit amount (AF)
        $totalAmountDebit += $amount
        #Write-Host $_.PrintInfo() -ForegroundColor Red
    }
    Else {
        # We have credit amount (BIJ)
        $totalAmountCredit += $amount
        #Write-Host $_.PrintInfo() -ForegroundColor Green
    }
    
    #$categorie = $_.Tegenpartij.ToUpper()
    $categorie = $_.Omschrijving.ToUpper()
    
    switch($categorie.ToUpper()) { 
    
        # Zijn het boodschappen?
        {($_.Contains('FOOD')) -or 
         ($_.Contains('LIDL')) -or 
         ($_.Contains('ALDI')) -or 
         ($_.Contains('ALBERT HEIJN')) -or 
         ($_.Contains('JUMBO')) -or 
         ($_.Contains('BROOD')) -or 
         ($_.Contains('NJAM')) -or 
         ($_.Contains('BARRIER')) -or 
         ($_.Contains('ACTION')) -or 
         ($_.Contains('PRIMERA')) -or
         ($_.Contains('KRUIDVAT'))}  { 
          $totalBoodschappen += $amount; Break
        }

        # Is het kleding?
        {($_.Contains('JEANS'))} { 
          $totalKleding += $amount; Break
        }

        # Is het entertainment?
        {($_.Contains('SPOTIFY'))} { 
          $totalEntertainment += $amount; Break
        }

        {($_.Contains('HYPOTHEEK'))} { 
          $totalHypotheek += $amount; Break
        }

        {($_.Contains('OVERLIJDENSRISICO'))} { 
          $totalNietLevenVerzekeringen += $amount; Break
        }

        {($_.Contains('BETAALGEMAK'))} { 
          $totalBankkostenAbn += $amount; Break
        }

        {($_.Contains('ASR SCHADEVERZEKERING'))} { 
          $totalverzekeringsPakketGezin += $amount; Break
        }
        
        {($_.Contains('GELDAUTOMAAT'))} { 
          $totalCashGepind += $amount; Break
        }
        
        {($_.Contains('GULF'))} { 
          $totalBrandstof += $amount; Break
        }
        
        {($_.Contains('RIOOL')) -and
         ($_.Contains('AFVAL'))} { 
          $totalRioolAfval += $amount; Break
        }

        {($_.Contains('BUDGET ALLES-IN-1'))} {
          $totalInternetTv += $amount; Break
        }

        {($_.Contains('OZB/WOZ'))} {
          $totalOzb += $amount; Break
        }

        {($_.Contains('WATER'))} {
          $totalWater += $amount; Break
        }

        {($_.Contains('SALARISBETALING'))} {
          $totalLoonB += $amount; Break
        }

        {($_.Contains('DAGACTIVITEITEN'))} {
          $totalLoonA += $amount; Break
        }

        {($_.Contains('DAGACTIVITEITEN'))} {
          $totalLoonA += $amount; Break
        }

        {($_.Contains('KADO'))} {
          $totalCadeaus += $amount; Break
        }
        
        {($_.Contains('OXXIO'))} {
          $totalStroomGas += $amount; Break
        }

        {($_.Contains('REMI/SPAREN/')) -and
         ($_.Contains('B.H.J.'))} {
          $totalSparenB += $amount; Break
        }

        {($_.Contains('REMI/SPAREN/')) -and
         ($_.Contains('A.G.T.'))} {
          $totalSparenA += $amount; Break
        }

        # Onbekend
        default {
          #Write-Host "Onbekend bedrag: $($amount) -> Categorie: $($categorie)" -ForegroundColor Yellow
          $unknowns += $_
          $totalOnbekend += $amount; Break
        }
    }
}

$totalKosten = `
    $totalBoodschappen +` 
    $totalKleding +` 
    $totalEntertainment +` 
    $totalHypotheek +` 
    $totalNietLevenVerzekeringen +`
    $totalBankkostenAbn +` 
    $totalverzekeringsPakketGezin +`
    $totalCashGepind +` 
    $totalBrandstof +` 
    $totalRioolAfval +` 
    $totalInternetTv +` 
    $totalOzb +` 
    $totalWater +` 
    $totalCadeaus +` 
    $totalStroomGas +` 
    $totalAutoBelasting +`
    $totalOnbekend

# Print totals
Write-Host "----------------------------- + "
Write-Host "TOTAL CREDIT:`t EUR $($totalAmountCredit)"
Write-Host "TOTAL DEBIT :`t EUR $($totalAmountDebit)`n"

Write-Host "Loon B:`t`t`t EUR $($totalLoonB)" -ForegroundColor Green
Write-Host "Loon A:`t`t`t EUR $($totalLoonA)`n" -ForegroundColor Green

Write-Host "Sparen B:`t`t EUR $($totalSparenB)" -ForegroundColor Cyan
Write-Host "Sparen A:`t`t EUR $($totalSparenA)`n" -ForegroundColor Cyan

# Print categories
Write-Host "Boodschappen:`t EUR $($totalBoodschappen)" -ForegroundColor Red
Write-Host "Kleding:`t`t EUR $($totalKleding)" -ForegroundColor Red
Write-Host "Entertainment:`t EUR $($totalEntertainment)" -ForegroundColor Red
Write-Host "Hypotheek:`t`t EUR $($totalHypotheek)" -ForegroundColor Red
Write-Host "Nietleven-verz:`t EUR $($totalNietLevenVerzekeringen)" -ForegroundColor Red
Write-Host "Bankkosten ABN:`t EUR $($totalBankkostenAbn)" -ForegroundColor Red
Write-Host "Verzekeringen:`t EUR $($totalverzekeringsPakketGezin)" -ForegroundColor Red
Write-Host "Cash/PIN:`t`t EUR $($totalCashGepind)" -ForegroundColor Red
Write-Host "Brandstof:`t`t EUR $($totalBrandstof)" -ForegroundColor Red
Write-Host "Autobelasting:`t EUR $($totalAutoBelasting)" -ForegroundColor Red
Write-Host "Riool/Afval:`t EUR $($totalRioolAfval)" -ForegroundColor Red
Write-Host "Internet/TV:`t EUR $($totalInternetTv)" -ForegroundColor Red
Write-Host "OZB:`t`t`t EUR $($totalOzb)" -ForegroundColor Red
Write-Host "Water:`t`t`t EUR $($totalWater)" -ForegroundColor Red
Write-Host "Cadeaus:`t`t EUR $($totalCadeaus)" -ForegroundColor Red
Write-Host "Stroom/gas:`t`t EUR $($totalStroomGas)" -ForegroundColor Red
Write-Host "Onbekend:`t`t EUR $($totalOnbekend)" -ForegroundColor Red
Write-Host "----------------------------- + "
Write-Host "KOSTEN TOTAAL:`t EUR $($totalKosten)" -ForegroundColor Red

# Print unknowns
Write-Host "`n`nOnbekend details:"
$unknowns | ForEach-Object {
    Write-Host $_ -ForegroundColor Yellow
}
