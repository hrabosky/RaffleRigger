Add-Type -AssemblyName Microsoft.VisualBasic
<#
    _   __              ____    __              _     
   / | / /__  _________/ / /_  / /___  ____ _  (_)___ 
  /  |/ / _ \/ ___/ __  / __ \/ / __ \/ __ `/ / / __ \
 / /|  /  __/ /  / /_/ / /_/ / / /_/ / /_/ / / / /_/ /
/_/ |_/\___/_/   \__,_/_.___/_/\____/\__, (_)_/\____/ 
                                    /____/            

 ____ ____ ____ ____ ____ ____ 
||R |||a |||f |||f |||l |||e ||
||__|||__|||__|||__|||__|||__||
|/__\|/__\|/__\|/__\|/__\|/__\|
 ____ ____ ____ ____ ____ ____ 
||R |||i |||g |||g |||e |||r ||
||__|||__|||__|||__|||__|||__||
|/__\|/__\|/__\|/__\|/__\|/__\|

V1.0

Hey all, thank you for your interest in my script. 
You may not know me, but my name is John and I have a site called nerdblog.io
I will 

Improvements to come:
-Ability to blacklist users
-Multiple sculpts with different quantities & prices
-Weighted raffles, based on users DKP points (Binge's Idea).
-Nerdblog.io auto win

#>

$AmountOfWinners = Read-Host "How many individual winners?" #This is the amount of people to win
$RaffleEntriesCsv = ".\RaffleEntries.csv" <#
This is the file location of your keycaps.csv file;
The idea is you have a google form for your raffle, with the fields paypal email, first name, last name,
and a yes or no question if they're international. You would then export your results from your google sheets
into an excel spreadsheet with the following cell values, A1 = "paypal", B1 = "fname", C1 = "lname", D1 = "international".
Save the excel file as a .csv (you can also do it with notepad, but have to separate with commas.)
#>
$WinnerPaypalCsv = ".\Winners.csv" #This is the location of where it will save the text file.

#Prompt user for necessary information (paypal invoicing).
$KeycapName = Read-Host "What is the keycap name you are raffling?"
$Keycapvalue = Read-Host "How much are you selling the $KeycapName for?"

$EntryList = Import-Csv $RaffleEntriesCsv | sort paypal -Unique
<#the line above imports the entries and removes duplicates based on Paypal emails addresses.
I have forseen issues if a husband and wife use the same paypal email for example;
or if someone wants to be sneaky and link a different email to their paypal account.
My thought to ignore duplicates would be by delivery address; but didn't include it for this script.
#>

#Declare arrays | Variables
$PaypalArray = @()
$PaypalEmail = @()
$International = @()
$ShippingArray = @()
$WinnerArray = @()
$ShippingCost = ""
$data = ""

#Cycles through each raffle entry, and stores their paypal email, along with if they're international or not.
ForEach ($u in $EntryList){
$PaypalEmail += $($u.paypal)
$International += $($u.international)
}

#Adds both entries to a jagged array
$PaypalArray = $PaypalEmail,$International


#Picks the winners based on how the $AmountOfWinners variable | will not chose same winner twice.
$WinnerArray += Get-Random -InputObject $PaypalArray[0] -Count $AmountOfWinners

#Imports an array so you can append.
if (Test-Path $WinnerPaypalCsv){
    [array]$data = Import-Csv -Path $WinnerPaypalCsv
}

#Finds the index of the winners from the first array, and finds the if they are international or not.
For ($i=0; $i -lt $WinnerArray.Length; $i++){
$Search = $PaypalArray[0].IndexOf($WinnerArray[$i])

#converts a yes or no to a dollar value, change th values here if you want to adjust shipping pricing.
if ($PaypalArray[1][$Search] -eq 'Yes'){
    $ShippingCost = '13.00'
}Else{
   $ShippingCost = '5.00'
}

<#
Adds the winners to csv file. / the properties below are what gets exported, while only the necessary columns
were added, you could hard code in due date, or Currency to USD... etc
#>
$data += New-Object PSObject -Property @{
    'Recipient Email' = $WinnerArray[$i]
    'Recipient First Name' = ""
    'Recipient Last Name' = ""
    'Invoice Number' = ""
    'Due Date' = ""
    'Reference' = ""
    'Item Name' = $KeycapName
    'Description' = ""
    'Item Amount' = $Keycapvalue
    'Shipping Amount' = $ShippingCost
    'Discount' = ""
    'Currency' = ""
    'Note to Customer' = ""
    'Terms and Conditions' = ""
    'Memo to Self' = ""
}

$data | Select-Object "Recipient Email", "Recipient First Name", "Recipient Last Name", "Invoice Number", "Due Date", "Reference",
 "Item Name", "Description", "Item Amount", "Shipping Amount", "Discount","Currency","Note to Customer","Terms and Conditions", "Memo to Self" | Export-CSV -Path $WinnerPaypalCsv -NoTypeInformation
}


Write-Host "
 ____ ____ ____                     
||T |||h |||e ||                    
||__|||__|||__||                    
|/__\|/__\|/__\|                    
 ____ ____ ____ ____ ____ ____ ____ 
||W |||i |||n |||n |||e |||r |||s ||
||__|||__|||__|||__|||__|||__|||__||
|/__\|/__\|/__\|/__\|/__\|/__\|/__\|
 ____ ____ ____ ____                
||H |||a |||v |||e ||               
||__|||__|||__|||__||               
|/__\|/__\|/__\|/__\|               
 ____ ____ ____ ____                
||B |||e |||e |||n ||               
||__|||__|||__|||__||               
|/__\|/__\|/__\|/__\|               
 ____ ____ ____ ____ ____ ____      
||C |||h |||o |||s |||e |||n ||     
||__|||__|||__|||__|||__|||__||     
|/__\|/__\|/__\|/__\|/__\|/__\|     
"
start-sleep -s 1
#Opens Csv file
Invoke-Item $WinnerPaypalCsv