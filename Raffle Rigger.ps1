Add-Type -AssemblyName Microsoft.VisualBasic
<#
    _   __              ____    __              _
   / | / /__  _________/ / /_  / /___  ____ _  (_)___
  /  |/ / _ \/ ___/ __  / __ \/ / __ \/ __ `/ / / __ \
 / /|  /  __/ /  / /_/ / /_/ / / /_/ / /_/ / / / /_/ /
/_/ |_/\___/_/   \__,_/_.___/_/\____/\__, (_)_/\____/
                                    /____/

 ____ ____ ____
||H |||W |||S ||
||__|||__|||__||
|/__\|/__\|/__\|
 ____ ____ ____ ____ ____ ____
||R |||a |||f |||f |||l |||e ||
||__|||__|||__|||__|||__|||__||
|/__\|/__\|/__\|/__\|/__\|/__\|
 ____ ____ ____ ____ ____ ____
||R |||i |||g |||g |||e |||r ||
||__|||__|||__|||__|||__|||__||
|/__\|/__\|/__\|/__\|/__\|/__\|

V1.2

Hey All, thank you for your interest in my script.
My name is John and I have a site called nerdblog.io

Improvements to come:
-Multiple sculpts with different quantities & prices
-Nerdblog.io auto win

The idea is you have a google form for your raffle, with the fields PaypalEmail, FirstName, LastName, Address,
and a yes or no question if they're international. You would then export your results from your google sheets
into an excel spreadsheet with the following cell values, A1 = "PaypalEmail", B1 = "FirstName", C1 = "LastName", D1 = "International",
E1 = "Address". Save the excel file as a .csv (you can also do it with notepad, but have to separate with commas.)
The Data object has a few fields for your paypal invoicing that you can write in. Should be at line # 116
#>

#This is used to add to paypal memos, "Thank you for supporting ______!"
$ArtisanID = Read-Host "What is your artisan handle?"

#This is the amount of people to win
$AmountOfWinners = Read-Host "How many individual winners?"

#This is the file location of your raffle entries.
$RaffleEntriesCsv = ".\RaffleEntries.csv"

#This is the location of where it will save the text file.
$WinnerPaypalCsv = ".\Winners.csv"
#Banned users.
$BlacklistCsv = ".\Blacklist.csv"
#Imports blacklist
$BlackListUsers = Import-Csv $BlacklistCsv | sort Address -Unique


#Prompt user for necessary information (paypal invoicing).
$KeycapName = Read-Host "What is the keycap name you are raffling?"
$Keycapvalue = Read-Host "How much are you selling the $KeycapName for?"

#The line below imports the entries and removes duplicates based on addresses.
$EntryList = Import-Csv $RaffleEntriesCsv | sort Address -Unique

#Declare arrays | Variables
$PaypalEmail = @()
$International = @()
$Firstname = @()
$LastName = @()
$Address = @()
$Entry = @()
$ShippingCost = ""
$Data = ""

<#Cycles through each raffle entry, compares it against the blacklist and stores
their paypal email, along with if they're international or not.
It also generates a random number 1-1000 and stores the combined points in a new object. #>
foreach ($u in $EntryList) {
	if ($BlackListUsers | Where-Object { $_.Address -eq $($u.Address) }) {
		Write-Host "" $u.Address " is banned"
	} else {
		$PaypalEmail = $($u.PaypalEmail)
		$International = $($u.International)
		$FirstName = $($u.FirstName)
		$LastName = $($u.LastName)
		$Address = $($u.Address)
		$CombinedPoints = Get-Random 1000

		$Entry += New-Object PSObject -Property @{
			PaypalEmail = $PaypalEmail
			International = $International
			FirstName = $Firstname
			LastName = $LastName
			Address = $Address
			Points = $Points
			CombinedPoints = [int]$CombinedPoints
		}
	}
}

#Write winners to new object
$Winners = $Entry | Sort-Object -Property CombinedPoints -Descending | Select-Object -First $AmountOfWinners

#Imports the Winners csv file.
if (Test-Path $WinnerPaypalCsv) {
	[array]$Data = Import-Csv -Path $WinnerPaypalCsv
}

#Export winners to Paypal csv
foreach ($w in $Winners) {

	if ($w.International -eq 'Yes') {
		$ShippingCost = '13.00'
	} else {
		$ShippingCost = '5.00'
	}

	$Data += New-Object PSObject -Property @{
		'Recipient Email' = $w.PaypalEmail
		'Recipient First Name' = $w.FirstName
		'Recipient Last Name' = $w.LastName
		'Invoice Number' = ""
		'Due Date' = ""
		'Reference' = ""
		'Item Name' = $KeycapName
		'Description' = ""
		'Item Amount' = $Keycapvalue
		'Shipping Amount' = $ShippingCost
		'Discount' = ""
		'Currency' = "USD"
		'Note to Customer' = "Thank you for supporting $ArtisanID!"
		'Terms and Conditions' = ""
		'Memo to Self' = ""
	}

	<#
Adds the winners to csv file. / the properties below are what gets exported, while only the necessary columns
were added, you could hard code in due date, or Currency to USD... etc
#>

	$Data | Select-Object "Recipient Email","Recipient First Name","Recipient Last Name","Invoice Number","Due Date","Reference",
	"Item Name","Description","Item Amount","Shipping Amount","Discount","Currency","Note to Customer","Terms and Conditions","Memo to Self" | Export-Csv -Path $WinnerPaypalCsv -NoTypeInformation
}

#One second delay
Start-Sleep -s 1

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

#Wait 1 second
Start-Sleep -s 1
