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

V2.00

Hey All, thank you for your interest in my script.
My name is John and I have a site called nerdblog.io

Improvements to come:
-Nerdblog.io auto win

The idea is you have a google form for your raffle, with the fields Paypal Email, First Name, Last Name, Address,
and a yes or no question if they're international.
If you're doing a multiple sculpt, it would be on a Multiple Checkbox with the names of the thing listed.
You would then export your results from your google sheets
into an excel spreadsheet with the following cell values, A1 = "Paypal Email", B1 = "First Name", C1 = "Last Name", D1 = "International",
E1 = "Address". Save the excel file as a .csv (you can also do it with notepad, but have to separate with commas.)
#>

#This is used to add to paypal memos, "Thank you for supporting ______!"
$ArtisanID = Read-Host "What is your artisan handle?"

#Prompt user for raffle style
Write-Host "Menu:
1) Blind Raffle
2) Multiple Sculpts/Colorways"

#This is the file location of your raffle entries.
$RaffleEntriesCsv = ".\RaffleEntries.csv"

#This is the location of where it will save the text file.
$WinnerPaypalCsv = ".\Winners.csv"
#Banned users.
$BlacklistCsv = ".\Blacklist.csv"
#Imports blacklist
$BlackListUsers = Import-Csv $BlacklistCsv

#The line below imports the entries and removes duplicates based on addresses.
$EntryList = Import-Csv $RaffleEntriesCsv | sort Address -Unique

do
{
	$RaffleType = Read-Host -Prompt "Please select an option: (1 or 2)"
}
while ($RaffleType -notlike 1 -and $RaffleType -notlike 2)



switch ($RaffleType) {
	1 {
		#This is the amount of people to win
		$AmountOfWinners = Read-Host "How many individual winners?"

		#Prompt user for necessary information (paypal invoicing).
		$KeycapName = Read-Host "What is the keycap name you are raffling?"
		$Keycapvalue = Read-Host "How much are you selling the $KeycapName for?"
		$IntShipCost = Read-Host "How much are you charging for international shipping?"
		$ConusShipCost = Read-Host "How much are you charging for conus shipping?"
		[int]$InvoiceNumber = Read-Host "What is your last Paypal invoice number?"

		#Declare arrays | Variables
		$ErrorActionPreference = "SilentlyContinue"
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
			$Check = $($u.Address.ToLower()) -replace "\W"
			if ($BlackListUsers | Where-Object { ($($_.Address.ToLower()) -replace "\W") -eq $Check -or $_. "Paypal Email".Split('+@')[0].ToLower() -eq $($u. "Paypal Email").Split('+@')[0].ToLower() }) {
				Write-Host "" $u . "Paypal Email" " is banned"
			} else {
				$PaypalEmail = $($u. "Paypal Email")
				$International = $($u.International)
				$FirstName = $($u. "First Name")
				$LastName = $($u. "Last Name")
				$Address = $($u.Address)
				$CombinedPoints = Get-Random 1000

				$Entry += New-Object PSObject -Property @{
					"Paypal Email" = $PaypalEmail
					International = $International
					"First Name" = $Firstname
					"Last Name" = $LastName
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
			$InvoiceNumber++
			if ($w.International -eq 'Yes') {
				$ShippingCost = $IntShipCost
			} else {
				$ShippingCost = $ConusShipCost
			}

			$Data += New-Object PSObject -Property @{
				'Recipient Email' = $w. "Paypal Email"
				'Recipient First Name' = $w. "First Name"
				'Recipient Last Name' = $w. "Last Name"
				'Invoice Number' = $InvoiceNumber
				'Due Date' = ""
				'Reference' = ""
				'Item Name' = $KeycapName
				'Description' = ""
				'Item Amount' = $Keycapvalue
				'Shipping Amount' = $ShippingCost
				'Discount' = ""
				'Currency' = "USD"
				'Note to Customer' = "Thank you for supporting $ArtisanID!"
				'Terms and Conditions' = "These keycaps are handmade. There is individual variation in shape, color, and pattern due to the nature of the creation process. By purchasing, you acknowledge that the keycap(s) you receive will not look exactly like the one pictured unless stated otherwise. This disclaimer does not cover defects in manufacturing."
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



	}
	2 {
		#This is the amount of people to win
		$AmountOfKeycaps = Read-Host "How many Individual Sculpts/Colorways?"
		#Prompt user for necessary information (paypal invoicing).
		$IntShipCost = Read-Host "How much are you charging for international shipping?"
		$ConusShipCost = Read-Host "How much are you charging for conus shipping?"
		$InvoiceNumber = Read-Host "What is your last Paypal invoice number?"

		for ($j = 0; $j -lt $AmountOfKeycaps; $j++) {
			$number = $j + 1
			$KeycapName = Read-Host "What is keycap/colorway name you are raffling for number $number ?"
			$AmountOfWinners = Read-Host "How many individual winners for this keycap?"
			$Keycapvalue = Read-Host "How much are you selling the $KeycapName for?"

			#Declare arrays | Variables
			$ErrorActionPreference = "SilentlyContinue"
			$PaypalEmail = @()
			$International = @()
			$Firstname = @()
			$LastName = @()
			$Address = @()
			$Entry = @()
			$Data = ""

			<#Cycles through each raffle entry, compares it against the blacklist and stores
	their paypal email, along with if they're international or not.
	It also generates a random number 1-1000 and stores the combined points in a new object. #>

			foreach ($u in $EntryList) {
				if ($($u. "Keycap [$KeycapName]") -eq "Yes") {
					$Check = $($u.Address.ToLower()) -replace "\W"
					if ($BlackListUsers | Where-Object { ($($_.Address.ToLower()) -replace "\W") -eq $Check -or $_. "Paypal Email".Split('+@')[0].ToLower() -eq $($u. "Paypal Email").Split('+@')[0].ToLower() }) {
						Write-Host "" $u . "Paypal Email" " is banned"
					} else {
						$PaypalEmail = $($u. "Paypal Email")
						$International = $($u.International)
						$FirstName = $($u. "First Name")
						$LastName = $($u. "Last Name")
						$Address = $($u.Address)
						$CombinedPoints = Get-Random 1000

						$Entry += New-Object PSObject -Property @{
							"Paypal Email" = $PaypalEmail
							International = $International
							"First Name" = $Firstname
							"Last Name" = $LastName
							Address = $Address
							Points = $Points
							CombinedPoints = [int]$CombinedPoints
						}
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
				$InvoiceNumber++
				if ($w.International -eq 'Yes') {
					$ShippingCost = $IntShipCost
				} else {
					$ShippingCost = $ConusShipCost
				}

				$Data += New-Object PSObject -Property @{
					'Recipient Email' = $w. "Paypal Email"
					'Recipient First Name' = $w. "First Name"
					'Recipient Last Name' = $w. "Last Name"
					'Invoice Number' = $InvoiceNumber
					'Due Date' = ""
					'Reference' = ""
					'Item Name' = $KeycapName
					'Description' = ""
					'Item Amount' = $Keycapvalue
					'Shipping Amount' = $ShippingCost
					'Discount' = ""
					'Currency' = "USD"
					'Note to Customer' = "Thank you for supporting $ArtisanID!"
					'Terms and Conditions' = "These keycaps are handmade. There is individual variation in shape, color, and pattern due to the nature of the creation process. By purchasing, you acknowledge that the keycap(s) you receive will not look exactly like the one pictured unless stated otherwise. This disclaimer does not cover defects in manufacturing."
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

		}

	}




}

Write-Host "Thank you for supporting

    _   __              ____    __              _
   / | / /__  _________/ / /_  / /___  ____ _  (_)___
  /  |/ / _ \/ ___/ __  / __ \/ / __ \/ __ ``/ / / __ \
 / /|  /  __/ /  / /_/ / /_/ / / /_/ / /_/ / / / /_/ /
/_/ |_/\___/_/   \__,_/_.___/_/\____/\__, (_)_/\____/
                                    /____/
"

Start-Sleep -Seconds 5

exit
