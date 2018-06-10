![Nerdblog.io](https://i.imgur.com/Lborxa2.png)

# RaffleRigger
Script for Importing Raffle Entries, choosing winners, and exporting into a csv file for paypal batch invoice upload.

Paypal only needs columns for their csv, Email,item,item amount, and shipping amount.
With that in mind, 

If it is your first time running a powershell script:
1. Open Powershell as administrator
2. Type the following command: Set-ExecutionPolicy Unrestricted
3. Say yes to all.

To run raffle rigger:
1. Download all csv files into the same directory.
2. Add banned user's addresses to the blacklist file.
3. Edit RaffleEntries.csv with your Raffle Entries, keep the headers the same.
Run Raffle Rigger.ps1; enter how many keys you are giving away.

It will create a new csv file called "Winners.csv"

you can upload the file at https://www.paypal.com/invoice/batch

![Nerdblog.io](https://i.imgur.com/K20uFGm.png)
