#Created by: Jacob Carl
#Editted to be used for personal use

#Needed for creating form dropbox
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Version Number
$Global:VERNUM = "1.4.2"

#Global variable for the SQL Server path
#=========================LIVE==================================
$Global:SCCMSQLSERVER = 1 #Placeholder: path to the server holding the database
#===============================================================

#=========================TEST==================================
#$Global:SCCMSQLSERVER = "localhost\SQLEXPRESS"
#===============================================================

#Global variable for the Database name
$Global:DBNAME = "WiseTrack"

#Global variable for CSV file for multiple asset add
$Global:CSVFILE = 1 #Placeholder for excel sheet


#Main Menu Display Function
function Main-Menu
{
    param(
        [string]$Title = 'WiseTrack Home'
    )
    Write-Host "==================$Title================="
    Write-Host " "
    Write-Host " Input    |       Action"
    Write-Host "----------------------------"
    Write-Host "A)dd      :  Go to Add menu"
    Write-Host "U)pdate   :  Update an Asset"
    Write-Host "S)earch   :  Search for Asset"
    Write-Host "L)ogging  :  Return logs of TSC Asset"
    Write-Host "P)rofile  :  Update Profile Settings"
    Write-Host "H)elp     :  Look at About & Features"
    Write-Host "Q)uit     :  Press 'Q' to quit"

}

#Add Menu Display Function
function Add-Menu
{
    param(
        [string]$Title = 'Add Home'
    )
    Write-Host "==================$Title================="
    Write-Host " "
    Write-Host " Input   |       Action"
    Write-Host "---------------------------------"
    Write-Host "A)sset   :  Create a new Asset"
    Write-Host "I)nvoice :  Create a new Invoice"
    Write-Host "U)ser    :  Create a new User"
    Write-Host "C)ancel  :  Go back to Home Menu"
}

#Search Menu Display Function
function Search-Menu
{
    param(
        [string]$Title = 'Search Home'
    )
    Write-Host "==================$Title================="
    Write-Host " "
    Write-Host " Input   |       Action"
    Write-Host "---------------------------------"
    Write-Host "A)sset   :  Search by TSC Number"
    Write-Host "U)ser    :  Search by User"
    Write-Host "S)erial  :  Search by Serial Number"
    Write-Host "C)ancel  :  Go back to Home Menu"
}

#Update Menu Display Function
function Update-Menu
{
    param(
        [string]$Title = 'Update Home'
    )
    Write-Host "==================$Title================="
    Write-Host " "
    Write-Host " Input   |       Action"
    Write-Host "---------------------------------"
    Write-Host "A)sset   :  Update an Asset"
    Write-Host "U)ser    :  Update a User"
    Write-Host "C)ancel  :  Go back to Home Menu"
}

function Display-About
{
    Write-Host "This program was created by Jacob Carl and is property of TechSmith Corporation."
    Write-Host "Version: $VERNUM"
    Write-Host "=============================================================================================="
    Write-Host ""
    Write-Host "Add:"
    Write-Host ""
    Write-Host "    -Asset:"
    Write-Host "     -Single    This function is used to create a new asset in the database. This"
    Write-Host "                function automates a lot of redundant input, resulting in faster"
    Write-Host "                input. As of Version 2.0.0, the script will add any asset over `$1000"
    Write-Host "                to the excel sheet hosted on the IT share, which then gets put into"
    Write-Host "                shared excel file for accounting by a daily job."
    Write-Host "     -Multiple  This function is used to create multiple assets in the database."
    Write-Host "                The data is pulled from:"
    Write-Host "                \\techsmith.com\departments\it\Scripts\WiseTrack\Staging\BatchAssetAdd.csv"
    Write-Host ""
    Write-Host "    -Invoice:   This function is used to create a new invoice in the database. Like"
    Write-Host "                the Asset function, this automates a lot of redundant input."
    Write-Host ""
    Write-Host "    -User:      This function is used to create a new user in the database."
    Write-Host ""
    Write-Host "Search:"
    Write-Host ""
    Write-Host "    -Asset:     This function is used to search for an asset if you know the TSC Number"
    Write-Host "                As of Version 3.1.1, you can now search by part of a TSC Number (i.e TSC-15)"
    Write-Host ""
    Write-Host "    -User:      This function is used to return a list of assets that belong to a"
    Write-Host "                specific user. In the case of duplicate last names, it will allow you"
    Write-Host "                to choose which user you are looking for."
    Write-Host ""
    Write-Host "    -Serial:    This function is used to search for an asset by its serial number in"
    Write-Host "                the case that there is no TSC Number on the computer."
    Write-Host ""
    Write-Host "    -NOTE:      As of Version 2.1.0, you can now update an asset straight from the search."
    Write-Host ""
    Write-Host "Update:"
    Write-Host ""
    Write-Host "    -Asset:     This allows you to update the TSC Number, Model Name, Serial Number,"
    Write-Host "                User, Warranty Expire Date, Invoice Number, and Purchase Price."
    Write-Host ""
    Write-Host "    -User:      This allows you to update a user's first and last name."
    Write-Host ""
    Write-Host "Logging:        This program also takes care of logging in the database. In the instance"
    Write-Host "                that an asset has been changed, and you need to find the previous value,"
    Write-Host "                you can connect to the database and view the table LOGTRANSACTION. Logging"
    Write-Host "                is done for any add or update to an asset."
    Write-Host "                As of Version 3.2.0, you can now search for the logs of a specific TSC asset"
    Write-Host ""
    Write-Host "Profile:        As of Version 4.0.0, you can now save settings that help you see what you"
    Write-Host "                want to see. Right now, there are only settings for search views, those being"
    Write-Host "                simple and extended."
    Write-Host ""
    Write-Host "    -Simple:    Use this setting if you only want to see info quick"
    Write-Host ""
    Write-Host "    -Extended:  Use this setting if you need to look at financial/other info on computer"
    Write-Host ""
    Write-Host "=============================================================================================="
    Write-Host ""
    Write-Host "Find some bugs that need squashing? Can't seem to do what you need? Put in a Feature Request!"
    Write-Host "Go to https://docs.google.com/forms/d/1qyfmlx6Y40nQMSQnsnyI8AVHBmNaBFusYywCYi99daI and submit"
    Write-Host "your request. We can't make this better without you!"
    Write-Host ""
    pause
}

function Check-AD
{
    #Verify $UserCredential
    $credentialCheck = "False"
    $credentialCounter = 0

    Do{
        Try
        {
            $Global:UserCredential = Get-Credential
        }
        Catch
        {
            exit
        }
        $username = $UserCredential.username
        $password = $UserCredential.GetNetworkCredential().password

        # Get current domain using logged-on user's credentials
        $currentDomain = "LDAP://" + ([ADSI]"").distinguishedName
        $domain = New-Object System.DirectoryServices.DirectoryEntry($currentDomain,$userName,$password)

        if ($domain.name -eq $null){
             write-host "Authentication failed - please verify your username and password." -BackgroundColor White -ForegroundColor Red
             $credentialCounter++
             #exit #terminate the script.
        }
        else{
            write-host "Successfully authenticated with domain $domain.name" -BackgroundColor White -ForegroundColor Blue
            #Global variable for the User logged on
            #CURRENTLY STATIC DURING PROOF OF CONCEPT
            $Global:LOGGEDINUSER = $username
            $Global:SECUREPASS = $password

            #Check user to see if it is a part of DL - IT Staff
            $Userlist = Get-ADGroupMember -Identity "ITStaff"
            foreach($Item in $Userlist)
            {
                if($username -eq $Item.SamAccountName){
                    Set-Variable -Name CredentialCheck -Value "True" -Scope 0
                    Write-Host "User Verified!"
                }
            }
            if($credentialCheck -eq "False")
            {
                Write-Host "Sorry, the user does not have correct permissions. Please try a different user account"
            }
        }
    } While (($credentialCheck -eq "False") -and ($credentialCounter -lt 5))

}
function Initialize-Profile
{
    $text = "search:simple"
    $text | Add-Content 'C:\WiseTrackProfile\WiseTrackProfile.txt'
}

function Set-Profile
{
    $Global:ProfileSearch = Get-Content -Path C:\WiseTrackProfile\WiseTrackProfile.txt | Where-Object {$_ -like 'search:*'}
    $Global:ProfileSearch = $Global:ProfileSearch.Replace("search:","")
}

#This function will be used to create the profile directory if one is not already made
function Profile-Check
{
    if(!(Test-Path -Path "C:\WiseTrackProfile"))
    {
        $response = Read-Host "You do not seem to have a profile yet. Would you like to create one?(Y/N)"
        $response = $response.ToLower()
        switch($response){
            'y'{
                Try{
                    [System.IO.Directory]::CreateDirectory("C:\WiseTrackProfile")
                }
                Catch{
                    Write-Host "Unable to make Directory. Please make sure you have sufficient permissions"
                    return $false
                }
                Try{
                    New-Item C:\WiseTrackProfile\WiseTrackProfile.txt -ItemType file
                }
                Catch{
                    Write-Host "Unable to make profile. Please make sure you have sufficient permissions"
                    Remove-Item -Path C:\WiseTrackProfile
                    return $false
                }
                #Initializes the profile with default settings
                Initialize-Profile
                Write-Host "Profile successfully created"
                pause
                Set-Profile
                return $true
            }
        }
        return $false
    }
    Set-Profile
    return $true
}

function Update-Profile
{
    $search = Get-Content -Path C:\WiseTrackProfile\WiseTrackProfile.txt | Where-Object {$_ -like 'search:*'}
    $search = $search.Replace("search:","")

    #Create Form with objects in it
    #Grabbing the Vendor
    #Taken from https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/selecting-items-from-a-list-box?view=powershell-6
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Select a Vendor"
    $form.Size = New-Object System.Drawing.Size(600,600) 
    $form.StartPosition = "CenterScreen"

    #Create OK button on form
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(150,400)
    $OKButton.Size = New-Object System.Drawing.Size(100,35)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    #Create Cancel button on form
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(275,400)
    $CancelButton.Size = New-Object System.Drawing.Size(100,35)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    #Create the label on the Form
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(400,40) 
    $label.Text = "Search Settings:"
    $form.Controls.Add($label)
    
    $SearchDropDown = new-object System.Windows.Forms.ComboBox
    $SearchDropDown.Location = new-object System.Drawing.Size(40,60)
    $SearchDropDown.Size = new-object System.Drawing.Size(260,20)

    [void] $SearchDropDown.Items.Add("simple")
    [void] $SearchDropDown.Items.Add("extended")

    $form.Controls.Add($SearchDropDown)

    #Makes the current setting the selected setting in the drop box
    if($search -eq "extended")
    {
        $SearchDropDown.SelectedIndex = 1
    }
    else
    {
        $SearchDropDown.SelectedIndex = 0
    }

    $form.Controls.Add($listBox) 

    $form.Topmost = $True

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $search = $SearchDropDown.SelectedItem
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        return
    }
    
    $text = "search:$search"
    $text | Out-File 'C:\WiseTrackProfile\WiseTrackProfile.txt'
    Set-Profile
}

#---------------------------------------------------------------
#Function that executes adding multiple assets to the Database
#---------------------------------------------------------------
function Return-Logs
{
    $TSCNum = Read-Host "Please enter TSC Number (TSC-####)"

    #Create a SQL Command Object
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand

    #Create a SQL Adapter object to allow for the execution of the SQL command
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SQLCommand.Connection = $SQLConnection

    #Grab the TSC Number ID in Database
    $SQLCommand.CommandText = "select DATESTAMP, USER_NAME, LOGTRANSACTION.DESCRIPTION from LOGTRANSACTION left join ITEM on ITEM.ID=LOGTRANSACTION.ITEM_ID where ITEM.DESCRIPTION = '$TSCNum'"
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null
    $TSCTable = $SQLDataset.Tables[0]

    if($TSCTable.Rows.Count -eq 0)
    {
        Write-Host "Asset not found/No logs for asset"
    }
    else
    {
        foreach($data in $TSCTable)
        {
            $data | Format-List
        }
    }
    $Response = Read-Host "Find another TSC Asset's logs?(y/n)"
    $Response = $Response.ToLower()
    if($Response -eq 'y')
    {
        Return-Logs
    }
}

#---------------------------------------------------------------
#Function that executes adding multiple assets to the Database
#---------------------------------------------------------------
function Add-Asset-Batch
{
    try
    {
        $Assets = import-csv $CSVFILE
    }
    catch [System.Exception]
    {
        Write-Error "Could not open CSV file for batch add"
        pause
        return
    }

    #Create a SQL Command Object
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand

    #Create a SQL Adapter object to allow for the execution of the SQL command
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SQLCommand.Connection = $SQLConnection

    #Grab all the Companies
    $SQLCommand.CommandText = "select ID, COMPANY from [dbo].COMPANY"
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null
    $Companies = $SQLDataset.Tables[0]

    #Grab all the Invoices
    $SQLCommand.CommandText = "select INVOICENUM, ID from [dbo].INVOICE order by ID asc"
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null

    $Invoices = $SQLDataset.Tables[0]

    $ErrorFlag = $false
    $AssetCount = 2
    foreach($Asset in $Assets)
    {
        #Check for NULL values
        #Returns if there is an issue with an asset
        if($Asset.TSCNumber -eq "") {$ErrorFlag = $true}
        if($Asset.ModelName -eq "") {$ErrorFlag = $true}
        if($Asset.ModelNumber -eq "") {$ErrorFlag = $true}
        if($Asset.Serial -eq "") {$ErrorFlag = $true}
        if($Asset.Supplier -eq "") {$ErrorFlag = $true}
        if($Asset.Manufacturer -eq "") {$ErrorFlag = $true}
        if($Asset.WarrantyExpireDate -eq "") {$ErrorFlag = $true}
        if($Asset.Invoice -eq "") {$ErrorFlag = $true}
        if($Asset.PurchasePrice -eq "") {$ErrorFlag = $true}
        if($Asset.PurchaseDate -eq "") {$ErrorFlag = $true}
        if($ErrorFlag)
        {
            Write-Host "Error at row $AssetCount. Please delete rows above row $AssetCount and fix error. DO NOT DELETE ROW 1"
            pause
            return
        }

        #Setting values to local variables
        $TSCNum = $Asset.TSCNumber
        $ModelName = $Asset.ModelName
        $ModelNum = $Asset.ModelNumber
        $Serial = $Asset.Serial
        $SupplierName = '*'+$Asset.Supplier+'*'
        $ManufacturerName = '*'+$Asset.Manufacturer+'*'

        #Makes Supplier and Manufacturer readable by database
        $Supflag = $true
        $Manflag = $true
        foreach($Item in $Companies)
        {
            if($Item[1] -like $SupplierName)
            {
                $Supplier = $Item[0]
                $Supflag = $false
            }
            if($Item[1] -like $ManufacturerName)
            {
                $Manufacturer = $Item[0]
                $Manflag = $false
            }
        }
        if($Supflag)
        {
            $Supplier = '1038'
        }
        if($Manflag)
        {
            $Manufacturer = '14'
        }

        $Warranty = $Asset.WarrantyExpireDate
        $InvoiceNumber = $Asset.Invoice
        $Invflag = $true
        #Setting invoice to be readable by Database
        foreach($Item in $Invoices)
        {
            if($Item[0] -eq $InvoiceNumber)
            {
                $Invoice = $Item[1]
                $Invflag = $false
            }
        }
        if($Invflag)
        {
            Write-Host "Invoice not found. Process stopped at row $AssetCount. Please delete rows above $AssetCount. DO NOT DELETE ROW 1."
            pause
            return
        }

        $PurchasePrice = $Asset.PurchasePrice
        $PurchaseDate = $Asset.PurchaseDate
        $date = (Get-Date).ToString("MM/dd/yyyy")
        $ServiceExpire = ((Get-Date).AddYears(3)).ToString("MM/dd/yyyy")

        #Adding the object to Wisetrack
        $SQLCommand.CommandText = "select DESCRIPTION from [dbo].ITEM where DESCRIPTION = '$TSCNum'"
        $SQLAdapter.SelectCommand = $SQLCommand
        $SQLDataset = New-Object System.Data.DataSet 
        $SQLAdapter.fill($SQLDataset) | Out-Null

        $TSCTable = $SQLDataset.Tables[0]


        if($TSCTable.Rows.Count -eq 0)
        {
            #Create the SQL Statement and Connection
            #NOTE: $SQLConnection is the Connection to the database established at the top of Main Program in first Try statement
            $SQLCommand.CommandText = "insert into [dbo].ITEM 
                                        (BARCODE, DESCRIPTION, SERIAL_NUMBER, MODEL_NUMBER, SUPPLIER_ID, MANUFACTURER_ID, LOCATION_ID, CUSTODIAN_NAME, INVOICE_ID, IN_SERVICE_DATE, DATE_ENTERED, ESTIMATED_LIFE, WARRANTY_EXPIRY_DATE, DATE_LAST_INVENTORIED,
                                        SERVICE_EXPIRY_DATE, SURPLUS, LOST, PURCHASE_PRICE, PURCHASE_DATE, SOLD, LEASED, BUSINESS_USE_PERCENT, DELETED, DYNAMIC, TAGID, INTERNAL_DEPRECIATED, MODEL_NAME)
                                        values 
                                        ('$TSCNum', '$TSCNum', '$Serial', '$ModelNum', '$Supplier', '$Manufacturer', '23', 'ITStockroom', '$Invoice', '$date', '$date', '36', '$Warranty', '$date', '$ServiceExpire', '0', '0', '$PurchasePrice', '$PurchaseDate', '0', '0', '100', '0', '0', '$TSCNum', '0', '$ModelName')"
            $SQLCommand.Connection = $SQLConnection

            #Creates SQL Adapter Object to execute the Insert Function
            $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
            $SQLCommand.ExecuteNonQuery()
            
            $SQLCommand.CommandText = "select ID from [dbo].ITEM where DESCRIPTION='$TSCNum' ORDER BY ID asc"
            $SQLAdapter.SelectCommand = $SQLCommand
            $SQLDataset = New-Object System.Data.DataSet 
            $SQLAdapter.fill($SQLDataset) | out-null

            #Grab the Asset ID
            foreach($Item in $SQLDataset.Tables[0])
            {
                $AssetID = $Item[0]
            }

            $UpdateDescrip = "BarCode:$TSCNum;TagID:$TSCNum;Description:$TSCNum;Custodian:ITStockroom;Serial Number:$Serial;Model Number:$ModelNum;Model Name:$ModelName;Manufacturer Name:"+$ManufacturerName+";Invoice:"+$InvoiceNumber+";Date Last Inventoried:$date;Purchase Price:$PurchasePrice;Estimated Life:36;Warranty Expiry Date:$Warranty"
            $LogObject = @{description = $UpdateDescrip;
                           type = 1;
                           ItemID = $AssetID
                          }

            $LogDate= Logging($LogObject)
            $PurchasePrice = [int]$PurchasePrice

            if($PurchasePrice -ge 1000)
            {
                #Write to Sharepoint Excel File
                #Code from: https://stackoverflow.com/questions/35606762/append-powershell-output-to-an-excel-file
                #Launch Excel
                $XL = New-Object -ComObject Excel.Application
                #Open the workbook
                $WB = $XL.Workbooks.Open("\\techsmith.com\departments\it\scripts\wisetrack\staging\UpdateSharepointAsset.xlsx")
                #Activate Sheet1, pipe to Out-Null to avoid 'True' output to screen
                $WB.Sheets.Item("Sheet1").Activate() | Out-Null
                #Find first blank row #, and activate the first cell in that row
                $FirstBlankRow = $($xl.ActiveSheet.UsedRange.Rows)[-1].Row + 1
                $XL.ActiveSheet.Range("A$FirstBlankRow").Activate()
                #Create PSObject with the properties that we want, convert it to a tab delimited CSV, and copy it to the clipboard
                $Record = [PSCustomObject]@{
                            'TSCNum' = $TSCNum
                            'Serial' = $Serial
                            'ModelName' = $ModelName
                            'Date' = $date
                            'Price' = $PurchasePrice
                            'Supplier' = $SupplierName.trim('*')
                            'Invoice' = $InvoiceNumber
                }
                $Record | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Select -Skip 1 | Clip
                #Paste at the currently active cell
                $XL.ActiveSheet.Paste() | Out-Null
                # Save and close
                $WB.Save() | Out-Null
                $WB.Close() | Out-Null
                $XL.Quit() | Out-Null
                #Release ComObject
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($XL)
            }
        }
        else
        {
            Write-Host "TSC Number Already exists (Row $AssetCount). Skipping Asset."
            pause 
        }
        $AssetCount = $AssetCount + 1
    }
}



#----------------------------------------------------------
#Function that executes adding a new user to the Database
#----------------------------------------------------------
function Add-User
{
    #Gather input from the User
    [string]$username = Read-Host "Username"
    [string]$first = Read-Host "First Name"
    [string]$last = Read-Host "Last Name"
    $date = (Get-Date).ToString('MM/dd/yyyy')

    #Create a SQL Command Object
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand

    #Create the SQL Statement and Connection
    #NOTE: $SQLConnection is the Connection to the database established at the top of Main Program in first Try statement
    $SQLCommand.CommandText = "insert into [dbo].USERS (USER_NAME, FIRST_NAME, LAST_NAME, DATE_ENTERED) values ('$username', '$first', '$last','$date');"
    $SQLCommand.Connection = $SQLConnection

    #Creates SQL Adapter Object to execute the Insert Function
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    try
    {
        $SQLCommand.ExecuteNonQuery()                   
    }
    catch
    {
        [System.Windows.MessageBox]::Show('User was not added to Database', 'Error', 'OK', 'Error')
    }
    #pause
    #----------------------------------------------------------------------------------
    #Comment Block (BELOW CODE IS FOR TESTING ONLY):
    #Lines below are used to check that the user has been input into the database 
    #----------------------------------------------------------------------------------
    #$SQLCommand.CommandText = "select * from [dbo].USERS where FIRST_NAME like '$first'"
    #$SQLCommand.Connection = $SQLConnection

    #$SQLAdapter.SelectCommand = $SQLCommand
    #$SQLDataset = New-Object System.Data.DataSet 
    #$SQLAdapter.fill($SQLDataset) | out-null
    #$tablevalue = @()
    #foreach ($data in $SQLDataset.tables[0])
    #{
    #    $tablevalue = $data[0]
    #    $tablevalue
    #}
    #pause
}

#----------------------------------------------------------
#Function that executes adding a new invoice to the Database
#----------------------------------------------------------
function Add-Invoice
{

    #Grabs the Invoice Number, Date, Subtotal, Shipping, Tax,
    #and computes the total
    [int]$number = Read-Host "Invoice Number"
    $date = (Get-Date).ToString('MM/dd/yyyy')
    [int]$subtotal = Read-Host "Subtotal Cost"
    [int]$shipping = Read-Host "Shipping Cost"
    [int]$tax = Read-Host "Tax Cost"
    [int]$total = $subtotal + $shipping + $tax

    #Create Form with objects in it
    #Grabbing the Vendor
    #Taken from https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/selecting-items-from-a-list-box?view=powershell-6
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Select a Vendor"
    $form.Size = New-Object System.Drawing.Size(600,600) 
    $form.StartPosition = "CenterScreen"

    #Create OK button on form
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(150,400)
    $OKButton.Size = New-Object System.Drawing.Size(100,35)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    #Create Cancel button on form
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(275,400)
    $CancelButton.Size = New-Object System.Drawing.Size(100,35)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    #Create the label on the Form
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(400,40) 
    $label.Text = "Please select a Vendor:"
    $form.Controls.Add($label)

    #Create the list object on the form
    $listBox = New-Object System.Windows.Forms.ListBox 
    $listBox.Location = New-Object System.Drawing.Point(40,70) 
    $listBox.Size = New-Object System.Drawing.Size(500,300) 
    $listBox.Height = 300

    #Create an updated list of Vendors to populate list box
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
    $SQLCommand.CommandText = "select * from [dbo].COMPANY order by COMPANY asc"
    $SQLCommand.Connection = $SQLConnection

    #Create a SQL Adapter object to allow for the execution of the SQL command
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter

    #Execute the SQL command and grab the resulting query
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null

    #Grabbing the first table in the results
    $table = $SQLDataset.tables[0]

    #Set-up list to display in box
    foreach($data in $table)
    {
        $datadisplay = $data[1]
        [void] $listBox.Items.Add("$datadisplay")
    }

    #Formatting to display the form
    $form.Controls.Add($listBox) 
    $form.Topmost = $True
    $result = $form.ShowDialog()

    #Grabbing the info that was selected
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $select = $listBox.SelectedItem
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        return
    }

    #Grab value database will understand
    foreach($data in $table)
    {
        $datadisplay = $data[1]
        if ($datadisplay -eq $select)
        {
            [int]$vendor = $data[0]
        }
    }



    #Create Form with objects in it
    #Grabbing the Supplier
    #Taken from https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/selecting-items-from-a-list-box?view=powershell-6
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Select a Supplier"
    $form.Size = New-Object System.Drawing.Size(600,600) 
    $form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(150,400)
    $OKButton.Size = New-Object System.Drawing.Size(100,35)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(275,400)
    $CancelButton.Size = New-Object System.Drawing.Size(100,35)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(400,40) 
    $label.Text = "Please select a Supplier:"
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox 
    $listBox.Location = New-Object System.Drawing.Point(40,70) 
    $listBox.Size = New-Object System.Drawing.Size(500,300) 
    $listBox.Height = 300

    #Create an updated list of Vendors to populate list box
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
    $SQLCommand.CommandText = "select * from [dbo].COMPANY order by COMPANY asc"
    $SQLCommand.Connection = $SQLConnection

    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter

    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null

    $table = $SQLDataset.tables[0]

    #Set-up list to be put in box
    foreach($data in $table)
    {
        $datadisplay = $data[1]
        [void] $listBox.Items.Add("$datadisplay")
    }

    $form.Controls.Add($listBox) 

    $form.Topmost = $True

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $select = $listBox.SelectedItem
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        return
    }

    #Get the value the database will understand
    foreach($data in $table)
    {
        $datadisplay = $data[1]
        if ($datadisplay -eq $select)
        {
            [int]$supplier = $data[0]
        }
    }

    #Create Form with objects in it
    #Grabbing the PO
    #Taken from https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/selecting-items-from-a-list-box?view=powershell-6
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Select a PO"
    $form.Size = New-Object System.Drawing.Size(600,600) 
    $form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(150,400)
    $OKButton.Size = New-Object System.Drawing.Size(100,35)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(275,400)
    $CancelButton.Size = New-Object System.Drawing.Size(100,35)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(400,40) 
    $label.Text = "Please select a PO:"
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox 
    $listBox.Location = New-Object System.Drawing.Point(40,70) 
    $listBox.Size = New-Object System.Drawing.Size(500,300) 
    $listBox.Height = 300

    #Create an updated list of Vendors to populate list box
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
    $SQLCommand.CommandText = "select * from [dbo].PO"
    $SQLCommand.Connection = $SQLConnection

    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter

    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null

    $table = $SQLDataset.tables[0]

    #Set-up list to be put in box
    foreach($data in $table)
    {
        $datadisplay = $data[1]
        [void] $listBox.Items.Add("$datadisplay")
    }

    $form.Controls.Add($listBox) 

    $form.Topmost = $True

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $select = $listBox.SelectedItem
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        return
    }

    #Get the value the database will understand
    foreach($data in $table)
    {
        $datadisplay = $data[1]
        if ($datadisplay -eq $select)
        {
            [int]$PO = $data[0]
        }
    }

    #Grab the ship date and description for the invoice
    $shipdate = Read-Host "Ship Date (MM/DD/YYYY)"
    $description = Read-Host "Please enter a description"
    

    #Create a SQL Command Object
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand

    #Create the SQL Statement and Connection
    #NOTE: $SQLConnection is the Connection to the database established at the top of Main Program in first Try statement
    $SQLCommand.CommandText = "insert into [dbo].INVOICE (INVOICENUM, DESCRIPTION, VENDOR, SUBTOTAL, SHIPPING, TAX1, TOTAL, SUPPLIER_ID, PO_ID, DATESTAMP, SHIPDATE) values ('$number', '$description', $vendor, $subtotal, $shipping, $tax, $total, $supplier, $PO, '$date', '$shipdate');"
    $SQLCommand.Connection = $SQLConnection

    #Creates SQL Adapter Object to execute the Insert Function
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter

    try
    {
        $SQLCommand.ExecuteNonQuery()
    }
    catch
    {
        [System.Windows.MessageBox]::Show('Invoice was not added to Database', 'Error', 'OK', 'Error')
    }
}

function Search-Asset-TSC
{
    $TSCNum = Read-Host "TSC Number (TSC-####)"

    #Create a SQL Command Object
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
    if($Global:ProfileSearch -eq "simple")
    {
        $SQLCommand.CommandText = "select [dbo].ITEM.DESCRIPTION, MODEL_NAME, SERIAL_NUMBER, [dbo].USERS.FIRST_NAME, [dbo].USERS.LAST_NAME, [dbo].LOCATION.NAME from [dbo].ITEM left join [dbo].USERS on [dbo].ITEM.CUSTODIAN_NAME=[dbo].USERS.USER_NAME left join [dbo].INVOICE on [dbo].ITEM.INVOICE_ID=[dbo].INVOICE.ID left join [dbo].LOCATION on [dbo].ITEM.LOCATION_ID=[dbo].LOCATION.ID where [dbo].ITEM.DESCRIPTION like '%$TSCNum%'"
    }
    else
    {
        $SQLCommand.CommandText = "select [dbo].ITEM.DESCRIPTION, MODEL_NAME, MODEL_NUMBER, SERIAL_NUMBER, [dbo].USERS.FIRST_NAME, [dbo].USERS.LAST_NAME, [dbo].ITEM.DATE_ENTERED, WARRANTY_EXPIRY_DATE, [dbo].INVOICE.INVOICENUM, PURCHASE_PRICE, [dbo].LOCATION.NAME from [dbo].ITEM left join [dbo].USERS on [dbo].ITEM.CUSTODIAN_NAME=[dbo].USERS.USER_NAME left join [dbo].INVOICE on [dbo].ITEM.INVOICE_ID=[dbo].INVOICE.ID left join [dbo].LOCATION on [dbo].ITEM.LOCATION_ID=[dbo].LOCATION.ID where [dbo].ITEM.DESCRIPTION like '%$TSCNum%'"
    }
    $SQLCommand.Connection = $SQLConnection

    #Create an adapter to execute a SQL command
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null
    $tablevalue = @()
    foreach ($data in $SQLDataset.tables[0])
    {
        $data | Format-List
    }

    $table = $SQLDataset.tables[0]

    if($table.Rows.Count -ne 0)
    {
        #Check to see if user would like to update asset
        $update = Read-Host "Would you like to update this asset? (Y/N)"
        $update = $update.ToLower()
        if($update -eq 'y')
        {
            if($table.Rows.Count -eq 1)
            {
                Update-Asset($TSCNum)
            }
            else
            {
                $TSCNum = Read-Host "Please specify which asset number (TSC-####)"
                Update-Asset($TSCNum)
            }        
        }
    }
    else
    {
        Write-Host "No Assets Found"
        pause
    }
}
function Easter-Egg
{
    #You found me! Have fun ;)
    iex (New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w")
}

function Search-Asset-Serial
{
    $Serial = Read-Host "Serial Number"
    
    #Create a SQL Command Object
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand

    if($Global:ProfileSearch -eq "simple")
    {
        $SQLCommand.CommandText = "select [dbo].ITEM.DESCRIPTION, MODEL_NAME, SERIAL_NUMBER, [dbo].USERS.FIRST_NAME, [dbo].USERS.LAST_NAME, [dbo].LOCATION.NAME from [dbo].ITEM left join [dbo].USERS on [dbo].ITEM.CUSTODIAN_NAME=[dbo].USERS.USER_NAME left join [dbo].INVOICE on [dbo].ITEM.INVOICE_ID=[dbo].INVOICE.ID left join [dbo].LOCATION on [dbo].ITEM.LOCATION_ID=[dbo].LOCATION.ID where [dbo].ITEM.SERIAL_NUMBER like '%$Serial%'"
    }
    else
    {
        $SQLCommand.CommandText = "select [dbo].ITEM.DESCRIPTION, MODEL_NAME, MODEL_NUMBER, SERIAL_NUMBER, [dbo].USERS.FIRST_NAME, [dbo].USERS.LAST_NAME, [dbo].ITEM.DATE_ENTERED, WARRANTY_EXPIRY_DATE, [dbo].INVOICE.INVOICENUM, PURCHASE_PRICE, [dbo].LOCATION.NAME from [dbo].ITEM left join [dbo].USERS on [dbo].ITEM.CUSTODIAN_NAME=[dbo].USERS.USER_NAME left join [dbo].INVOICE on [dbo].ITEM.INVOICE_ID=[dbo].INVOICE.ID left join [dbo].LOCATION on [dbo].ITEM.LOCATION_ID=[dbo].LOCATION.ID where [dbo].ITEM.SERIAL_NUMBER like '%$Serial%'"
    }
    $SQLCommand.Connection = $SQLConnection

    #Create an adapter to execute a SQL command
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null
    $tablevalue = @()
    $table = $SQLDataset.Tables[0]

    if($table.Rows.Count -eq 0)
    {
        Write-Host "No Assets found"
        pause
        return
    }
    foreach ($data in $SQLDataset.tables[0])
    {
        $data | Format-List
        $TSCNum = $data.DESCRIPTION
    }
    
    $table = $SQLDataset.Tables[0]
    #Check to see if user would like to update asset
    $update = Read-Host "Would you like to update this asset? (Y/N)"
    $update = $update.ToLower()
    if($update -eq 'y')
    {
        if($table.Rows.Count -eq 1)
        {
            Update-Asset($TSCNum)
        }
        else
        {
            $TSCNum = Read-Host "Please specify which asset number (TSC-####)"
            Update-Asset($TSCNum)
        }        
    }
    
}

function Search-Asset-User
{
    $Last = Read-Host "User's Last Name"

    #Create a SQL Command Object
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
    $SQLCommand.CommandText = "select * from [dbo].USERS WHERE LAST_NAME like '$Last'"
    $SQLCommand.Connection = $SQLConnection

    #Create an updated list of Vendors to populate list box
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null
    $tablevalue = @()

    $table = $SQLDataset.Tables[0]

    if($table.Rows.Count -eq 1)
    {
        foreach($data in $table)
        {
            $username = $data[0]
        }
    }
    elseif($table.Rows.Count -eq 0)
    {
        Write-Host "User does not exist"
        pause
        return
    }
    else
    {
        #Create Form with objects in it
        #Grabbing the PO
        #Taken from https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/selecting-items-from-a-list-box?view=powershell-6
        $form = New-Object System.Windows.Forms.Form 
        $form.Text = "Select a User"
        $form.Size = New-Object System.Drawing.Size(600,600) 
        $form.StartPosition = "CenterScreen"

        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(150,400)
        $OKButton.Size = New-Object System.Drawing.Size(100,35)
        $OKButton.Text = "OK"
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $OKButton
        $form.Controls.Add($OKButton)

        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(275,400)
        $CancelButton.Size = New-Object System.Drawing.Size(100,35)
        $CancelButton.Text = "Cancel"
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $CancelButton
        $form.Controls.Add($CancelButton)

        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20) 
        $label.Size = New-Object System.Drawing.Size(400,40) 
        $label.Text = "Please select a User:"
        $form.Controls.Add($label)

        $listBox = New-Object System.Windows.Forms.ListBox 
        $listBox.Location = New-Object System.Drawing.Point(40,70) 
        $listBox.Size = New-Object System.Drawing.Size(500,300) 
        $listBox.Height = 300

        foreach ($data in $SQLDataset.tables[0])
        {
                $firstname = $data[5]
                $lastname = $data[4]
                [void] $listBox.Items.Add("$lastname, $firstname")
        }

        $form.Controls.Add($listBox) 

        $form.Topmost = $True

        $result = $form.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $select = $listBox.SelectedItem
        }
        elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
        {
            return
        }

        #Get the value the database will understand
        foreach($data in $SQLDataset.tables[0])
        {
            $datadisplay = $data[4]+", "+$data[5]
            if ($datadisplay -eq $select)
            {
                $username = $data[0]
            }
        }
    }

    if($Global:ProfileSearch -eq "simple")
    {
        $SQLCommand.CommandText = "select [dbo].ITEM.DESCRIPTION, MODEL_NAME, SERIAL_NUMBER, [dbo].USERS.FIRST_NAME, [dbo].USERS.LAST_NAME, [dbo].LOCATION.NAME from [dbo].ITEM left join [dbo].USERS on [dbo].ITEM.CUSTODIAN_NAME=[dbo].USERS.USER_NAME left join [dbo].INVOICE on [dbo].ITEM.INVOICE_ID=[dbo].INVOICE.ID left join [dbo].LOCATION on [dbo].ITEM.LOCATION_ID=[dbo].LOCATION.ID where [dbo].USERS.USER_NAME like '$username'"
    }
    else
    {
        $SQLCommand.CommandText = "select [dbo].ITEM.DESCRIPTION, MODEL_NAME, MODEL_NUMBER, SERIAL_NUMBER, [dbo].USERS.FIRST_NAME, [dbo].USERS.LAST_NAME, [dbo].ITEM.DATE_ENTERED, WARRANTY_EXPIRY_DATE, [dbo].INVOICE.INVOICENUM, PURCHASE_PRICE, [dbo].LOCATION.NAME from [dbo].ITEM left join [dbo].USERS on [dbo].ITEM.CUSTODIAN_NAME=[dbo].USERS.USER_NAME left join [dbo].INVOICE on [dbo].ITEM.INVOICE_ID=[dbo].INVOICE.ID left join [dbo].LOCATION on [dbo].ITEM.LOCATION_ID=[dbo].LOCATION.ID where [dbo].USERS.USER_NAME like '$username'"
    }
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null
    $tablevalue = @()
    foreach ($data in $SQLDataset.tables[0])
    {
        $data | Format-List
        $TSCNum = $data.DESCRIPTION
    }
    
    $table = $SQLDataset.Tables[0]
    #Check to see if user would like to update asset
    $update = Read-Host "Would you like to update this asset? (Y/N)"
    $update = $update.ToLower()
    if($update -eq 'y')
    {
        if($table.Rows.Count -eq 1)
        {
            Update-Asset($TSCNum)
        }
        else
        {
            $TSCNum = Read-Host "Please specify which asset number (TSC-####)"
            Update-Asset($TSCNum)
        }        
    }
    
}

function Update-User
{
    $Last = Read-Host "User's Last Name"

    #Create a SQL Command Object
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
    $SQLCommand.CommandText = "select * from [dbo].USERS WHERE LAST_NAME like '$Last'"
    $SQLCommand.Connection = $SQLConnection

    #Create Form with objects in it
    #Grabbing the User
    #Taken from https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/selecting-items-from-a-list-box?view=powershell-6
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Select a User"
    $form.Size = New-Object System.Drawing.Size(600,600) 
    $form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(150,400)
    $OKButton.Size = New-Object System.Drawing.Size(100,35)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(275,400)
    $CancelButton.Size = New-Object System.Drawing.Size(100,35)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(400,40) 
    $label.Text = "Please select a User:"
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox 
    $listBox.Location = New-Object System.Drawing.Point(40,70) 
    $listBox.Size = New-Object System.Drawing.Size(500,300) 
    $listBox.Height = 300

    #Create an updated list of Vendors to populate list box
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null
    $tablevalue = @()
    $table = $SQLDataset.Tables[0]

    if($table.Rows.Count -eq 0)
    {
        Write-Host "User does not exist"
        pause
        return
    }
    elseif($table.Rows.Count -eq 1)
    {
        foreach($data in $table)
        {
            $username = $data[0]
            $OGFirstname = $data[5]
            $OGLastname = $data[4]
        }
    }
    else
    {
        foreach ($data in $SQLDataset.tables[0])
        {
                $firstname = $data[5]
                $lastname = $data[4]
                [void] $listBox.Items.Add("$lastname, $firstname")
        }

        $form.Controls.Add($listBox) 

        $form.Topmost = $True

        $result = $form.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $select = $listBox.SelectedItem
        }
        #Cancel button was hit
        if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
        {
            return
        }

        #Get the value the database will understand
        foreach($data in $SQLDataset.tables[0])
        {
            $datadisplay = $data[4]+", "+$data[5]
            if ($datadisplay -eq $select)
            {
                $username = $data[0]
                $OGFirstname = $data[5]
                $OGLastname = $data[4]
            }
        }
    }

    #Create a form to edit the data
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Edit User"
    $form.Size = New-Object System.Drawing.Size(600,600) 
    $form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(150,400)
    $OKButton.Size = New-Object System.Drawing.Size(100,35)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(275,400)
    $CancelButton.Size = New-Object System.Drawing.Size(100,35)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(400,40) 
    $label.Text = "Please edit information:"
    $form.Controls.Add($label)

    $FirstLabel = New-Object System.Windows.Forms.Label
    $FirstLabel.Location = New-Object System.Drawing.Point(40,110) 
    $FirstLabel.Size = New-Object System.Drawing.Size(400,40) 
    $FirstLabel.Text = "First Name:"
    $form.Controls.Add($FirstLabel)

    $FtextBox = New-Object System.Windows.Forms.TextBox 
    $FtextBox.Location = New-Object System.Drawing.Point(100,150) 
    $FtextBox.Size = New-Object System.Drawing.Size(260,20)
    $FtextBox.Text = $OGFirstname 
    $form.Controls.Add($FtextBox)

    $LastLabel = New-Object System.Windows.Forms.Label
    $LastLabel.Location = New-Object System.Drawing.Point(40,190) 
    $LastLabel.Size = New-Object System.Drawing.Size(400,40) 
    $LastLabel.Text = "Last Name:"
    $form.Controls.Add($LastLabel)

    $LtextBox = New-Object System.Windows.Forms.TextBox 
    $LtextBox.Location = New-Object System.Drawing.Point(100,230) 
    $LtextBox.Size = New-Object System.Drawing.Size(260,20)
    $LtextBox.Text = $OGLastname 
    $form.Controls.Add($LtextBox)

    $form.Topmost = $True
    $result = $form.ShowDialog()


    $eventHandler = [System.EventHandler]{
    $FtextBox.Text;
    $LtextBox.Text;
    $form.Close();};
    $OKButton.Add_Click($eventHandler);

    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        return
    }

    $UpFirst = $FtextBox.Text
    $UpLast = $LtextBox.Text

    $SQLCommand.CommandText = "update [dbo].USERS set FIRST_NAME='$UpFirst', LAST_NAME='$UpLast' where USER_NAME = '$username'"
    $SQLCommand.Connection = $SQLConnection
    $rowsaffected = $SQLCommand.ExecuteNonQuery()

    #Debugging
    #if($OGFirstname -ne $UpFirst){
    #    Write-Host "First Name Updated"
    #}
    #if($OGLastname -ne $UpLast){
    #    Write-Host "Last Name Updated"
    #}
    #
    #pause
}

function Update-Asset($TSCNum)
{
    #Debugging
    #Write-Host $TSCNum
    #pause

    #Create a SQL Command Object
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
    $SQLCommand.CommandText = "select [dbo].ITEM.ID, [dbo].ITEM.DESCRIPTION, MODEL_NAME, SERIAL_NUMBER, [dbo].USERS.USER_NAME, WARRANTY_EXPIRY_DATE, [dbo].INVOICE.INVOICENUM, PURCHASE_PRICE, [dbo].ITEM.LOCATION_ID from [dbo].ITEM left join [dbo].USERS on [dbo].ITEM.CUSTODIAN_NAME=[dbo].USERS.USER_NAME left join [dbo].INVOICE on [dbo].ITEM.INVOICE_ID=[dbo].INVOICE.ID where [dbo].ITEM.DESCRIPTION like '%$TSCNum%'"
    $SQLCommand.Connection = $SQLConnection

    #Create an adapter to execute a SQL command
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null

    $table = $SQLDataset.tables[0]

    if($table.Rows.Count -eq 1)
    {

        #Grab all the data from the SQL Query
        $AssetID = $table.ID
        $OGDescription = $table.DESCRIPTION
        $OGModelName = $table.MODEL_NAME
        $OGSerial = $table.SERIAL_NUMBER
        $OGUsername = $table.USER_NAME
        $OGWarrDate = $table.WARRANTY_EXPIRY_DATE
        $OGInvoice = $table.INVOICENUM
        $OGPrice = $table.PURCHASE_PRICE
        $OGLocation = $table.LOCATION_ID

        #Grab all the Users
        $SQLCommand.CommandText = "select USER_NAME, FIRST_NAME, LAST_NAME from [dbo].USERS order by LAST_NAME asc"
        $SQLAdapter.SelectCommand = $SQLCommand
        $SQLDataset = New-Object System.Data.DataSet 
        $SQLAdapter.fill($SQLDataset) | out-null

        $Users = $SQLDataset.tables[0]

        #Grab all the Invoices
        $SQLCommand.CommandText = "select INVOICENUM, ID from [dbo].INVOICE"
        $SQLAdapter.SelectCommand = $SQLCommand
        $SQLDataset = New-Object System.Data.DataSet 
        $SQLAdapter.fill($SQLDataset) | out-null

        $Invoices = $SQLDataset.Tables[0]

        #Grab all locations
        $SQLCommand.CommandText = "select ID, NAME from [dbo].LOCATION"
        $SQLAdapter.SelectCommand = $SQLCommand
        $SQLDataset = New-Object System.Data.DataSet 
        $SQLAdapter.fill($SQLDataset) | out-null

        $Locations = $SQLDataset.Tables[0]

        #Create a form to edit the data
        $form = New-Object System.Windows.Forms.Form 
        $form.Text = "Edit Asset"
        $form.Size = New-Object System.Drawing.Size(900,600) 
        $form.StartPosition = "CenterScreen"

        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(330,475)
        $OKButton.Size = New-Object System.Drawing.Size(100,35)
        $OKButton.Text = "OK"
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $OKButton
        $form.Controls.Add($OKButton)

        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(470,475)
        $CancelButton.Size = New-Object System.Drawing.Size(100,35)
        $CancelButton.Text = "Cancel"
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $CancelButton
        $form.Controls.Add($CancelButton)

        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20) 
        $label.Size = New-Object System.Drawing.Size(200,40) 
        $label.Text = "Please edit information:"
        $form.Controls.Add($label)

        $TSCLabel = New-Object System.Windows.Forms.Label
        $TSCLabel.Location = New-Object System.Drawing.Point(40,110) 
        $TSCLabel.Size = New-Object System.Drawing.Size(200,40) 
        $TSCLabel.Text = "TSC Number:"
        $form.Controls.Add($TSCLabel)

        $TSCtextBox = New-Object System.Windows.Forms.TextBox 
        $TSCtextBox.Location = New-Object System.Drawing.Point(100,150) 
        $TSCtextBox.Size = New-Object System.Drawing.Size(260,20)
        $TSCtextBox.Text = $OGDescription 
        $form.Controls.Add($TSCtextBox)

        $ModelLabel = New-Object System.Windows.Forms.Label
        $ModelLabel.Location = New-Object System.Drawing.Point(40,190) 
        $ModelLabel.Size = New-Object System.Drawing.Size(200,40) 
        $ModelLabel.Text = "Model Name:"
        $form.Controls.Add($ModelLabel)

        $ModeltextBox = New-Object System.Windows.Forms.TextBox 
        $ModeltextBox.Location = New-Object System.Drawing.Point(100,230) 
        $ModeltextBox.Size = New-Object System.Drawing.Size(260,20)
        $ModeltextBox.Text = $OGModelName 
        $form.Controls.Add($ModeltextBox)

        $SerialLabel = New-Object System.Windows.Forms.Label
        $SerialLabel.Location = New-Object System.Drawing.Point(40,270) 
        $SerialLabel.Size = New-Object System.Drawing.Size(200,40) 
        $SerialLabel.Text = "Serial Number:"
        $form.Controls.Add($SerialLabel)

        $SerialtextBox = New-Object System.Windows.Forms.TextBox 
        $SerialtextBox.Location = New-Object System.Drawing.Point(100,310) 
        $SerialtextBox.Size = New-Object System.Drawing.Size(260,20)
        $SerialtextBox.Text = $OGSerial
        $form.Controls.Add($SerialtextBox)

        $UserLabel = New-Object System.Windows.Forms.Label
        $UserLabel.Location = New-Object System.Drawing.Point(40,350) 
        $UserLabel.Size = New-Object System.Drawing.Size(200,40) 
        $UserLabel.Text = "User:"
        $form.Controls.Add($UserLabel)

        $DropDown = new-object System.Windows.Forms.ComboBox
        $DropDown.Location = new-object System.Drawing.Size(100,390)
        $DropDown.Size = new-object System.Drawing.Size(260,20)

        $c = 0
        ForEach ($Item in $Users)
        {
            $display = $Item[2] + ", " + $Item[1]
            [void] $DropDown.Items.Add($display)

            if($OGUsername -eq $Item[0])
            {
                $UserNum = $c
            }

            $c = $c+1
        }

        $form.Controls.Add($DropDown)

        #Makes the current user the selected user in the drop box
        $DropDown.SelectedIndex = $UserNum

        $WarrantyLabel = New-Object System.Windows.Forms.Label
        $WarrantyLabel.Location = New-Object System.Drawing.Point(400,110) 
        $WarrantyLabel.Size = New-Object System.Drawing.Size(200,40) 
        $WarrantyLabel.Text = "Warranty Expire Date:"
        $form.Controls.Add($WarrantyLabel)

        $WarrantytextBox = New-Object System.Windows.Forms.TextBox 
        $WarrantytextBox.Location = New-Object System.Drawing.Point(460,150) 
        $WarrantytextBox.Size = New-Object System.Drawing.Size(260,20)
        $WarrantytextBox.Text = $OGWarrDate
        $form.Controls.Add($WarrantytextBox)

        $InvoiceLabel = New-Object System.Windows.Forms.Label
        $InvoiceLabel.Location = New-Object System.Drawing.Point(400,190) 
        $InvoiceLabel.Size = New-Object System.Drawing.Size(200,40) 
        $InvoiceLabel.Text = "Invoice Number:"
        $form.Controls.Add($InvoiceLabel)

        $InvoiceDropDown = new-object System.Windows.Forms.ComboBox
        $InvoiceDropDown.Location = new-object System.Drawing.Size(460,230)
        $InvoiceDropDown.Size = new-object System.Drawing.Size(260,20)

        $c = 0
        [void] $InvoiceDropDown.Items.Add("")
        ForEach ($Item in $Invoices)
        {
            $display = $Item[0]
            [void] $InvoiceDropDown.Items.Add($display)

            if($OGInvoice -eq $Item[0])
            {
                $Invoicenum = $c
            }

            $c = $c+1
        }

        $form.Controls.Add($InvoiceDropDown)

        #Makes the current invoice the selected user in the drop box
        $InvoiceDropDown.SelectedIndex = $Invoicenum

        $PurchaseLabel = New-Object System.Windows.Forms.Label
        $PurchaseLabel.Location = New-Object System.Drawing.Point(400,270) 
        $PurchaseLabel.Size = New-Object System.Drawing.Size(200,40) 
        $PurchaseLabel.Text = "Purchase Price:"
        $form.Controls.Add($PurchaseLabel)

        $PurchasetextBox = New-Object System.Windows.Forms.TextBox 
        $PurchasetextBox.Location = New-Object System.Drawing.Point(460,310) 
        $PurchasetextBox.Size = New-Object System.Drawing.Size(260,20)
        $PurchasetextBox.Text = $OGPrice
        $form.Controls.Add($PurchasetextBox)

        $LocationLabel = New-Object System.Windows.Forms.Label
        $LocationLabel.Location = New-Object System.Drawing.Point(400,350) 
        $LocationLabel.Size = New-Object System.Drawing.Size(200,40) 
        $LocationLabel.Text = "Location:"
        $form.Controls.Add($LocationLabel)

        $LocationDropDown = new-object System.Windows.Forms.ComboBox
        $LocationDropDown.Location = new-object System.Drawing.Size(460,390)
        $LocationDropDown.Size = new-object System.Drawing.Size(260,20)

        #This will set the Location for the dropdown
        $c = 0
        ForEach ($Item in $Locations)
        {
            $display = $Item[1]
            [void] $LocationDropDown.Items.Add($display)

            if($OGLocation -eq $Item[0])
            {
                $LocationNum = $c
            }

            $c = $c+1
        }

        $form.Controls.Add($LocationDropDown)
        $LocationDropDown.SelectedIndex = $LocationNum

        #Displaying and receiving form
        $form.Topmost = $True
        $result = $form.ShowDialog()


        $eventHandler = [System.EventHandler]{
        $TSCtextBox.Text;
        $ModeltextBox.Text;
        $SerialtextBox.Text;
        $DropDown.SelectedItem;
        $WarrantytextBox.Text;
        $PurchasetextBox.Text;
        $InvoiceDropDown.SelectedItem;
        $LocationDropDown.SelectedItem;
        $form.Close();};
        $OKButton.Add_Click($eventHandler);

        if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
        {
            return
        }

        $UpdateDescrip = ""

        if($TSCtextBox.Text -ne $OGDescription)
        {
            $UpdateDescrip = $UpdateDescrip + "Description: From:"+$OGDescription+" To:"+$TSCtextBox.Text+";"
            $OGDescription = $TSCtextBox.Text
        }
        if($ModeltextBox.Text -ne $OGModelName){
            $UpdateDescrip = $UpdateDescrip + "Model Name: From:"+$OGModelName+" To:"+$ModeltextBox.Text+";"
            $OGModelName = $ModeltextBox.Text
        }
        if($SerialtextBox.Text -ne $OGSerial)
        {
            $UpdateDescrip = $UpdateDescrip + "Serial Number: From:"+$OGSerial+" To:"+$SerialtextBox.Text+";"
            $OGSerial = $SerialtextBox.Text
        }
        foreach($Item in $Users){
            $User = $Item[2] + ", " + $Item[1]
            if($DropDown.SelectedItem -eq $User)
            {
                $UpdatedUser = $Item[0]
                if($UpdatedUser -ne $OGUsername)
                {
                    $UpdateDescrip = $UpdateDescrip + "Custodian: From:"+$OGUsername+" To:"+$UpdatedUser+";"
                    $OGUsername = $UpdatedUser
                }
            }
        }
        if($WarrantytextBox.Text -ne $OGWarrDate)
        {
            $UpdateDescrip = $UpdateDescrip + "Warranty Date: From:"+$OGWarrDate+" To:"+$WarrantytextBox.Text+";"
            $OGWarrDate = $WarrantytextBox.Text
        }
        if($InvoiceDropDown.SelectedItem -ne $OGInvoice)
        {
            $UpdateDescrip = $UpdateDescrip + "Invoice: From:"+$OGInvoice+" To:"+$InvoiceDropDown.SelectedItem+";"
            foreach($Item in $Invoices)
            {
                if($InvoiceDropDown.SelectedItem -eq $Item[0])
                {
                    $OGInvoice = $Item[1]
                }
            }
        }
        else{
            foreach($Item in $Invoices)
            {
                if($OGInvoice -eq $Item[0])
                {
                    $OGInvoice = $Item[1]
                }
            }
        }
        if($PurchasetextBox.Text -ne $OGPrice)
        {
            $UpdateDescrip = $UpdateDescrip + "Purchase Price: From:"+$OGPrice+" To:"+$PurchasetextBox.Text+";"
            $OGPrice = $PurchasetextBox.Text
        }
        if($LocationDropDown.SelectedItem -ne $OGLocation)
        {
            $UpdateDescrip = $UpdateDescrip + "Location: From:"+$OGLocation+" To:"+$LocationDropDown.SelectedItem+";"
            foreach($Item in $Locations)
            {
                if($LocationDropDown.SelectedItem -eq $Item[1])
                {
                    $OGLocation = $Item[0]
                }
            }
        }

        $LogObject = @{description = $UpdateDescrip;type = 3;ItemID = $AssetID}

        if($UpdateDescrip -ne "")
        {
            $dateNow = Logging($LogObject)
            $SQLCommand.CommandText = "update [dbo].ITEM set
                                            DESCRIPTION = '$OGDescription',
                                            BARCODE = '$OGDescription',
                                            TAGID = '$OGDescription',
                                            SERIAL_NUMBER = '$OGSerial',
                                            MODEL_NAME = '$OGModelName',
                                            CUSTODIAN_NAME = '$OGUsername',
                                            LOCATION_ID = '$OGLocation'
                                            "
            if($InvoiceDropDown.SelectedItem -ne "")
            {
                $SQLCommand.CommandText = $SQLCommand.CommandText + ",INVOICE_ID = '$OGInvoice'"
            }
            if($PurchasetextBox.Text -ne "")
            {
                $SQLCommand.CommandText = $SQLCommand.CommandText + ",PURCHASE_PRICE = '$OGPrice'"
            }
            if($WarrantytextBox.Text -ne "")
            {
                $SQLCommand.CommandText = $SQLCommand.CommandText + ",WARRANTY_EXPIRY_DATE = '$OGWarrDate'"
            }
            $SQLCommand.CommandText = $SQLCommand.CommandText + " where ID = $AssetID"

            #Creates SQL Adapter Object to execute the Insert Function
            $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
            $SQLCommand.ExecuteNonQuery()
            #pause
        }
    }
    ElseIf($table.Rows.Count -eq 0)
    {
        Write-Host "Asset does not exist"
        pause
    }
    else
    {
        #Naming the error Blueberry Sunrise was Jason's idea (It's a little easter egg)
        Write-Host "Multiple assets with same TSC Number. Please fix issue manually in database. (Error Code: Blueberry Sunrise)"
        pause
        return
    }
}
function Add-Asset-Single
{
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
    $SQLCommand.Connection = $SQLConnection

    #Create an adapter to execute a SQL command
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter

    #Grab all the Users
    $SQLCommand.CommandText = "select USER_NAME, FIRST_NAME, LAST_NAME from [dbo].USERS order by LAST_NAME asc"
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null

    $Users = $SQLDataset.tables[0]

    #Grab all the Invoices
    $SQLCommand.CommandText = "select INVOICENUM, ID from [dbo].INVOICE order by ID asc"
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null

    $Invoices = $SQLDataset.Tables[0]

    #Grab all the Companies
    $SQLCommand.CommandText = "select ID, COMPANY from [dbo].COMPANY"
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null

    $Companies = $SQLDataset.Tables[0]

    #Grab all the Locations
    $SQLCommand.CommandText = "select ID, NAME from [dbo].LOCATION"
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | out-null

    $Locations = $SQLDataset.Tables[0]

    #Create a form to edit the data
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Edit Asset"
    $form.Size = New-Object System.Drawing.Size(900,900) 
    $form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(330,775)
    $OKButton.Size = New-Object System.Drawing.Size(100,35)
    $OKButton.Text = "OK"
    #$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #$form.AcceptButton = $OKButton
    #$form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(470,775)
    $CancelButton.Size = New-Object System.Drawing.Size(100,35)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(200,40) 
    $label.Text = "Please edit information:"
    $form.Controls.Add($label)

    $TSCLabel = New-Object System.Windows.Forms.Label
    $TSCLabel.Location = New-Object System.Drawing.Point(40,110) 
    $TSCLabel.Size = New-Object System.Drawing.Size(200,40) 
    $TSCLabel.Text = "TSC Number:"
    $form.Controls.Add($TSCLabel)

    $TSCtextBox = New-Object System.Windows.Forms.TextBox 
    $TSCtextBox.Location = New-Object System.Drawing.Point(100,150) 
    $TSCtextBox.Size = New-Object System.Drawing.Size(260,20) 
    $form.Controls.Add($TSCtextBox)

    $ModelLabel = New-Object System.Windows.Forms.Label
    $ModelLabel.Location = New-Object System.Drawing.Point(40,190) 
    $ModelLabel.Size = New-Object System.Drawing.Size(200,40) 
    $ModelLabel.Text = "Model Name:"
    $form.Controls.Add($ModelLabel)

    $ModeltextBox = New-Object System.Windows.Forms.TextBox 
    $ModeltextBox.Location = New-Object System.Drawing.Point(100,230) 
    $ModeltextBox.Size = New-Object System.Drawing.Size(260,20) 
    $form.Controls.Add($ModeltextBox)

    $ModelNumLabel = New-Object System.Windows.Forms.Label
    $ModelNumLabel.Location = New-Object System.Drawing.Point(40,270) 
    $ModelNumLabel.Size = New-Object System.Drawing.Size(200,40) 
    $ModelNumLabel.Text = "Model Number:"
    $form.Controls.Add($ModelNumLabel)

    $ModelNumtextBox = New-Object System.Windows.Forms.TextBox 
    $ModelNumtextBox.Location = New-Object System.Drawing.Point(100,310) 
    $ModelNumtextBox.Size = New-Object System.Drawing.Size(260,20) 
    $form.Controls.Add($ModelNumtextBox)

    $SerialLabel = New-Object System.Windows.Forms.Label
    $SerialLabel.Location = New-Object System.Drawing.Point(40,350) 
    $SerialLabel.Size = New-Object System.Drawing.Size(200,40) 
    $SerialLabel.Text = "Serial Number:"
    $form.Controls.Add($SerialLabel)

    $SerialtextBox = New-Object System.Windows.Forms.TextBox 
    $SerialtextBox.Location = New-Object System.Drawing.Point(100,390) 
    $SerialtextBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($SerialtextBox)

    $SupplierLabel = New-Object System.Windows.Forms.Label
    $SupplierLabel.Location = New-Object System.Drawing.Point(40,430) 
    $SupplierLabel.Size = New-Object System.Drawing.Size(200,40) 
    $SupplierLabel.Text = "Supplier:"
    $form.Controls.Add($SupplierLabel)

    $SupplierDropDown = new-object System.Windows.Forms.ComboBox
    $SupplierDropDown.Location = new-object System.Drawing.Size(100,470)
    $SupplierDropDown.Size = new-object System.Drawing.Size(260,20)

    #This will set the Supplier default as Connection
    $c = 0
    ForEach ($Item in $Companies)
    {
        $display = $Item[1]
        [void] $SupplierDropDown.Items.Add($display)

        if("1038" -eq $Item[0])
        {
            $SupplierNum = $c
        }

        $c = $c+1
    }

    $form.Controls.Add($SupplierDropDown)
    $SupplierDropDown.SelectedIndex = $SupplierNum

    $ManufacturerLabel = New-Object System.Windows.Forms.Label
    $ManufacturerLabel.Location = New-Object System.Drawing.Point(40,510) 
    $ManufacturerLabel.Size = New-Object System.Drawing.Size(200,40) 
    $ManufacturerLabel.Text = "Manufacturer:"
    $form.Controls.Add($ManufacturerLabel)

    $ManufacturerDropDown = new-object System.Windows.Forms.ComboBox
    $ManufacturerDropDown.Location = new-object System.Drawing.Size(100,550)
    $ManufacturerDropDown.Size = new-object System.Drawing.Size(260,20)

    #This will set the Manufacturer default as Lenovo
    $c = 0
    ForEach ($Item in $Companies)
    {
        $display = $Item[1]
        [void] $ManufacturerDropDown.Items.Add($display)
        if("14" -eq $Item[0])
        {
            $ManufacturerNum = $c
        }

        $c = $c+1
    }

    $form.Controls.Add($ManufacturerDropDown)
    $ManufacturerDropDown.SelectedIndex = $ManufacturerNum

    $UserLabel = New-Object System.Windows.Forms.Label
    $UserLabel.Location = New-Object System.Drawing.Point(40,590) 
    $UserLabel.Size = New-Object System.Drawing.Size(200,40) 
    $UserLabel.Text = "User:"
    $form.Controls.Add($UserLabel)

    $DropDown = new-object System.Windows.Forms.ComboBox
    $DropDown.Location = new-object System.Drawing.Size(100,630)
    $DropDown.Size = new-object System.Drawing.Size(260,20)

    #This will set the User default as IT,Stockroom
    $c = 0
    ForEach ($Item in $Users)
    {
        $display = $Item[2] + ", " + $Item[1]
        [void] $DropDown.Items.Add($display)

        if("ITStockroom" -eq $Item[0])
        {
            $UserNum = $c
        }

        $c = $c+1
    }

    $form.Controls.Add($DropDown)
    #Makes the current user the selected user in the drop box
    $DropDown.SelectedIndex = $UserNum

    $LocationLabel = New-Object System.Windows.Forms.Label
    $LocationLabel.Location = New-Object System.Drawing.Point(400,110) 
    $LocationLabel.Size = New-Object System.Drawing.Size(200,40) 
    $LocationLabel.Text = "Location:"
    $form.Controls.Add($LocationLabel)

    $LocationDropDown = new-object System.Windows.Forms.ComboBox
    $LocationDropDown.Location = new-object System.Drawing.Size(460,150)
    $LocationDropDown.Size = new-object System.Drawing.Size(260,20)

    #This will set the Location default as 2369 IT Storage
    $c = 0
    ForEach ($Item in $Locations)
    {
        $display = $Item[1]
        [void] $LocationDropDown.Items.Add($display)

        if("23" -eq $Item[0])
        {
            $LocationNum = $c
        }

        $c = $c+1
    }

    $form.Controls.Add($LocationDropDown)
    $LocationDropDown.SelectedIndex = $LocationNum

    $WarrantyLabel = New-Object System.Windows.Forms.Label
    $WarrantyLabel.Location = New-Object System.Drawing.Point(400,190) 
    $WarrantyLabel.Size = New-Object System.Drawing.Size(200,40) 
    $WarrantyLabel.Text = "Warranty Expire Date:"
    $form.Controls.Add($WarrantyLabel)

    $WarrantytextBox = New-Object System.Windows.Forms.TextBox 
    $WarrantytextBox.Location = New-Object System.Drawing.Point(460,230) 
    $WarrantytextBox.Size = New-Object System.Drawing.Size(260,20)
    $WarrantytextBox.Text = ((Get-Date).AddYears(3)).ToString("MM/dd/yyyy")
    $form.Controls.Add($WarrantytextBox)

    $InvoiceLabel = New-Object System.Windows.Forms.Label
    $InvoiceLabel.Location = New-Object System.Drawing.Point(400,270) 
    $InvoiceLabel.Size = New-Object System.Drawing.Size(200,40) 
    $InvoiceLabel.Text = "Invoice Number:"
    $form.Controls.Add($InvoiceLabel)

    $InvoiceDropDown = new-object System.Windows.Forms.ComboBox
    $InvoiceDropDown.Location = new-object System.Drawing.Size(460,310)
    $InvoiceDropDown.Size = new-object System.Drawing.Size(260,20)

    #Initializes the items for Invoices
    $c=0
    ForEach ($Item in $Invoices)
    {
        $display = $Item[0]
        [void] $InvoiceDropDown.Items.Add($display)
        $c = $c + 1
    }

    $form.Controls.Add($InvoiceDropDown)
    #Makes the last input invoice in database as the default invoice
    $InvoiceDropDown.SelectedIndex = $c - 1

    $PurchaseLabel = New-Object System.Windows.Forms.Label
    $PurchaseLabel.Location = New-Object System.Drawing.Point(400,350) 
    $PurchaseLabel.Size = New-Object System.Drawing.Size(200,40) 
    $PurchaseLabel.Text = "Purchase Price:"
    $form.Controls.Add($PurchaseLabel)

    $PurchasetextBox = New-Object System.Windows.Forms.TextBox 
    $PurchasetextBox.Location = New-Object System.Drawing.Point(460,390) 
    $PurchasetextBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($PurchasetextBox)

    $PurchaseDateLabel = New-Object System.Windows.Forms.Label
    $PurchaseDateLabel.Location = New-Object System.Drawing.Point(400,430) 
    $PurchaseDateLabel.Size = New-Object System.Drawing.Size(200,40) 
    $PurchaseDateLabel.Text = "Purchase Date:"
    $form.Controls.Add($PurchaseDateLabel)

    $PurchaseDatetextBox = New-Object System.Windows.Forms.TextBox 
    $PurchaseDatetextBox.Location = New-Object System.Drawing.Point(460,470) 
    $PurchaseDatetextBox.Size = New-Object System.Drawing.Size(260,20)
    $PurchaseDatetextBox.Text = (Get-Date).ToString("MM/dd/yyyy")
    $form.Controls.Add($PurchaseDatetextBox)
<#
    $eventHandler = [System.EventHandler]{
    $TSCtextBox.Text;
    $ModeltextBox.Text;
    $ModelNumtextBox.Text;
    $SerialtextBox.Text;
    $SupplierDropDown.SelectedItem;
    $ManufacturerDropDown.SelectedItem;
    $DropDown.SelectedItem;
    $LocationDropDown.SelectedItem;
    $WarrantytextBox.Text;
    $PurchasetextBox.Text;
    $InvoiceDropDown.SelectedItem;
    $PurchaseDatetextBox.Text;
    $form.Close();};

    $OKButton.Add_Click($eventHandler);
#>
    # Add Validation Control
    $ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

    
    $eventHandler = [System.EventHandler]{
        $Global:OKFlag = $true
        if($TSCtextBox.Text.Length -eq 0)
        {
            $ErrorProvider.SetError($TSCtextBox, "Please input TSC Number")
            $Global:OKFlag = $false
        }
        else
        {
            $ErrorProvider.SetError($TSCtextBox, "")
        }
        if($ModeltextBox.Text.Length -eq 0)
        {
            $ErrorProvider.SetError($ModeltextBox, "Please input Model")
            $Global:OKFlag = $false
        }
        else
        {
            $ErrorProvider.SetError($ModeltextBox, "")
        }
        if($ModelNumtextBox.Text.Length -eq 0)
        {
            $ErrorProvider.SetError($ModelNumtextBox, "Please input Model Number")
            $Global:OKFlag = $false
        }
        else
        {
            $ErrorProvider.SetError($ModelNumtextBox, "")
        }
        if($SerialtextBox.Text.Length -eq 0)
        {
            $ErrorProvider.SetError($SerialtextBox, "Please input Serial Number")
            $Global:OKFlag = $false
        }
        else
        {
            $ErrorProvider.SetError($SerialtextBox, "")
        }
        if($OKFlag)
        {
            $TSCtextBox.Text;
            $ModeltextBox.Text;
            $ModelNumtextBox.Text;
            $SerialtextBox.Text;
            $SupplierDropDown.SelectedItem;
            $ManufacturerDropDown.SelectedItem;
            $DropDown.SelectedItem;
            $LocationDropDown.SelectedItem;
            $WarrantytextBox.Text;
            $PurchasetextBox.Text;
            $InvoiceDropDown.SelectedItem;
            $PurchaseDatetextBox.Text;
            $Global:OKFlag = $false
            $form.Close()
        }

    }
    $OKButton.Add_Click($eventHandler);
    $form.Controls.Add($OKButton)

    #Outputs the form to the display
    $form.Topmost = $True
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel -and $Global:OKFlag)
    {
        return
    }

    #Setting all of the values that will be put into the SQL Command
    #Foreach blocks are to make selected inputs readable by the database
    $TSCNum = $TSCtextBox.Text
    $Serial = $SerialtextBox.Text
    $ModelName = $ModeltextBox.Text
    $ModelNum = $ModelNumtextBox.Text
    foreach($Item in $Companies)
    {
        if($Item[1] -eq $SupplierDropDown.SelectedItem)
        {
            $Supplier = $Item[0]
        }
        if($Item[1] -eq $ManufacturerDropDown.SelectedItem)
        {
            $Manufacturer = $Item[0]
        }
    }
    foreach($Item in $Locations)
    {
        if($Item[1] -eq $LocationDropDown.SelectedItem)
        {
            $Location = $Item[0]
        }
    }
    foreach($Item in $Users){
        $User = $Item[2] + ", " + $Item[1]
        if($DropDown.SelectedItem -eq $User)
        {
             $Username = $Item[0]
        }
    }
    foreach($Item in $Invoices)
    {
        if($Item[0] -eq $InvoiceDropDown.SelectedItem)
        {
            $Invoice = $Item[1]
        }
    }
    $Warranty = [datetime]$WarrantytextBox.Text
    $Warranty = $Warranty.ToString("MM/dd/yyyy")
    $date = (Get-Date).ToString("MM/dd/yyyy")
    $ServiceExpire = ((Get-Date).AddYears(3)).ToString("MM/dd/yyyy")
    $PurchasePrice = $PurchaseTextBox.Text
    $PurchaseDate = $PurchaseDatetextBox.Text


    #Adding the object to Wisetrack
    $SQLCommand.CommandText = "select DESCRIPTION from [dbo].ITEM where DESCRIPTION = '$TSCNum'"
    $SQLAdapter.SelectCommand = $SQLCommand
    $SQLDataset = New-Object System.Data.DataSet 
    $SQLAdapter.fill($SQLDataset) | Out-Null

    $TSCTable = $SQLDataset.Tables[0]

    if($TSCTable.Rows.Count -eq 0)
    {
        #Debugging
        #Write-Host "TSCNumber: $TSCNum"
        #pause

        $SQLCommand.CommandText = "insert into [dbo].ITEM 
                                    (BARCODE, DESCRIPTION, SERIAL_NUMBER, MODEL_NUMBER, SUPPLIER_ID, MANUFACTURER_ID, LOCATION_ID, CUSTODIAN_NAME, INVOICE_ID, IN_SERVICE_DATE, DATE_ENTERED, ESTIMATED_LIFE, WARRANTY_EXPIRY_DATE, DATE_LAST_INVENTORIED,
                                    SERVICE_EXPIRY_DATE, SURPLUS, LOST, PURCHASE_PRICE, PURCHASE_DATE, SOLD, LEASED, BUSINESS_USE_PERCENT, DELETED, DYNAMIC, TAGID, INTERNAL_DEPRECIATED, MODEL_NAME)
                                    values 
                                    ('$TSCNum', '$TSCNum', '$Serial', '$ModelNum', '$Supplier', '$Manufacturer', '$Location', '$Username', '$Invoice', '$date', '$date', '36', '$Warranty', '$date', '$ServiceExpire', '0', '0', '$PurchasePrice', '$PurchaseDate', '0', '0', '100', '0', '0', '$TSCNum', '0', '$ModelName')"
        try
        {
            $SQLCommand.ExecuteNonQuery()
        }
        catch
        {
            Write-Host "Could not add Asset to database. Please re-input data"
            pause
            Add-Asset-Single
        }
        $SQLCommand.CommandText = "select ID from [dbo].ITEM where DESCRIPTION='$TSCNum' ORDER BY ID asc"
        $SQLAdapter.SelectCommand = $SQLCommand
        $SQLDataset = New-Object System.Data.DataSet 
        $SQLAdapter.fill($SQLDataset) | out-null

        #Grab the Asset ID
        foreach($Item in $SQLDataset.Tables[0])
        {
            $AssetID = $Item[0]
        }
        #Logging the add
        $UpdateDescrip = "BarCode:$TSCNum;TagID:$TSCNum;Description:$TSCNum;Custodian:$Username;Serial Number:$Serial;Model Number:$ModelNum;Model Name:$ModelName;Manufacturer Name:"+$ManufacturerDropDown.SelectedItem+";Invoice:"+$InvoiceDropDown.SelectedItem+";Date Last Inventoried:$date;Purchase Price:$PurchasePrice;Estimated Life:36;Warranty Expiry Date:$Warranty"
        $LogObject = @{description = $UpdateDescrip;type = 1;ItemID = $AssetID}

        $LogDate= Logging($LogObject)
        $PurchasePrice = [int]$PurchasePrice

        if($PurchasePrice -ge 1000)
        {
            #Write to Sharepoint Excel File
            #Code from: https://stackoverflow.com/questions/35606762/append-powershell-output-to-an-excel-file
            #Launch Excel
            $XL = New-Object -ComObject Excel.Application
            #Open the workbook
            $WB = $XL.Workbooks.Open("\\techsmith.com\departments\it\scripts\wisetrack\staging\UpdateSharepointAsset.xlsx")
            #Activate Sheet1, pipe to Out-Null to avoid 'True' output to screen
            $WB.Sheets.Item("Sheet1").Activate() | Out-Null
            #Find first blank row #, and activate the first cell in that row
            $FirstBlankRow = $($xl.ActiveSheet.UsedRange.Rows)[-1].Row + 1
            $XL.ActiveSheet.Range("A$FirstBlankRow").Activate()
            #Create PSObject with the properties that we want, convert it to a tab delimited CSV, and copy it to the clipboard
            $Record = [PSCustomObject]@{
                        'TSCNum' = $TSCNum
                        'Serial' = $Serial
                        'ModelName' = $ModelName
                        'Date' = $date
                        'Price' = $PurchasePrice
                        'Supplier' = $SupplierDropDown.SelectedItem
                        'Invoice' = $InvoiceDropDown.SelectedItem
            }
            $Record | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Select -Skip 1 | Clip
            #Paste at the currently active cell
            $XL.ActiveSheet.Paste() | Out-Null
            # Save and close
            $WB.Save() | Out-Null
            $WB.Close() | Out-Null
            $XL.Quit() | Out-Null
            #Release ComObject
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($XL)
        }
    }
    else
    {
        $Rows = $TSCTable.Rows.Count
        Write-Host "TSC Number already exists. Please try again; Number of Rows found: $Rows"
        Add-Asset-Single
    }
    
}

#Expects a powershell object with 3 values
#-Description
#-Transaction Type
#-ItemID
function Logging($LogObject)
{
    $dateNow = (Get-Date).ToString('yyyy-MM-dd hh:mm:ss')
    $description = $LogObject.description
    $type = $LogObject.type
    $itemID = $LogObject.ItemID

    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand

    #Create the SQL Statement and Connection
    #NOTE: $SQLConnection is the Connection to the database established at the top of Main Program in first Try statement
    $SQLCommand.CommandText = "insert into [dbo].LOGTRANSACTION (DATESTAMP, TYPE_ID, ITEM_ID, USER_NAME, DESCRIPTION) values ('$dateNow', $type, '$itemID','$LOGGEDINUSER', '$description');"
    $SQLCommand.Connection = $SQLConnection

    #Creates SQL Adapter Object to execute the Insert Function
    $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SQLCommand.ExecuteNonQuery()
    
    return $dateNow
}

#--------------------------█▄─▄█─▄▀▄─█─█▄─█───█▀▄─█▀▄─▄▀▀▄─▄▀▀──█▀▄─▄▀▄─█▄─▄█---------------------------------
#--------------------------█─▀─█─█▀█─█─█─▀█───█▀──██▀─█──█─█─▀█─██▀─█▀█─█─▀─█---------------------------------
#--------------------------▀───▀─▀─▀─▀─▀──▀───▀───▀─▀──▀▀───▀▀──▀─▀─▀─▀─▀───▀---------------------------------

#Debugging
#pause

function main
{
    Check-AD
    #$SQLConnection = New-Object System.Data.SqlClient.SqlConnection
    #$SQLConnection.ConnectionString = "server=$SCCMSQLSERVER;database=$DBNAME;user id=TSCCORP\$LOGGEDINUSER;password=$SECUREPASS;"
    #Write-Host $SQLConnection.ConnectionString
    #$SQLConnection.Open()
    #pause
    #Try to connect to the Database defined in global variable
    Try
    {
        $SQLConnection = New-Object System.Data.SqlClient.SqlConnection
        $SQLConnection.ConnectionString = "server=$SCCMSQLSERVER;database=$DBNAME;user id=TSCCORP\$LOGGEDINUSER;password=$SECUREPASS;trusted_connection=true;"
        #Write-Host $SQLConnection.ConnectionString
        #pause
        $SQLConnection.Open()
    }
    #Catch if cannot connect to database
    catch
    {
        [System.Windows.MessageBox]::Show('Could not connect to SQL Database', 'Error', 'OK', 'Error')
        exit
    }

    #Output main menu
    do
    {
        Clear-Host
        Write-Host "Logged in as $LOGGEDINUSER"
        Main-Menu
        $input = Read-Host
        $input = $input.ToLower()
        switch($input)
        {
            #Initialize the Add Menu
            'a'{
                do{
                    Clear-Host
                    Add-Menu
                    $AddInput = Read-Host
                    $AddInput = $AddInput.ToLower()
                    switch($AddInput)
                    {
                        'a'{
                            Clear-Host
                            $AddAsset = Read-Host "Single or Multiple?(S/M)"
                            $AddAsset = $AddAsset.ToLower()
                            switch($AddAsset)
                            {
                                's'{
                                    #Stable
                                    Add-Asset-Single
                                }'m'{
                                    #TODO
                                    Write-Host "Make sure file is updated and closed!"
                                    pause
                                    Add-Asset-Batch
                                }
                            }
                        }'i'{
                            Clear-Host
                            #Stable
                            Add-Invoice
                        }'u'{
                            Clear-Host
                            #Stable
                            Add-User
                        }
                    }
                }
                until($AddInput -eq 'c')
            #COMPLETE
            }'u'{
                do{
                    Clear-Host
                    Update-Menu
                    $UpdateInput = Read-Host
                    $UpdateInput = $UpdateInput.ToLower()
                    switch($UpdateInput)
                    {
                        'a'{
                            Clear-Host
                            #Stable
                            $TSCNum = Read-Host "Enter TSC Number (TSC-####)"
                            Update-Asset($TSCNum)
                        }'u'{
                            Clear-Host
                            #Stable
                            Update-User
                        }
                    }
                }
                until($UpdateInput -eq 'c')
            #COMPLETE
            }'s'{
                do{
                    Clear-Host
                    Search-Menu
                    $SearchInput = Read-Host
                    $SearchInput = $SearchInput.ToLower()
                    switch($SearchInput)
                    {
                        'a'{
                            Clear-Host
                            #Stable
                            Search-Asset-TSC
                        }'u'{
                            Clear-Host
                            #Stable
                            Search-Asset-User
                        }'s'{
                            Clear-Host
                            #
                            Search-Asset-Serial
                        }
                    }
                }
                until($SearchInput -eq 'c')
            }'l'{
                Clear-Host
                #Returns logs of a specific TSC Asset
                Return-Logs
            }'p'{
                Clear-Host
                #Checks to see if profile has been made
                $exists = Profile-Check
                if($exists)
                {
                    Update-Profile
                }
            }'h'{
                Clear-Host
                #Displays the about and functionality of the program
                Display-About

            }'ee'{
                Clear-Host
                #Easter Egg!
                Easter-Egg
            }'q'{
                return
            }
        }
        #pause
    }
    until($input -eq 'q')
    $SQLConnection.close()
}
main

#Verifies the user has access to the database
#if($env:COMPUTERNAME -eq "TSC-1084")
#{
#    $SQLConnection = new-object System.Data.Odbc.OdbcConnection
#    $SQLConnection.ConnectionString = "DSN=SQL05`$OTS_WiseTrack"
#    #Write-Host $SQLConnection.ConnectionString
#    #pause
#    $SQLConnection.Open()
#    main
#}
#else
#{
#    main
#}