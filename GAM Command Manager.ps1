<#
GAM Command Manager | STABLE
#>

<# Parameter #>
param([switch] $Help)

<# Global Varibles #>
    $global:Output =""
    $ScriptVersion = "STABLE 1.4"

<# Pre Initilization #>
cls

<# CHANGE OPTIONS HERE #>

    # General GAM Settings
        $GAMLocation = "" #Full path of server. EG. \\MainServer\c$\GAM
        $GamServerName = "" # NAME OF GAM SERVER. EG. Google-Server

    # Banned users settings
        # If you have a banned group to stop sending email or accessing a service, add the options here
        $ADGroup = "" # The name of the group on AD EG. 'BannedUsers'
        $GoogleGroup = "" # The name of the Google Group. EG 'BannedUsers@SomeSchool.com'

<# END OF OPTIONS #>

<# Help Section #>

    if ($Help)
    {
    write-output "********* GAM Command Manager | Version $ScriptVersion | By Iri5s @ Iri5s.com *********" 
    write-output "***************************************************************************************"
    write-output "**** This script will is a GUI for GAM (Google Apps Manager) to help manage Google ****"
    write-output "***************************************************************************************"
    write-output "************************************ Help Section *************************************"
    write-output "***************************************************************************************"

    write-output "- Global varibles are set in the script. Please make sure these are changed. "
    write-output "- Some features are marked [WIP], so are not implemented, please bare this in mind."
    exit
    }

<# Menu Functions #>

    function Show-Menu {
        param (
            [string]$Title = "$Title"
        )
        Clear-Host

        if ($Warning)
            {Write-Host "WARNING: You are currently running this script remotely, results will be printed in console only." -ForegroundColor Yellow }

        Write-Host "================ $Title ================" -ForegroundColor Cyan
        
        if ($Option1) {Write-Host "1: Press '1' for $Option1" -ForegroundColor Yellow}
        if ($Option2) {Write-Host "2: Press '2' for $Option2" -ForegroundColor Yellow}
        if ($Option3) {Write-Host "3: Press '3' for $Option3" -ForegroundColor Yellow}
        if ($Option4) {Write-Host "4: Press '4' for $Option4" -ForegroundColor Yellow}
        if ($Option5) {Write-Host "5: Press '5' for $Option5" -ForegroundColor Yellow}
        if ($Option6) {Write-Host "6: Press '6' for $Option6" -ForegroundColor Yellow}
        if ($Option7) {Write-Host "7: Press '7' for $Option7" -ForegroundColor Yellow}
        if ($Option8) {Write-Host "8: Press '8' for $Option8" -ForegroundColor Yellow}
        Write-Host "M: Press 'M' to go to menu" -ForegroundColor Cyan
        Write-Host "Q: Press 'Q' to quit." -ForegroundColor Cyan
        }

    function Main-Menu {
        param (
            [string]$Title="Main Menu"
        )
        PreMenu
        Clear-Variable Option*
        $Title = "Main Menu"
        $Option1 = "Google Classroom Management"
        $Option2 = "Gmail Management"
        Show-Menu
        do
         {
            $selection = Read-Host "Please select an option"
            switch ($selection)
            {
            '1' {ClassroomMainMenu} 
            '2' {GmailManagementMenu}
            'Q' {QuitScript}
            }
            
         }
         until ($selection -eq'Q')
        }


    function ClassroomMainMenu {
        param (
            [string]$Title="Google Classroom Management"
        )
        PreMenu
        Clear-Variable Option*
        $Title = "Google Classroom Management"
        $Option1 = "Class Creation and Deletion"
        $Option2 = "Adding/Removing Teachers"
        $Option3 = "Adding/Removing Students"
        $Option4 = "General Info"
        $Option5 = "Course Admin"
        $Option6 = "Misc"
        Show-Menu
        do
         {
            $selection = Read-Host "Please select an option"
            switch ($selection)
            {
            '1' {CreateAndDeleteMenu} 
            '2' {AddingAndRemovingTeachersMenu}
            '3' {AddingAndRemovingStudentsMenu}
            '4' {GeneralInfoMenu}
            '5' {CourseAdminMenu}
            '6' {MiscMenu}
            'M' {Main-Menu}
            'Q' {QuitScript}
            }
            
         }
         until ($selection -eq'Q')
        }

        function GmailManagementMenu {
        param (
            [string]$Title="Google Mail Management"
        )
        PreMenu
        Clear-Variable Option*
        $Title = "Gmail Management Menu"
        $Option1 = "Delegate Menu"
        $Option2 = "Manage Email Bans"
        Show-Menu
        do
         {
            $selection = Read-Host "Please select an option"
            switch ($selection)
            {
            '1' {GmailDelegateMenu} 
            '2' {BannedUsersMenu}
            'M' {Main-Menu}
            'Q' {QuitScript}
            }
         }
         until ($selection -eq'Q')
        }

#// GOOGLE CLASSROOM FUNCTIONS // #

    function CreateAndDeleteMenu {
        param (
            [string]$Title="Create/Delete Menu"
        )
        PreMenu
        Clear-Variable Option*
        $Title = "Create/Delete Menu"
        $Option1 = "Create a class"
        $Option2 = "Delete a class"
        $Option3 = "Main Menu"
        Show-Menu
        do
         {
            $selection = Read-Host "Please select an option"
            switch ($selection)
            {
            '1' {# Create Class
                    $Alias = Read-Host "Please type in name of class"
                    $Teacher = Read-Host "Please type in teacher's email"
                    ./Gam create course alias $Alias teacher $Teacher name $Alias
                    ./gam update course $Alias status Active
                    Write-Host "Course $Alias created" -ForegroundColor Green
                    Read-Host "Press 'enter' to continue.."
                    ClassroomMainMenu} 
            '2' {# Deletes a class
                    $Alias = Read-Host "Please type in ID of class"
                    CourseLookup
                    Write-Host "WARNINGS: You are about to delete class " -ForegroundColor Red -NoNewLine
                    Write-Host "$CourseName! " -ForegroundColor Yellow -NoNewLine
                    Write-Host "are you sure?" -ForegroundColor Red
                    $Confirm = Read-Host "Y/N"

                    if ($Confirm -eq "Y")
                         {
                        ./gam delete course $Alias
                        Write-Host "Course $Alias deleted" -ForegroundColor Green
                        Read-Host "Press 'enter' to continue.."}
                    else { Write-Host "Command Aborted" -ForegroundColor Yellow 
                    Read-Host "Press 'enter' to continue.."}
                    ClassroomMainMenu}
            '3' {ClassroomMainMenu}
            'M' {ClassroomMainMenu}
            'Q' {QuitScript}
            }

         }
         until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

    function AddingAndRemovingTeachersMenu {
        param (
            [string]$Title="Adding And Removing Teachers"
        )
        PreMenu
        Clear-Variable Option*
        $Title = "Adding And Removing Teachers"
        $Option1 = "Add a Teacher to a class"
        $Option2 = "Remove a Teacher from a class"
        Show-Menu
        do
         {
            $selection = Read-Host "Please select an option"
            switch ($selection)
            {
            '1' {
                Write-Host "Type 'csv' to run from csv file"
                $Alias = Read-Host "Type in Course ID/Name"

                if (($Alias -eq "CSV"))
                    {
                        $CSVPath = "$GAMLocation\_Config\AddITeacher.txt"
                        $Header1 = "`"~Class`""
                        $Header2 = "`"~Email`""
                        .\gam csv $CSVPath gam course $Header1 add teacher $Header2
                        Read-Host "Press 'enter' to continue.."
                        ClassroomMainMenu
                    }
                else{
                    $Teacher = Read-Host "Please type in teacher's email"
                    .\gam course $Alias add teacher $Teacher
                    Read-Host "Press 'enter' to continue.."
                    ClassroomMainMenu
                    }
                } 
            '2' {
                Write-Host "Type 'csv' to run from csv file"
                $Alias = Read-Host "Type in Course ID/Name"

                if (($Alias -eq "CSV"))
                    {
                        $CSVPath = "$GAMLocation\_Config\RemoveITeacher.txt"
                        $Header1 = "`"~Class`""
                        $Header2 = "`"~Email`""
                        .\gam csv $CSVPath gam course $Header1 remove teacher $Header2
                        Read-Host "Press 'enter' to continue.."
                        ClassroomMainMenu
                    }
                else{
                    $Teacher = Read-Host "Please type in teacher's email"
                    .\gam course $Alias remove teacher $Teacher
                    Read-Host "Press 'enter' to continue.."
                    ClassroomMainMenu
                    }
                }
            'M' {ClassroomMainMenu}
            'Q' {QuitScript}
            }
            
         }
         until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

    function AddingAndRemovingStudentsMenu {
        param (
            [string]$Title="Add and Remove Students menu"
        )
        PreMenu
        Clear-Variable Option*
        $Title = "Add and Remove Students menu"
        $Option1 = "Add a Student"
        $Option2 = "Remove a Student"
        Show-Menu
        do
         {
            $selection = Read-Host "Please select an option"
            switch ($selection)
            {
            '1' {
                Write-Host "Type 'CSV' to run from csv file"
                $Alias = Read-Host "Type in Course ID/Name"

                if (($Alias -eq "CSV"))
                    {
                        $CSVPath = "$GAMLocation\_Config\AddIStudent.txt"
                        $Header1 = "`"~Class`""
                        $Header2 = "`"~Email`""
                        .\gam csv $CSVPath gam course $Header1 add student $Header2
                        Read-Host "Press 'enter' to continue.."
                        ClassroomMainMenu
                    }
                else{
                    $Student = Read-Host "Please type in email of student"
                    .\gam course $Alias add student $Student
                    Read-Host "Press 'enter' to continue.."
                    ClassroomMainMenu
                    }
                } 
            '2' {
                Write-Host "Type 'csv' to run from csv file"
                $Alias = Read-Host "Type in Course ID/Name"

                if (($Alias -eq "CSV"))
                    {
                        $CSVPath = "$GAMLocation\_Config\RemoveIStudent.txt"
                        $Header1 = "`"~Class`""
                        $Header2 = "`"~Email`""
                        .\gam csv $CSVPath gam course $Header1 add student $Header2
                        Read-Host "Press 'enter' to continue.."
                        ClassroomMainMenu
                    }
                else{
                    $Student = Read-Host "Please type in email of student"
                    .\gam course $Alias remove student $Student
                    Read-Host "Press 'enter' to continue.."
                    ClassroomMainMenu
                    }
                } 
            #Could have an option to remove all students from a class..
            'M' {ClassroomMainMenu}
            'Q' {QuitScript}
            }
            
         }
         until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

    function GeneralInfoMenu 
        {
        param (
            [string]$Title="General Info Menu"
        )
        PreMenu
        Clear-Variable Option*
        $Title = "General Info Menu"
        $Option1 = "Get Course Info"
        $Option2 = "List Teacher Courses"
        Show-Menu
        do
         {
            $selection = Read-Host "Please select an option"
                switch ($selection)
                {
                '1' {   # Get Course Info
                        $SingleorMulti = Read-Host "Run command from a CSV file?"
                        $OutPath = "$GAMLocation\_Results\CourseInfo.txt"

                        if (($SingleorMulti -eq "Y") -or ($SingleorMulti -eq "Yes"))
                            {
                                $CSVPath = Read-Host "Path of CSV"
                                $Header = "`"~ID`""
                                .\gam csv $CSVPath gam info course $Header > $OutPath 
                            }
                        else{
                            $Alias = Read-Host "Type in Course ID"
                            .\gam info course $Alias > $OutPath
                            }
                        Write-Host "Done! Do you want to open '$OutPath'?" -ForegroundColor Yellow
                        Write-Host "Selecting No will display results in console!" -ForegroundColor Yellow
                        $Answer = Read-Host "Yes/No"
                    
                            if (($Answer -eq "Y") -or ($Answer -eq "Yes"))
                            {OpenOutFile}
                            else {DisplayResults}

                            Read-Host "Press 'enter' to continue.."
                            ClassroomMainMenu
                    }
                    
                '2' {   # List Teacher Courses
                        $Teacher = Read-Host "Please type in teacher's email"
                        $global:OutPath = "$GAMLocation\_Results\ITeacherCourses.txt"
                        ./Gam print courses teacher $Teacher > $OutPath
                        Write-Host "Course $Alias created" -ForegroundColor Yellow

                        Write-Host "Done! Do you want to open '$OutPath'?" -ForegroundColor Yellow
                        Write-Host "Selecting No will display results in console!" -ForegroundColor Yellow
                        $Answer = Read-Host "Yes/No"
                    
                            if (($Answer -eq "Y") -or ($Answer -eq "Yes"))
                            {OpenOutFile}
                            else {DisplayResults}
                            
                            Read-Host "Press 'enter' to continue.."
                            ClassroomMainMenu
                    } 
                'M' {ClassroomMainMenu}
                'Q' {QuitScript}
            }
            }        
        until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }
    
    function CourseAdminMenu 
        {
        param (
            [string]$Title="Course Admin Menu"
        )
        PreMenu
        Clear-Variable Option*
        $Title = "Course Admin Menu"
        $Option1 = "Set Class Status"
        $Option2 = "Change Class Owner"
        Show-Menu
        do
         {
            $selection = Read-Host "Please select an option"
            switch ($selection)
            {
            '1' {   #Set Class Status
                    $Alias = Read-Host "Type in Course ID"
                    CourseLookup
                    Write-Host "Current course $Alias state is" -NoNewLine
                    Write-Host " $CourseStatus" -ForegroundColor Yellow
                    $Status = Read-Host "Please type wanted Status 'Active/Archived'"
                    .\gam update course $Alias status $Status
                    Start-Sleep -s 5
                    ClassroomMainMenu } 
            '2' {   #Change class owner
                    $Alias = Read-Host "Type in Course ID"
                    $Teacher = Read-Host "Please type in teachers email"
                    .\gam update course $Alias owner $Teacher
                    CourseLookup
                    Write-Host "$Teacher is now owner of '$CourseName'" -ForegroundColor Green
                    Start-Sleep -s 5
                    ClassroomMainMenu} 
            'M' {ClassroomMainMenu}
            'Q' {QuitScript}
            }
            
         }
         until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

        function MiscMenu 
        {
        param (
            [string]$Title="Misc Menu"
        )
        PreMenu
        Clear-Variable Option*
        $Title = "Misc Menu"
        $Option1 = "Archive ALL Classes"
        $Option2 = "Run Class Sync [WIP]"
        Show-Menu
        do
         {
            $selection = Read-Host "Please select an option"
            switch ($selection)
            { 
            '1' {   #Archive All Classes
                    $ResultsFile = Read-Host "$GAMLocation\_Results\AllCoursesRAW.csv"
                    .\gam print courses state active > "$GAMLocation\_Results\AllCoursesRAW.csv"
                    $CSVConvert = ConvertTo-Csv -inputobject $ResultsFile -NoTypeInformation
                    $AllCourseName = Import-Csv $ResultsFile | select name | Select-Object -ExpandProperty name # This returns a list of names of all the courses. Now to filter..
                    $FilteredCourseName = $AllCourseName
                    $ImportCSV = Import-Csv $ResultsFile | export-csv "$GAMLocation\_Results\AllCourses.csv" -NoTypeInformation   
                    Write-Host "Are you sure you want to archive all classes in $ResultsFile?" -ForegroundColor Red
                    $Confirm = Read-Host "Y/N"

                    if ($Confirm -eq "Y")
                    {
                    foreach ($Names in $FilteredCourseName)
                    {
                    .\gam update course $Names state archived 
                    }
                    Write-Host "Operation Completed" -ForegroundColor Green
                    $Choice = Read-Host "Would you like to view current active classes? Y/N"
                    if (($Choice -eq "Y") -or ($Choice -eq "Yes"))
                    {
                        $ResultsFile = "\\server73\c$\GAM\_Results\NewAllCoursesRAW.csv"
                        .\gam print courses state active > "\\server73\c$\GAM\_Results\NewAllCoursesRAW.csv"
                        $CSVConvert = ConvertTo-Csv -inputobject $ResultsFile -NoTypeInformation 
                        $NewAllCourseName = Import-Csv $ResultsFile | select name | Select-Object -ExpandProperty name
                        $ImportCSV = Import-Csv $ResultsFile | export-csv "\\server73\c$\GAM\_Results\NewAllCourses.csv" -NoTypeInformation
                        Write-Host "Current active courses" -ForegroundColor Yellow
                        foreach ($NewCourseName in $NewAllCourseName)
                        {
                            Write-Host $NewCourseName
                        }
                        Remove-Item $ResultsFile
                        Write-Host "End of active course list" -ForegroundColor Yellow
                    }
                    else { 
                    }
                        Write-Host "Command Aborted" -ForegroundColor Yellow 
                    }
                    Read-Host "Press 'enter' to continue.."
                    ClassroomMainMenu } 
            'M' {ClassroomMainMenu}
            'Q' {QuitScript}
            }
            
         }
         until (($selection -eq '1') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

#// GMAIL FUNCTIONS // #

    function GmailDelegateMenu {
        param (
            [string]$Title="Gmail Delegate Menu"
        )
        PreMenu
        Clear-Variable Option*
        $Title = "Gmail Delegate Menu"
        $Option1 = "See delegates"
        $Option2 = "Add delegates"
        $Option3 = "Remove delegates"
        $Option4 = "Remove all your delegates[WIP]"
        Show-Menu
        do
         {
            $selection = Read-Host "Please select an option"
            switch ($selection)
            {
            '1' {# Show Delegate
                    $SharedEmail = Read-Host "Show delegates for"
                    Write-Host "Listing delegates for: $SharedEmail" -ForegroundColor Yellow
                    ./Gam user $SharedEmail show delegates
                    Write-Host "All delegates displayed" -ForegroundColor Yellow
                    Read-Host "Press 'enter' to continue.."
                    GmailDelegateMenu
            } 
            '2' {# Add Delegate
                    $SharedEmail = Read-Host "Email to delegate to"
                    $DelegateToAdd = Read-Host "Delegate to give access"
                    ./Gam user $SharedEmail add delegate $DelegateToAdd
                    Write-Host "Processed." -ForegroundColor Yellow
                    Read-Host "Press 'enter' to continue.."
                    GmailDelegateMenu}
            '3' {# Remove Delegate
                    $SharedEmail = Read-Host "Email to delegate to"
                    $DelegateToRemove = Read-Host "Delegate to give access"
                    ./Gam user $SharedEmail delete delegate $DelegateToRemove
                    Write-Host "Processed." -ForegroundColor Yellow
                    Read-Host "Press 'enter' to continue.."
                    GmailDelegateMenu}
            '4' {# Remove all your Delegates
                    #$DelegateToRemove = Read-Host "Remove all delegates from"

                    GmailDelegateMenu}
            'M' {GmailManagementMenu}
            'Q' {QuitScript}
            }
            
         }
         until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

    function BannedUsersMenu {
        param (
            [string]$Title="Banned Users Management"
        )
        Clear-Variable Option*
        $Title = "Banned Users Management"
        $Option1 = "Add a user to the ban list"
        $Option2 = "Remove a user from the ban list"
        $Option3 = "See banned users"
        Show-Menu

        if ($ADGroup -eq "")
        {
            Write-Host "ERROR! ADGroup is empty." -ForegroundColor Red
            $ADGroup = Read-Host "Please type the AD group used for banned users. (It's best to hardcode this)"
            Write-Host "Now using '$ADGroup' as AD group for banned users" -ForegroundColor Green
        }
        if ($GoogleGroup -eq "")
        {
            Write-Host "ERROR! GoogleGroup is empty." -ForegroundColor Red
            $GoogleGroup = Read-Host "Please type the name of the Google group used for banned users. (It's best to hardcode this)"
            Write-Host "Now using '$GoogleGroup' as Google group for banned users" -ForegroundColor Green
        }
        do
         {
            $selection = Read-Host "Please select an option"
            switch ($selection)
            {
            '1' {
                $Email = Read-Host "Email"
                $user = Get-ADGroupMember -Identity $ADGroup | Where-Object {$_.mail -eq $Email}
                if($user){
                Write-Host "$Email is already a banned user! - Exiting" -ForegroundColor Blue 
                }
                else {
                    # Add to AD Group.
                    Write-Host "Adding $Email to AD group.." -ForegroundColor Yellow
                    try {
                        Add-ADGroupMember -Identity $ADGroup -Members $Email
                        Write-Host "Added $Email to AD group.." -ForegroundColor Green
                    }
                    catch {Write-Host "Something went wrong when adding $Email to AD Group: $ADGroup |" $Error[0] -ForegroundColor Red}
                    
                    # Add to Google Group.
                    Write-Host "Adding $Email to Google group.." -ForegroundColor Yellow
                    try {
                        .\gam update group $GoogleGroup add member "$Email"
                        Write-Host "Added $Email Google group.." -ForegroundColor Green
                    }
                    catch {Write-Host "Something went wrong when adding $Email to Google Group: $GoogleGroup |" $Error[0] -ForegroundColor Red}
                    Write-Host "$Email is now banned." -ForegroundColor Green
                }
                Read-Host "Press 'enter' to continue"
                BannedUsersMenu
                } 
            '2' {
                $Email = Read-Host "Email"
                $user = Get-ADGroupMember -Identity $ADGroup -Confirm $false | Where-Object {$_.mail -eq $Email} 
                if(!$user){
                Write-Host "$Email is not a banned user! - Exiting" -ForegroundColor Blue 
                }
                else {
                    # Remove from AD Group.
                    Write-Host "Removing $Email from AD group.." -ForegroundColor Yellow
                    try {
                        Remove-ADGroupMember -Identity $ADGroup -Members $Email
                        Write-Host "Removed $Email from AD group.." -ForegroundColor Green
                    }
                    catch {Write-Host "Something went wrong when removing $Email from AD Group: $ADGroup |" $Error[0] -ForegroundColor Red}
                    
                    # Remove from Google Group.
                    Write-Host "Removing $Email from Google group.." -ForegroundColor Yellow
                    try {
                        .\gam update group $GoogleGroup delete member "$Email"
                        Write-Host "Removed $Email from Google group.." -ForegroundColor Green
                    }
                    catch {Write-Host "Something went wrong when removing $Email from Google Group: $GoogleGroup |" $Error[0] -ForegroundColor Red}
                    Write-Host "$Email is now unbanned." -ForegroundColor Green
                }
                Read-Host "Press 'enter' to continue"
                BannedUsersMenu
                }
            '3' {
                Import-Module ActiveDirectory -Verbose:$false
                $BannedUsers = Get-ADGroupMember -Identity $ADGroup -Recursive | Get-ADUser -Property mail | Select-object -ExpandProperty sAMAccountName
                Write-Host "Found" $BannedUsers.Count "banned users on AD." -ForegroundColor Cyan
                Write-Host "=====================================================" -ForegroundColor Yellow
                foreach ($i in $BannedUsers)
                {
                    Write-Host "$i"
                }
                Write-Host "=====================================================" -ForegroundColor Yellow
                Read-Host "Press 'enter' to continue"
                BannedUsersMenu
                }
            'M' {GmailManagementMenu}
            'Q' {QuitScript}
            }
            
         }
         until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

<# Other functions on the backend #>

    function CourseLookup
        {
        param (
            [string]$Title="Course Lookup"
        )
        $TempOutPath = "$GAMLocation\_Results\Temp.txt"
        .\gam info course $Alias > $TempOutPath
        
        $CourseNameR = (Get-Content -Path "$TempOutPath" -TotalCount 2)[-1]
        $global:CourseName = $CourseNameR.Trim() -replace "name: ", "" 

        $CourseStatusR = (Get-Content -Path "$TempOutPath" -TotalCount 7)[-1]
        $global:CourseStatus = $CourseStatusR.Trim() -replace "courseState: ", "" 

        }

    function DisplayResults
        {
        param (
            [string]$Title="Display Results"
        )
        Write-Host "Displaying results" -ForegroundColor Yellow
        if ($OutPath){
        Import-Csv $OutPath | Format-List}
        else {Write-Host "ERROR: Failed to find an Out file" -ForegroundColor Red} 
        }

    function OpenOutFile
        {
        param (
            [string]$Title="OpenOutFile"
        )
        Write-Host "Opening '$OutPath'" -ForegroundColor Yellow
        Invoke-Command -ScriptBlock {start notepad.exe "$OutPath"}

        }  

    function PreMenu
        {
        param (
            [string]$Title="PreMenu"
        )
        Clear-Variable Option* -Scope Global
        Clear-Variable Command* -Scope Global
        Clear-Variable Command* -Scope Global

        }  

    function QuitScript {
    cls 
    Write-Host "Terminating Script" -ForegroundColor Green
    cd $Location
    Exit
    }

<# Checks & Script #>

    $Hostname = [System.Net.Dns]::GetHostName()
    $Location = Get-Location

    if ($GAMLocation -eq "")
    {
        Write-Host "ERROR! GAMLocation is empty." -ForegroundColor Red
        $GAMLocation = Read-Host "Please type the path of your GAM install. (It's best to hardcode this)"
        Write-Host "Now using '$GAMLocation' as GAM Install Location" -ForegroundColor Green
    }
    if ($GamServerName -eq "")
    {
        Write-Host "ERROR! GamServerName is empty." -ForegroundColor Red
        $GamServerName = Read-Host "Please type the name of the server GAM is installed on. (It's best to hardcode this)"
        Write-Host "Now using '$GamServerName' as GAM Server Name" -ForegroundColor Green
    }

    if (-not(Test-Path "$GAMLocation\_Results")) {
        Write-Host "ERROR!: Missing folder $GAMLocation\_Results" -ForegroundColor Red
    }
    elseif (-not(Test-Path "$GAMLocation\_Results")) {
        Write-Host "ERROR!: Missing folder $GAMLocation\_Config" -ForegroundColor Red
    }
    else {
        if ($Hostname -eq $GamServerName)
        {
            cd "C:\GAM" # Local path of GAM.
            Main-Menu
        }
        else { 

            $Warning = "Yes"
            cd $GAMLocation
            $global:Out = $Null 
            $global:Remote = "-ComputerName $GamServerName" 
            }
            Main-Menu
        }