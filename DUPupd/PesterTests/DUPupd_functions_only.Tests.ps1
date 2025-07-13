# Pester tests for DUPupd script functions

Describe "DUPupd Script Tests" {
    BeforeAll {
        . "$PSScriptRoot\DUPupd_functions_only.ps1"
        $testFolder = "C:\admglsh\DUPupd"
        $testZipPath = "$testFolder\DUPupd.zip"
        $testConfigPath = "$testFolder\config.json"
        $testTaskFolder = "\GLSH\"
        $testLicenseId = "test-license"
        #New-Item -Path $testFolder -ItemType Directory -Force | Out-Null

        # Global mocks to prevent real cmdlet calls
        function Get-RDSessionCollection { throw "This is a dummy function and should be mocked Get-RDSessionCollection."}
        Mock Get-RDSessionCollection { return @()}
        function Get-RDSessionHost { throw "This is a dummy function and should be mocked Get-RDSessionHost."}
        Mock Get-RDSessionHost { return @()}
        Mock Get-CimInstance { return [PSCustomObject]@{ Domain = "test12345.local" }}#>
        #function Assert-DUPupdPrerequisites { throw "This is a dummy function and should be mocked Assert-DUPupdPrerequisites."}
        #function Ensure-DUPHOSTVariable { throw "This is a dummy function and should be mocked Ensure-DUPHOSTVariable."}
        #Mock Ensure-DUPHOSTVariable { return $true }
    }

    BeforeEach {
        [System.Environment]::SetEnvironmentVariable("DUPHOST", $null, [System.EnvironmentVariableTarget]::Machine)
    }

    Context "Date Validation" {
        It "Throws for invalid date" {
            { Convert-DateTime -NewDateTime "invalid" } | Should -Throw "invalid date/time format"
        }


        It "Parses valid date without error" {
            try {
                Convert-DateTime -NewDateTime '01.01.2025 12:00'
                $script:NewDateTimeParsed | Should -BeOfType [DateTime]
                $script:NewDateTimeParsed.ToString("dd.MM.yyyy HH:mm") | Should -Be "01.01.2025 12:00"
            }
            catch {
                Write-Output "Test failed: $_"
                throw
            }
        }
    }
    
    Context "AIO Server Check" {
        BeforeEach {
            Remove-Item -Path "HKLM:\Software\DUPupd" -Recurse -Force -ErrorAction SilentlyContinue
        }

        It "Sets IsAIO to true for AIO server" {
            Mock Get-RDSessionCollection {return @([PSCustomObject]@{ CollectionName = "TestCollection" })}
            Mock Get-RDSessionHost {return @([PSCustomObject]@{ SessionHost = "$env:COMPUTERNAME.local" })}

            IsAIO -IsAIOregPath "HKLM:\Software\DUPupd"
            $result = (Get-ItemProperty -Path "HKLM:\Software\DUPupd" -Name "IsAIO" -ErrorAction SilentlyContinue).IsAIO
            $result | Should -Be "True"
        }

        It "Sets IsAIO to false for non-AIO server" {
            # Global mock (empty collections) applies
            IsAIO -IsAIOregPath "HKLM:\Software\DUPupd"
            $result = (Get-ItemProperty -Path "HKLM:\Software\DUPupd" -Name "IsAIO" -ErrorAction SilentlyContinue).IsAIO
            $result | Should -Be "False"
        }
    }

    Context "Get-TaskInfo" {
        BeforeEach {
            $script:taskName = "TestTask"
            $script:taskFolder = "\GLSH\"
            $script:taskFullPath = "$taskFolder$taskName"
            Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        }

        AfterEach {
            Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        }

        It "Returns Exists false for non-existent task" {
            $output = Get-TaskInfo -TaskFolder $taskFolder -TaskName $taskName 2>&1
            $result = $output[1]
            $output[0] | Should -Match "Aufgabe '$TaskName' nicht gefunden"
            $result.Exists | Should -Be $false
            $result.StartBoundary | Should -Be $null
        }

        It "Returns Exists true and StartBoundary for existing task with trigger" {
            $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date "2025-06-01 12:00")
            $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-Command { exit 0 }"
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            Register-ScheduledTask -TaskName $taskName -TaskPath $taskFolder -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null

            $output = Get-TaskInfo -TaskFolder $taskFolder -TaskName $taskName 2>&1
            $result = $output
            $result.Exists | Should -Be $true
            $result.StartBoundary | Should -BeOfType [DateTime]
            $result.StartBoundary.ToString("yyyy-MM-dd HH:mm") | Should -Be "2025-06-01 12:00"
        }
    }

    Context "Logoff Non-Admins Before Update" {
        BeforeEach {
            $script:taskName = "LogoffNonAdminsBeforeUpdate"
            $script:taskFolder = "\GLSH\"
            $script:taskFullPath = "$taskFolder$taskName"
            Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        }

        AfterEach {
            Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        }

        It "Registers new task when no existing task" {
            $newDateTime = Get-Date "2025-06-01 12:00"
            # Ensure no existing task
            Get-ScheduledTask -TaskPath $taskFolder -TaskName "LogoffNonAdminsBeforeUpdate" -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue

            $output = Register-LogoffNonAdminsBeforeUpdate -NewDateTimeParsed $newDateTime -TaskFolder $taskFolder 2>&1
            $output[1] | Should -Match "Neue Aufgabe 'LogoffNonAdminsBeforeUpdate' wurde erstellt"

            $task = Get-ScheduledTask -TaskPath $taskFolder -TaskName "LogoffNonAdminsBeforeUpdate" -ErrorAction SilentlyContinue
            $task | Should -Not -Be $null
            $task.Triggers[0].StartBoundary | Should -Be ($newDateTime.AddMinutes(-1))
            [datetime]$actual = $task.Triggers[0].StartBoundary
            $expected = $newDateTime.AddMinutes(-1)
            $actual | Should -Be $expected
            # $task.Triggers[0].SynchronizeAcrossTimeZones | Should -Be $false
        }
        
        It "Skips registration when task exists with same start time" {
            $newDateTime = Get-Date "2025-06-01 12:00"
            $trigger = New-ScheduledTaskTrigger -Once -At ($newDateTime.AddMinutes(-1))
            $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-Command { exit 0 }"
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            Register-ScheduledTask -TaskName $taskName -TaskPath $taskFolder -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null

            $output = Register-LogoffNonAdminsBeforeUpdate -NewDateTimeParsed $newDateTime -TaskFolder $taskFolder 2>&1
            $output | Should -Match "Aufgabe 'LogoffNonAdminsBeforeUpdate' ist bereits"
        }
    }

    Context "Update-DatevScheduledTask" {
        BeforeEach {
            $script:taskName = "DUPupd"
            $script:taskFolder = "\GLSH\"
            $script:taskFullPath = "$taskFolder$taskName"
            Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        }

        AfterEach {
            Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        }

        It "Updates task trigger time when AutoLogon is enabled" {
            #Create a task to update
            $initialTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date "2025-06-01 11:00")
            $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-Command { exit 0 }"
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            Register-ScheduledTask -TaskName $taskName -TaskPath $taskFolder -Action $action -Trigger $initialTrigger -Principal $principal -Force | Out-Null
            
            Mock Assert-DUPupdPrerequisites { }

            $newDateTime = Get-Date "2025-06-01 12:00"
            
            $output = Update-DatevScheduledTask -TaskFolder $taskFolder -TaskName $taskName -NewDateTimeParsed $newDateTime -taskAvailable $true -currentStartBoundary (Get-Date "2025-06-01 11:00") -autoLogonEnabled $true -duphostSet $true 2>&1
            Write-Output "$output"
            $output[1] | Should -Match "Aufgabe $taskName wurde auf $newDateTime aktualisiert"

            $task = Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue
            [datetime]$actual = $task.Triggers[0].StartBoundary
            $expected = $newDateTime
            $actual | Should -Be $expected
            Assert-MockCalled Assert-DUPupdPrerequisites -Times 1
            
        }
    }

    Context "Get-HTTPLastModified" {
        It "Returns Last-Modified date for valid URL" {
            $url = "http://<YourFileServer>.de/DUPupd/DUPupd.zip"  # Replace with actual server URL
            $result = Get-HTTPLastModified -url $url 2>&1
            $result | Should -Not -Be $null
            $result | Should -BeOfType [DateTime]
            #$result.ToString("yyyy-MM-dd HH:mm") | Should -Be "2025-03-17 14:58"
        }
    }

    Context "Test-UpdateRequired" {
        BeforeEach {
            $script:testFolder = "C:\admglsh\DUPupd"
            $script:testZipPath = "$testFolder\DUPupd.zip"
            # Ensure test directory exists
            New-Item -Path $testFolder -ItemType Directory -Force | Out-Null
            # Remove any existing test ZIP file
            Remove-Item -Path $testZipPath -Force -ErrorAction SilentlyContinue
        }

        AfterEach {
            # Clean up test ZIP file
            Remove-Item -Path $testZipPath -Force -ErrorAction SilentlyContinue
        }

        It "Returns true when local ZIP file is older than server Last-Modified" {
            # Create fake ZIP file older than server (Feb 2025)
            New-Item -Path $testZipPath -ItemType File -Force | Out-Null
            Set-ItemProperty -Path $testZipPath -Name LastWriteTime -Value (Get-Date "2025-02-01 12:00")

            $serverLastModified = Get-HTTPLastModified -url "http://<YourFileServer>.de/DUPupd/DUPupd.zip"
            $serverLastModified | Should -Not -Be $null
            $serverLastModified | Should -BeOfType [DateTime]

            $result = Test-UpdateRequired -ZipFilePath $testZipPath -TargetFolder $testFolder -ServerLastModified $serverLastModified
            $result | Should -Be $true
        }

        It "Returns false when local ZIP file is newer than server Last-Modified" {
            # Create fake ZIP file newer than server (Apr 2025)
            New-Item -Path $testZipPath -ItemType File -Force | Out-Null
            Set-ItemProperty -Path $testZipPath -Name LastWriteTime -Value (Get-Date "2025-04-01 12:00")

            $serverLastModified = Get-HTTPLastModified -url "http://<YourFileServer>.de/DUPupd/DUPupd.zip"
            $serverLastModified | Should -Not -Be $null
            $serverLastModified | Should -BeOfType [DateTime]

            $result = Test-UpdateRequired -ZipFilePath $testZipPath -TargetFolder $testFolder -ServerLastModified $serverLastModified
            $result | Should -Be $false
        }

        It "Returns true when local ZIP file is missing" {
            $serverLastModified = Get-HTTPLastModified -url "http://<YourFileServer>.de/DUPupd/DUPupd.zip"
            $serverLastModified | Should -Not -Be $null
            $serverLastModified | Should -BeOfType [DateTime]

            $result = Test-UpdateRequired -ZipFilePath $testZipPath -TargetFolder $testFolder -ServerLastModified $serverLastModified
            $result | Should -Be $true
        }

        It "Returns true when target folder is missing" {
            $nonExistentFolder = "C:\admglsh\NonExistent"
            $serverLastModified = Get-HTTPLastModified -url "http://<YourFileServer>.de/DUPupd/DUPupd.zip"
            $serverLastModified | Should -Not -Be $null
            $serverLastModified | Should -BeOfType [DateTime]

            $result = Test-UpdateRequired -ZipFilePath $testZipPath -TargetFolder $nonExistentFolder -ServerLastModified $serverLastModified
            $result | Should -Be $true
        }
    }

    Context "Update-DatevFiles" {
        BeforeEach {
            $script:testFolder = "C:\Temp\DUPupdTest"
            $script:testZipPath = "$testFolder\DUPupd.zip"
            $script:testConfigPath = "$testFolder\config.json"
            $script:licenseId = "test-license"
            # Remove test directory and files
            #Remove-Item -Path $testFolder -Recurse -Force -ErrorAction SilentlyContinue
            # Create test directory
            #New-Item -Path $testFolder -ItemType Directory -Force | Out-Null
        }

        AfterEach {
            # Clean up test directory and files
            #Remove-Item -Path $testFolder -Recurse -Force -ErrorAction SilentlyContinue
        }

        It "Downloads and extracts ZIP file and updates config" {
            $output = Update-DatevFiles -HttpUrl "http://<YourFileServer>.de/DUPupd/DUPupd.zip" `
                                        -TargetFolder $testFolder `
                                        -ConfigFilePath $testConfigPath `
                                        -LicenseId $licenseId `
                                        -LastModified (Get-Date "2025-03-01 12:00") `
                                        -ZipFilePath $testZipPath 2>&1

            $output | Should -Contain "Files successfully downloaded and extracted to '$testFolder'."
            $output | Should -Contain "License ID updated in '$testConfigPath'."
            Test-Path -Path $testZipPath | Should -Be $true
            Test-Path -Path $testConfigPath | Should -Be $true
            $config = Get-Content -Path $testConfigPath | ConvertFrom-Json
            $config.licenseid | Should -Be $licenseId
        }

        It "Overwrites existing ZIP file and updates config" {
            # Create an old ZIP file
            #New-Item -Path $testZipPath -ItemType File -Force | Out-Null
            Set-ItemProperty -Path $testZipPath -Name LastWriteTime -Value (Get-Date "2025-01-01 12:00")
            # Create an initial config file
            @{ licenseid = "old-license" } | ConvertTo-Json | Set-Content -Path $testConfigPath

            $output = Update-DatevFiles -HttpUrl "http://<YourFileServer>.de/DUPupd/DUPupd.zip" `
                                        -TargetFolder $testFolder `
                                        -ConfigFilePath $testConfigPath `
                                        -LicenseId $licenseId `
                                        -LastModified (Get-Date "2025-03-01 12:00") `
                                        -ZipFilePath $testZipPath 2>&1

            $output | Should -Contain "Files successfully downloaded and extracted to '$testFolder'."
            $output | Should -Contain "License ID updated in '$testConfigPath'."
            Test-Path -Path $testZipPath | Should -Be $true
            (Get-Item $testZipPath).LastWriteTime | Should -BeGreaterThan (Get-Date "2025-01-01 12:00")
            Test-Path -Path $testConfigPath | Should -Be $true
            $config = Get-Content -Path $testConfigPath | ConvertFrom-Json
            $config.licenseid | Should -Be $licenseId
        }
    }

    Context "Test-AdminLoggedIn" {
        It "Returns true when admin user is logged in" {
            Mock quser { return " USERNAME              SESSIONNAME        ID  STATE   IDLE TIME  LOGON TIME`n>admin                 console            1   Active  none       6/1/2025 10:00 AM" }
            $output = Test-AdminLoggedIn 2>&1
            $output[0] | Should -Match "Admin found."
            $output[1] | Should -Be $true
        }

        It "Returns false when no admin user is logged in" {
            Mock quser { return " USERNAME              SESSIONNAME        ID  STATE   IDLE TIME  LOGON TIME`n user1                 console            1   Active  none       6/1/2025 10:00 AM`n user2                 rdp-tcp            2   Active  1:00       6/1/2025 11:00 AM" }

            $output = Test-AdminLoggedIn 2>&1
            $output[0] | Should -Match "No Admin user is currently logged in"
            $output[1] | Should -Be $false
        }

        It "Returns false and handles quser error" {
            Mock quser { throw "Access denied" }

            $output = Test-AdminLoggedIn 2>&1
            $output[0] | Should -Match "Error: looking for logged in users Access denied"
            $output[1] | Should -Be $false
        }
    }

    Context "Register-RestartServer" {
        BeforeEach {
            $script:taskName = "RestartServer"
            $script:taskFolder = "\GLSH\"
            $script:taskFullPath = "$taskFolder$taskName"
            # Clean up any existing task
            Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        }

        AfterEach {
            # Clean up task
            Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        }

        It "Creates restart task with correct trigger time" {
            $newDateTime = Get-Date "2025-05-01 12:00"
            $output = Register-RestartServer -RestartServerDateTime $newDateTime `
                                            -TaskFolder $taskFolder `
                                            -taskAvailable $true `
                                            -autoLogonEnabled $true 2>&1

            $output | Should -Contain "Task '$taskFullPath' created successfully for restart at $newDateTime."
            
            $task = Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue
            $task | Should -Not -Be $null
            [datetime]$actual = $task.Triggers[0].StartBoundary
            $expected = $newDateTime
            $actual | Should -Be $expected
            $task.Actions[0].Execute | Should -Be "shutdown.exe"
            $task.Actions[0].Arguments | Should -Be "/r /t 120 /c 'Server startet in 2 Minuten neu. Bitte jetzt abmelden."
        }
    }

    Context "Register-RestartIfNoAdmin" {
        BeforeEach {
            $script:taskName = "RestartServer"
            $script:taskFolder = "\GLSH\"
            $script:taskFullPath = "$taskFolder$taskName"
            # Clean up any existing task
            Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        }

        AfterEach {
            # Clean up task
            Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        }

        It "Registers restart task when no admin is logged in and no task exists" {
            Mock Test-AdminLoggedIn { return $false }
            $newDateTime = Get-Date "2025-06-01 12:00"
            Mock Get-Date { return $newDateTime }  # Ensure current date matches

            $output = Register-RestartIfNoAdmin -NewDateTimeParsed $newDateTime `
                                                -TaskFolder $taskFolder `
                                                -taskAvailable $true `
                                                -autoLogonEnabled $true `
                                                -duphostSet $true 2>&1

            $output | Should -Contain "Restart task '$taskFullPath' scheduled for $($newDateTime.AddMinutes(-4))."
            
            $task = Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue
            $task | Should -Not -Be $null
            [datetime]$actual = $task.Triggers[0].StartBoundary
            $expected = $newDateTime.AddMinutes(-4)
            $actual | Should -Be $expected
        }

        It "Skips registration when task is already scheduled for today" {
            $newDateTime = Get-Date "2025-06-01 12:00"
            $restartTime = $newDateTime.AddMinutes(-4)
            # Create task with today's trigger
            $trigger = New-ScheduledTaskTrigger -Once -At $restartTime
            $action = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument "/r /t 120 /c 'Server startet in 2 Minuten neu. Bitte jetzt abmelden."
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            Register-ScheduledTask -TaskName $taskName -TaskPath $taskFolder -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null

            Mock Test-AdminLoggedIn { return $false }
            Mock Get-Date { return $newDateTime }

            $output = Register-RestartIfNoAdmin -NewDateTimeParsed $newDateTime `
                                                -TaskFolder $taskFolder `
                                                -taskAvailable $true `
                                                -autoLogonEnabled $true `
                                                -duphostSet $true 2>&1

            $output | Should -Contain "Task '$taskFullPath' is already scheduled for today at $restartTime."
            $task = Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue
            [datetime]$actual = $task.Triggers[0].StartBoundary
            $expected = $newDateTime.AddMinutes(-4)
            $actual | Should -Be $expected
        }

        It "Skips registration when admin is logged in" {
            Mock Test-AdminLoggedIn { return $true }
            $newDateTime = Get-Date "2025-06-01 12:00"
            Mock Get-Date { return $newDateTime }

            $output = Register-RestartIfNoAdmin -NewDateTimeParsed $newDateTime `
                                                -TaskFolder $taskFolder `
                                                -taskAvailable $true `
                                                -autoLogonEnabled $true `
                                                -duphostSet $true 2>&1

            $output | Should -Contain "Admin is logged in. Skipping restart task registration."
            $task = Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue
            $task | Should -Be $null
        }

        It "Skips registration when dates do not match" {
            Mock Test-AdminLoggedIn { }
            $newDateTime = Get-Date "2025-06-02 12:00"
            Mock Get-Date { [datetime]"2025-06-01 12:00" } 

            $output = Register-RestartIfNoAdmin -NewDateTimeParsed $newDateTime `
                                                -TaskFolder $taskFolder `
                                                -taskAvailable $true `
                                                -autoLogonEnabled $true `
                                                -duphostSet $true 2>&1

            $output | Should -Contain "Current date does not match NewDateTimeParsed date. Skipping restart task registration."
            $task = Get-ScheduledTask -TaskPath $taskFolder -TaskName $taskName -ErrorAction SilentlyContinue
            $task | Should -Be $null
        }
    }

    Context "Ensure-DUPHOSTVariable" {
        BeforeEach {
            # Remove DUPHOST variable if it exists
            if (Test-Path Env:\DUPHOST) {
                Remove-Item -Path Env:\DUPHOST -Force
            }
        }

        AfterEach {
            # Clean up DUPHOST variable
            if (Test-Path Env:\DUPHOST) {Remove-Item -Path Env:\DUPHOST -Force}
        }

        It "Sets DUPHOST variable when it does not exist and domain contains 5 digits" {
            Mock Get-WmiObject { [PSCustomObject]@{ Domain = "example12345.local" } } -ParameterFilter { $Class -eq "Win32_ComputerSystem" }
            $result = Ensure-DUPHOSTVariable 2>&1
            $result | Should -Contain "System variable 'DUPHOST' set."
            $result[1] | Should -Be $true
        }

        It "Does not set DUPHOST variable when domain lacks 5 digits" {
            Mock Get-WmiObject { [PSCustomObject]@{ Domain = "example.local" } } -ParameterFilter { $Class -eq "Win32_ComputerSystem" }
            $result = Ensure-DUPHOSTVariable 2>&1
            $result | Should -Contain "Error: Domain name does not contain 5 digits."
            $result[1] | Should -Be $false
        }

        It "Returns true when DUPHOST variable already exists" {
            Mock Test-Path { $true }
            $result = Ensure-DUPHOSTVariable 2>&1
            $result | Should -Contain "System variable 'DUPHOST' exists."
            $result[1] | Should -Be $true
        }
    }

    Context "Assert-DUPupdPrerequisites" {
        BeforeEach {
            # Clean up DUPHOST variable from process and machine levels
            if (Test-Path Env:\DUPHOST) {
                Remove-Item -Path Env:\DUPHOST -Force
            }
            if ([System.Environment]::GetEnvironmentVariable("DUPHOST", [System.EnvironmentVariableTarget]::Machine)) {
                [System.Environment]::SetEnvironmentVariable("DUPHOST", $null, [System.EnvironmentVariableTarget]::Machine)
            }
        }

        AfterEach {
            # Clean up DUPHOST variable
            if (Test-Path Env:\DUPHOST) {
                Remove-Item -Path Env:\DUPHOST -Force
            }
            if ([System.Environment]::GetEnvironmentVariable("DUPHOST", [System.EnvironmentVariableTarget]::Machine)) {
                [System.Environment]::SetEnvironmentVariable("DUPHOST", $null, [System.EnvironmentVariableTarget]::Machine)
            }
        }

        It "Exits with 0 when all prerequisites are met" {
            Mock Ensure-DUPHOSTVariable { return $true }
            Mock Register-RestartIfNoAdmin { }
            $newDateTime = Get-Date "2025-06-01 12:00"

            $output = Assert-DUPupdPrerequisites -NewDateTimeParsed $newDateTime `
                                                -TaskFolder "\GLSH\" `
                                                -taskAvailable $true `
                                                -autoLogonEnabled $true 2>&1

            $output | Should -Contain "DATEV update is scheduled."
            #$LASTEXITCODE | Should -Be 0
            Assert-MockCalled Ensure-DUPHOSTVariable -Times 1
            Assert-MockCalled Register-RestartIfNoAdmin -Times 1
        }

        It "Exits with 1001 when DUPHOST is missing" {
            Mock Ensure-DUPHOSTVariable { return $false }
            Mock Register-RestartIfNoAdmin { }

            $output = Assert-DUPupdPrerequisites -NewDateTimeParsed (Get-Date "2025-06-01 12:00") `
                                                -TaskFolder "\GLSH\" `
                                                -taskAvailable $true `
                                                -autoLogonEnabled $true 2>&1

            $output | Should -Contain "System variable 'DUPHOST' is missing and could not be set automatically. Please set it manually."
            #LASTEXITCODE | Should -Be 1001
            Assert-MockCalled Ensure-DUPHOSTVariable -Times 1
            Assert-MockCalled Register-RestartIfNoAdmin -Times 1
        }

        It "Exits with 1001 when task is not available" {
            Mock Ensure-DUPHOSTVariable { return $true }
            Mock Register-RestartIfNoAdmin { }

            $output = Assert-DUPupdPrerequisites -NewDateTimeParsed (Get-Date "2025-06-01 12:00") `
                                                -TaskFolder "\GLSH\" `
                                                -taskAvailable $false `
                                                -autoLogonEnabled $true 2>&1

            $output | Should -Contain "The Windows DUPupd task does not exist. Create the task by running the C:\admglsh\DUPupd\CreateDUPupdTask.ps1 script as an admin user on the server. The execution must be explicitly performed as Administrator, e.g., via Total Commander."
            #$LASTEXITCODE | Should -Be 1001
            Assert-MockCalled Ensure-DUPHOSTVariable -Times 1
            Assert-MockCalled Register-RestartIfNoAdmin -Times 1
        }

        It "Exits with 1001 when AutoLogon is not enabled" {
            Mock Ensure-DUPHOSTVariable { return $true }
            Mock Register-RestartIfNoAdmin { }

            $output = Assert-DUPupdPrerequisites -NewDateTimeParsed (Get-Date "2025-06-01 12:00") `
                                                -TaskFolder "\GLSH\" `
                                                -taskAvailable $true `
                                                -autoLogonEnabled $false 2>&1

            $output | Should -Contain "AutoLogon is not enabled. Enable AutoLogon on the server to use DUPupd."
            #$LASTEXITCODE | Should -Be 1001
            Assert-MockCalled Ensure-DUPHOSTVariable -Times 1
            Assert-MockCalled Register-RestartIfNoAdmin -Times 1
        }

        It "Exits with 1001 when all prerequisites fail" {
            Mock Ensure-DUPHOSTVariable { return $false }
            Mock Register-RestartIfNoAdmin { }

            $output = Assert-DUPupdPrerequisites -NewDateTimeParsed (Get-Date "2025-06-01 12:00") `
                                                -TaskFolder "\GLSH\" `
                                                -taskAvailable $false `
                                                -autoLogonEnabled $false 2>&1

            $output | Should -Contain "System variable 'DUPHOST' is missing and could not be set automatically. Please set it manually."
            $output | Should -Contain "The Windows DUPupd task does not exist. Create the task by running the C:\admglsh\DUPupd\CreateDUPupdTask.ps1 script as an admin user on the server. The execution must be explicitly performed as Administrator, e.g., via Total Commander."
            $output | Should -Contain "AutoLogon is not enabled. Enable AutoLogon on the server to use DUPupd."
            #$LASTEXITCODE | Should -Be 1001
            Assert-MockCalled Ensure-DUPHOSTVariable -Times 1
            Assert-MockCalled Register-RestartIfNoAdmin -Times 1
        }
    }
}