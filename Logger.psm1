<#
https://overpoweredshell.com/Introduction-to-PowerShell-Classes/
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_classes?view=powershell-7.3
https://powershell.one/powershell-internals/attributes/validation
https://learn.microsoft.com/en-us/powershell/scripting/developer/module/how-to-write-a-powershell-module-manifest?view=powershell-7.3
Place this module (PSM1 and PS1) in one of the following folders:
   C:\Users\%User%\OneDrive\Documents\WindowsPowerShell\Modules\Logger
   C:\Users\%User%\Documents\WindowsPowerShell\Modules\Logger
#>

Function Test-IsFileLocked {
    <#
    https://mcpmag.com/articles/2018/07/10/check-for-locked-file-using-powershell.aspx
    #>
    [cmdletbinding()]
    Param (
        [parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('FullName','PSPath')]
        [string[]]$Path
    )
    Process {
        ForEach ($Item in $Path) {
            #Ensure this is a full path
            #$Item = Convert-Path $Item
            #Verify that this is a file and not a directory
            If ([System.IO.File]::Exists($Item)) {
                Try {
                    $FileStream = [System.IO.File]::Open($Item,'Open','Write')
                    $FileStream.Close()
                    $FileStream.Dispose()
                    $IsLocked = $False
                } Catch [System.UnauthorizedAccessException] {
                    $IsLocked = 'AccessDenied'
                } Catch {
                    $IsLocked = $True
                }
                [pscustomobject]@{
                    File = $Item
                    IsLocked = $IsLocked
                }
            }
        }
    }
}

class LoggingLevel
{
    [ValidateSet ("Error", "Warning", "Information", "Debug", "Verbose", "Trace")]
    [string]
    $LoggingLevelName

    [ValidateRange(0,4)]
    [int]
    $LoggingLevelInt

    LoggingLevel(){}

    LoggingLevel(
        [string]$LevelName
    )
    {
        $this.LoggingLevelName = $LevelName
        $this.LoggingLevelInt = $this.LevelToInt($LevelName)
    }

    [int]LevelToInt($LevelName)
    {
        # Convert logging level to INT
        switch ($LevelName)
        {
            "Error" 
            {
                return 0
            }

            "Warning" 
            {
                return 1
            }

            "Information" 
            {
                return 2
            }

            "Debug" 
            {
                return 3
            }

            "Verbose" 
            {
                return 4
            }

            "Trace" 
            {
                return 5
            }
        }
        return -1
    }

    [int]Level()
    {
        return $this.LoggingLevelInt
    }

    [string]Name()
    {
        return $this.LoggingLevelName
    }
}

class LoggingDestination
{
    [ValidateSet ("Console", "File", "Both")]
    [string]
    $LogDestination

    LoggingDestination()
    {
        $this.LogDestination = "Both"
    }

    LoggingDestination(
        [string]$Destination
    )
    {
        $this.LogDestination = $Destination
    }
    [bool]Console()
    {
        if (($this.LogDestination -ieq "Console") -or ($this.LogDestination -ieq "Both"))
        {
            return $true
        }
        else
        {
            return $false
        }
    }

    [bool]File()
    {
        if (($this.LogDestination -ieq "File") -or ($this.LogDestination -ieq "Both"))
        {
            return $true
        }
        else
        {
            return $false
        }
    }
}

class LoggingObject
{
    [string]
    hidden $LoggingFile

    [string]
    hidden $LoggingMutexName

    [System.Threading.Mutex]
    hidden $loggingMutex

    [timespan]
    hidden $Timeout

    # Constructor
    LoggingObject(){}
    LoggingObject(
        [string]$LogFileUNC,
        [int]$TimeOutSeconds
    )
    {
        $this.Timeout = New-TimeSpan -Seconds $TimeOutSeconds
        $this.LoggingFile = $LogFileUNC
        $this.LoggingMutexName = Split-Path -Leaf $LogFileUNC
        try
        {
            $this.loggingMutex = [System.Threading.Mutex]::OpenExisting($this.LoggingMutexName)
            $this.loggingMutex.ReleaseMutex()
            Write-Debug "Opened EXISTING handle."
        }
        catch
        {
            $this.loggingMutex = New-Object System.Threading.Mutex($false, $this.LoggingMutexName)
            Write-Debug "Opened NEW handle."
        }
        
        
    }
    
    # Methods
    [string]GetLoggingFile()
    {
        return $this.LoggingFile
    }

    [void]SetLoggingFile(
        [string]$LogFileUNC
    )
    {
        $this.LoggingFile = $LogFileUNC
    }

    [System.Threading.Mutex] GetLoggingMutex()
    {
        return $this.loggingMutex
    }

    [string] GetLoggingMutexName()
    {
        return $this.LoggingMutexName
    }

    [void] SetTimeOutSeconds(
        [int]$TimeOutSeconds
    )
    {
        $this.Timeout = $TimeOutSeconds
    }

    [void]WriteConsole(
        [string]$Message,
        [LoggingLevel]$Level
    )
    {
        # Pick the console write type.
        if ($Level.Name() -ieq "Trace")
        {
            Write-Host $Message  -ForegroundColor Green -BackgroundColor DarkGrey
        }
        elseif ($Level.Name() -ieq "Verbose")
        {
            Write-Host $Message  -ForegroundColor Green -BackgroundColor Black
        }
        elseif ($Level.Name() -ieq "Debug")
        {
            Write-Host $Message -ForegroundColor DarkBlue -BackgroundColor Black
        }
        elseif ($Level.Name() -ieq "Information")
        {
            Write-Host $Message -ForegroundColor Gray -BackgroundColor Black
        }
        elseif ($Level.Name() -ieq "Warning")
        {
            Write-Host $Message -ForegroundColor DarkYellow -BackgroundColor DarkGray
        }
        elseif ($Level.Name() -ieq "Error")
        {
            Write-Host $Message -ForegroundColor DarkRed -BackgroundColor White
        }
        else
        {
            Write-Host $Message
        }
        
    }

    [void]WriteFile(
        [string]$Message
    )
    {
        # Write to the log file
        #  An odd issue keeps popping up where the file cannot be written to even though nothing else should
        #  be accessing the file.
        #  The mutex is being used to help with parallel processes writing to the file.
        #  The file locked check is for the odd error.  An external process is locking the file for some reason.
        #https://social.technet.microsoft.com/Forums/SECURITY/en-US/0c0e36a2-4935-4385-af79-9d7576a8aac5/appending-to-file-getting-error-file-is-being-used-by-another-process?forum=winserverpowershell

        #Try to add the content to the log file.
		<#
            while((!$this.loggingMutex.WaitOne(7000)) -and ((Test-IsFileLocked -Path $this.LoggingFile).IsLocked))
			{
                $_WaitAttempts += 1
				Write-Host ("{0} | Waiting for log file access." -f $(Get-Date -Format "yyyy-dd-MM hh:mm:ss"))
                Write-Debug ("Mutex Handle ({0}) | Mutex Closed ({1}) | Mutex Name ({2}) | " -f $this.loggingMutex.Handle, $this.loggingMutex.SafeWaitHandle.IsClosed, $this.LoggingMutexName)
			}
			if($this.loggingMutex.WaitOne())
			{
                $_WaitAttempts += 1
				# Test for external file lock.
                while((Test-IsFileLocked -Path $this.LoggingFile).IsLocked)
                {
                    Write-Host ("{0} | Log file locked.  Waiting" -f $(Get-Date -Format "yyyy-dd-MM hh:mm:ss"))
                }
                #  Write the line once we get the Write
				Add-Content -Path $this.LoggingFile -Value $Message -ErrorAction Stop
			}
            # Capture the current date and time for the timeout
            

            finally
                        {
                            #Write-Host ("Mutex Waits:  {0}" -f $_WaitAttempts)
                            if ($_WaitAttempts -gt 1)
                            {
                                for (($i = 0); $i -le $_WaitAttempts; $i++)
                                {
                                    $this.loggingMutex.ReleaseMutex()
                                }
                            }
                        }
        #>
        # Keep track of the Wait attempts
        $_WaitAttempts = 0

        # Get the current date and time for the timeout
        $_AttemptWriteStartTime = Get-Date
        
        #Add the entry to the log file.  If it doesn't exist we'll create it.
        try {
			# Loop until we write successfully or the timeout is reached.
            while($true)
            {
                # Check if the file is locked.
                if(-not (Test-IsFileLocked -Path $this.LoggingFile).IsLocked)
                {
                    # Get the mutex for the log file.  For parallel processing
                    if ($this.loggingMutex.WaitOne(700))
                    {
                        $_WaitAttempts += 1

                        # Try writing
                        try
                        {
                            # At this point the file should be ours to write to.
                            #  Write the line once we get the Write
				            Add-Content -Path $this.LoggingFile -Value $Message -ErrorAction Stop
                            $this.loggingMutex.ReleaseMutex()
                            break
                        }
                        catch
                        {
                            # A write failure occured.  Wait and restart the loop.
                            Start-Sleep -Milliseconds 50
                        }
                        

                        # Check to see if the timeout has been exceeded and exit
                        if (($null -ne $this.Timeout) -and ((Get-Date) -gt ($_AttemptWriteStartTime + $this.Timeout)))
                        {
                            # The timeout has expired.  Throw an exception.
                            throw "Timeout Expired."
                        }
                    }
                    else
                    {
                        $this.loggingMutex.ReleaseMutex()
                    }
                }
                else
                {
                    Start-Sleep -Milliseconds 50
                }
            }
			
		}
		catch {
			throw $("Could not write to log file: {0}" -f $_.exception.message)
		}
    }
}

class Logger
{
    # Class Variables
    [LoggingLevel]
    hidden $TargetLogLevel

    [LoggingObject]
    hidden $LogFileObj

    [string]
    hidden $LogFileName

    [string]
    hidden $LogFilePath

    [int]
    hidden $TimeOut

    # Constructors
    Logger(){
        # Set the default logging level to:
        $this.TargetLogLevel = [LoggingLevel]::new("Verbose")

        # Set the default log file name:
        #  Get the name of the running script
        $scriptName = (Split-Path -Leaf $PSCommandPath) -replace '\.ps1$|\.psm1$'
        
        #  Set the Log File Name
        $this.LogFileName = ('{0}_{1}.log' -f $(Get-Date -Format "yyyy-MM-dd"), $scriptName)

        # Set the log file object
        #  Get the path to the currently running script.
        $this.LogFilePath = Split-Path -Parent $PSCommandPath

        # Set the timeout
        $this.TimeOut = 5

        #  Set the Log File Object
        $this.LogFileObj = [LoggingObject]::new((Join-Path $this.LogFilePath $this.LogFileName), $this.TimeOut)

    }

    Logger(
        [string]$TargetLogLevel
    )
    {
        # Set the default logging level to:
        $this.TargetLogLevel = [LoggingLevel]::new($TargetLogLevel)

        # Set the default log file name:
        #  Get the name of the running script
        $scriptName = (Split-Path -Leaf $PSCommandPath) -replace '\.ps1$|\.psm1$'
        
        #  Set the Log File Name
        $this.LogFileName = ('{0}_{1}.log' -f $(Get-Date -Format "yyyy-MM-dd"), $scriptName)

        # Set the log file object
        #  Get the path to the currently running script.
        $this.LogFilePath = Split-Path -Parent $PSCommandPath

        # Set the timeout
        $this.TimeOut = 5

        #  Set the Log File Object
        $this.LogFileObj = [LoggingObject]::new((Join-Path $this.LogFilePath $this.LogFileName), $this.TimeOut)
    }

    Logger(
        [string]$TargetLogLevel,
        [string]$LogFileName
    ){
        # Set the default logging level to:
        $this.TargetLogLevel = [LoggingLevel]::new($TargetLogLevel)

        # Set the default log file name:
        #  Set the Log File Name
        if ($LogFileName -eq "")
        {
            #  Get the name of the running script
            $scriptName = (Split-Path -Leaf $PSCommandPath) -replace '\.ps1$|\.psm1$'
            $this.LogFileName = ('{0}_{1}.log' -f $(Get-Date -Format "yyyy-MM-dd"), $scriptName)
        }
        else
        {
            $this.LogFileName = ('{0}_{1}.log' -f $(Get-Date -Format "yyyy-MM-dd"), $LogFileName)
        }
        
        # Set the log file object
        #  Get the path to the currently running script.
        $this.LogFilePath = Split-Path -Parent $PSCommandPath

        # Set the timeout
        $this.TimeOut = 5

        #  Set the Log File Object
        $this.LogFileObj = [LoggingObject]::new((Join-Path $this.LogFilePath $this.LogFileName), $this.TimeOut)
    }

    Logger(
        [string]$TargetLogLevel,
        [string]$LogFileName,
        [string]$LogFolder
    ){
        # Set the default logging level to:
        $this.TargetLogLevel = [LoggingLevel]::new($TargetLogLevel)

        # Set the default log file name:
        if ($LogFileName -eq "")
        {
            #  Get the name of the running script
            $scriptName = (Split-Path -Leaf $PSCommandPath) -replace '\.ps1$|\.psm1$'
            $this.LogFileName = ('{0}_{1}.log' -f $(Get-Date -Format "yyyy-MM-dd"), $scriptName)
        }
        else
        {
            $this.LogFileName = ('{0}_{1}.log' -f $(Get-Date -Format "yyyy-MM-dd"), $LogFileName)
        }

        # Set the log file object
        #  Get the path to the currently running script.
        $scriptPath = Split-Path -Parent $PSCommandPath

        #  Check if the provided path is fully qualified or relative.
        if ([System.IO.Path]::IsPathRooted($LogFolder))
        {
            #  The provided path is absolute/fully qualified
            $this.LogFilePath = $LogFolder
        }
        else 
        {
            $this.LogFilePath = Join-Path $scriptPath $LogFolder
        }

        # Set the timeout
        $this.TimeOut = 5

        #  Set the Log File Object
        $this.LogFileObj = [LoggingObject]::new((Join-Path $this.LogFilePath $this.LogFileName), $this.TimeOut)
    }

    Logger(
        [string]$TargetLogLevel,
        [string]$LogFileName,
        [string]$LogFolder,
        [bool]$UseFileNameAsIs
    ){
        # Set the default logging level to:
        $this.TargetLogLevel = [LoggingLevel]::new($TargetLogLevel)

        # Set the default log file name:
        if ($LogFileName -eq "")
        {
            #  Get the name of the running script
            $scriptName = (Split-Path -Leaf $PSCommandPath) -replace '\.ps1$|\.psm1$'
            $this.LogFileName = ('{0}_{1}.log' -f $(Get-Date -Format "yyyy-MM-dd"), $scriptName)
        }
        elseif ($UseFileNameAsIs)
        {
            $this.LogFileName = $LogFileName
        }
        else
        {
            $this.LogFileName = ('{0}_{1}.log' -f $(Get-Date -Format "yyyy-MM-dd"), $LogFileName)
        }

        # Set the log file object
        #  Get the path to the currently running script.
        $scriptPath = Split-Path -Parent $PSCommandPath

        #  Check if the provided path is fully qualified or relative.
        if ([System.IO.Path]::IsPathRooted($LogFolder))
        {
            #  The provided path is absolute/fully qualified
            $this.LogFilePath = $LogFolder
        }
        else 
        {
            $this.LogFilePath = Join-Path $scriptPath $LogFolder
        }

        # Set the timeout
        $this.TimeOut = 5

        #  Set the Log File Object
        $this.LogFileObj = [LoggingObject]::new((Join-Path $this.LogFilePath $this.LogFileName), $this.TimeOut)
    }

    [void] SetLoggingLevel([LoggingLevel]$LoggingLevel)
    {
        $this.TargetLogLevel = $LoggingLevel
    }

    [string] GetLoggingLevel()
    {
        return ("{0}" -f $this.TargetLogLevel.LoggingLevelName)
    }

    [void] SetLoggingPath([string]$LogFolder)
    {
        #  Get the path to the currently running script.
        $scriptPath = Split-Path -Parent $PSCommandPath

        #  Check if the provided path is fully qualified or relative.
        if ([System.IO.Path]::IsPathRooted($LogFolder))
        {
            #  The provided path is absolute/fully qualified
            $this.LogFilePath = $LogFolder
        }
        else 
        {
            $this.LogFilePath = Join-Path $scriptPath $LogFolder
        }

        #  Set the Log File Object
        $this.LogFileObj.SetLoggingFile((Join-Path $this.LogFilePath $this.LogFileName))
    }

    [string] GetLoggingPath()
    {
        return $this.LogFilePath
    }

    [void] SetLoggingFile([string]$FileName)
    {
        $this.LogFileName = $FileName

        #  Set the Log File Object
        $this.LogFileObj.SetLoggingFile((Join-Path $this.LogFilePath $this.LogFileName))
    }

    [string] GetLoggingFile()
    {
        return $this.LogFileName
    }

    [System.Threading.Mutex] GetLoggingMutex()
    {
        return $this.LogFileObj.GetLoggingMutex()
    }

    [int] GetTimeOutSeconds()
    {
        return $this.TimeOut
    }

    [void] SetTimeOutSeconds(
        [int]$TimeOutSeconds
    )
    {
        $this.TimeOut = $TimeOutSeconds
        $this.LogFileObj.SetTimeOutSeconds($this.TimeOut)
    }
    [string] ToString()
    {
        $_loggerTable = [ordered]@{
            #
            Level = $this.TargetLogLevel.Name()
            FileName = $this.LogFileName
            Path = $this.LogFilePath
            MutexName = $this.LogFileObj.GetLoggingMutexName()
        }

        $_retString = ""

        foreach ($key in $_loggerTable.keys)
        {
            if ($_retString -eq "")
            {
                $_retString = ("`r`n`t{0} | {1}" -f $key.PadLeft(12), $_loggerTable[$key])
            }
            else
            {
                $_retString += ("`r`n`t{0} | {1}" -f $key.PadLeft(12), $_loggerTable[$key])
            }
        }
        return $_retString
    }
    # Method:  Write Output
    [void] Write(
        [string]$Message
    )
    {
        # Write out the message: Console & Log, No level.
        if ($this.TargetLogLevel.Level() -ge ([LoggingLevel]("Information")).Level())
        {
            $this._WriteOutput($Message, "Information", "Both", $False)
        }
        
    }
    
    [void] Write(
        [string]$Message,
        [LoggingLevel]$Level
    )
    {
        # Write out the message: Console & Log, With level .
        if ($this.TargetLogLevel.Level() -ge ([LoggingLevel]($Level).Name()).Level())
        {
            $this._WriteOutput($Message, $Level, "Both", $False)
        }
    }
    
    [void] Write(
        [string]$Message,
        [LoggingLevel]$Level,
        [LoggingDestination]$Destination
    )
    {
        # Write out the message: Console or Log, With level and destination.
        if ($this.TargetLogLevel.Level() -ge ([LoggingLevel]($Level).Name()).Level())
        {
            $this._WriteOutput($Message, $Level.Name(), $Destination, $False)
        }
    }

    [void] Write(
        [string]$Message,
        [LoggingLevel]$Level,
        [LoggingDestination]$Destination,
        [bool]$Bare
    )
    {
        # Write out the message: Console or Log, With level, destination, and no prefix.
        if ($this.TargetLogLevel.Level() -ge ([LoggingLevel]($Level).Name()).Level())
        {
            $this._WriteOutput($Message, $Level.Name(), $Destination, $True)
        }
    }

    hidden [void] _WriteOutput(
        [string]$Message,
        [LoggingLevel]$Level,
        [LoggingDestination]$Destination,
        [bool]$Bare
    )
    {
        Write-Information ("Starting Output")
        # Create formated string for output
        if (-not $Bare)
        {
            # Add date stamp
            $writeString = ("{0} | {1:} | {2}" -f $(Get-Date -Format "yyyy-dd-MM hh:mm:ss"), $Level.Name().PadLeft(11,' '), $Message)
        }
        else
        {
            $writeString = $Message
        }

        # Write the message
        if ($Destination.Console())
        {
            # Write to the console
            $this.LogFileObj.WriteConsole($writeString, $Level)
        }

        if ($Destination.File())
        {
            # Write to the file
            $this.LogFileObj.WriteFile($writeString)
        }
    }
<#
    .SYNOPSIS
    This class will allow easier access to logging to the console and to a file.

    .DESCRIPTION
    You can write to the console only, the log file only or both.  The message can be prefixed with a date stamp
    or it can be written bare.  You can specify the logging level that you want to have outputted without having
    to modify the write statements.

    To use the logging module in your PowerShell scripts use the "Using" command.  To import all classes.
        Example:  Using module ".\Logger.psm1"

        Example:    $Logger = [Logger]::new("Verbose")
                    $Logger.Write("Hello World", "Verbose")
    
        This will create a new logging obect with the logging level set to Verbose.
#>
}
function Write-Test {
    Write-Host "Hello World"
    Write-Host "This is a test of the Logger Module"
}

Export-ModuleMember -Function Write-Test
