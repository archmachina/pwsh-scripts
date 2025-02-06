param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNull()]
    [int]$Duration = 60,

    [Parameter(Mandatory=$false)]
    [ValidateNotNull()]
    [int]$IterationSize = 10000
)

# Global Settings
Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

$iteration = 0
$start = [DateTime]::Now
$window = $start
$lastIteration = 0

while (([DateTime]::Now) -lt $start.AddSeconds($Duration))
{
    $iteration++

    # Print current status
    if ([DateTime]::Now -gt $window.AddSeconds(5))
    {
        # Start the new window from this point
        $oldWindow = $window
        $window = [DateTime]::Now

        # Local iterations is the number of iterations since the last output
        $localIterations = $iteration - $lastIteration
        $lastIteration = $iteration

        # Calculate averages
        $runningAverage = [Math]::Round($iteration / ($window - $start).TotalSeconds, 2)
        $currentAverage = [Math]::Round($localIterations / ($window - $oldWindow).TotalSeconds, 2)

        # Write status information
        Write-Information ("Iterations {0} - average: {1} p/sec since start - {2} p/sec since last threshold" -f $iteration, $runningAverage, $currentAverage)
    }

    # Perform sqrt calculation
    1..$IterationSize | ForEach-Object {
        $sum = [Math]::Sqrt($_)
    }
}

Write-Information "Finished: $iteration"
Write-Information ("Elapsed time: {0} seconds" -f [Math]::Round(([DateTime]::Now - $start).TotalSeconds, 2))

