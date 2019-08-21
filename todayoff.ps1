
#todays off powershell script
# this script make a list of today's off members from excel file and post to the Teams
# via imcoming Webhook
# Version 10.1


$todayflag = 0
$nextdayflag = 1
$nameCol = 2　　#constant
$aliasCol = 3   #constant
$dayRaw = 5     #constant

#following valuables are count from excel sheet.
$memberRawStart = 1
$numMembers = 0

# debug flag
if($Env:DEBUG_TODAYSOFF -eq 1)  {
    $debugflag = $true
}
else {
    $debugflag = $true
}


function isholiday($daycol)
{
    if( $exSheet.Cells.Item(6,$daycol ).Text -eq "土") {return $true}
    if( $exSheet.Cells.Item(6,$daycol ).Text -eq "日") {return $true}
    if( $exSheet.Cells.Item(2,$daycol ).Text -ne "")   {return $true}
    
    return $false
}

function isOFF($term)
{
    if( $term -eq "特" ) {return $true}
    if( $term -eq "夏" ) {return $true}
    if( $term -eq "休" ) {return $true}

    return $false
}

function isOOF($term)
{
    if( $term -eq "特" ) {return $true}
    if( $term -eq "夏" ) {return $true}
    if( $term -eq "休" ) {return $true}
    if( $term -eq "出") { return $true}
    if( $term -eq "T")  { return $true}
    return $false
    
}


function formatState( $term )
{
    
    if($term -eq "T") { return "Ｔ"}
    if($term -eq "AM") { return "㏂"}
    if($term -eq "PM") { return "㏘"}
    if($term -eq "WH") { return "　" }
    if($term -eq "出") { return "出"}
    if( isOFF($term) ) { return $term }

    return "　"
  
}

if($debugflag)  {
    # for testing purpose
    out-host -InputObject "Debug mode"
    $uri = "https://outlook.office.com/webhook/5b59c0eb-7c38-4749-b866-7665ce4ddb68@72f988bf-86f1-41af-91ab-2d7cd011db47/IncomingWebhook/88c4784a9992461c8feb28d8de967cf5/f27e8b14-e89d-4a23-ba91-6f6556bb4d79"
}
else {    
    #for お休みチャンネル
    Out-Host -InputObject "Release Mode"
    $uri = "https://outlook.office.com/webhook/ed8b025d-1eb0-4940-8ae9-960a6669515f@72f988bf-86f1-41af-91ab-2d7cd011db47/IncomingWebhook/ac4ef01f403f468d973f1b8b85f3786e/f27e8b14-e89d-4a23-ba91-6f6556bb4d79"
}




$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$book = $excel.Workbooks.Open("https://microsoft.sharepoint.com/teams/SfBDailyCaseAssign/Shared%20Documents/%E3%81%8A%E4%BC%91%E3%81%BF%20Request%20%E3%83%81%E3%83%A3%E3%83%8D%E3%83%AB%20-%20%E9%81%8B%E7%94%A8%E5%89%8D/Schedule_UC_FY20.xlsx")
$exSheet = $book.Worksheets.Item("FY20")

try{

    for( $memberRawStart = 1 ; $memberRawStart -le 60 ; $memberRawStart++ )
    {
        if( $exSheet.Cells.Item( $memberRawStart, 1).Text -eq "UC" )
        {
            $memberRawStart++;
            break;
        }
    }
    if( $memberRawStart -gt 60 ) { Throw } 
    
    for( $numMembers = 0 ; $numMembers -le (60-$memberRawStart) ; $numMembers++)
    {
        if( $exSheet.Cells.Item( $memberRawStart+$numMembers, 2).Text -eq "" )
        {
            break;
        }
    }
    if( $numMembers -gt (60-$memberRawStart)) { Throw }


    $past = [DateTime]::ParseExact("20190701","yyyyMMdd", $null)

    $today = Get-Date
    if($debugflag)  {
        #for debug purpose to override specify a date for today.
        $today = [DateTime]::ParseExact("20190819","yyyyMMdd", $null)
    }

    $pastDays = $today - $past
    $todaysCol = $pastDays.Days + 4

    if( !( isholiday($todaysCol) ) ) {

        #Create a today and next day OOF members list and post to teams

        #skip holidays to find next business day.
        for($nextdayCol = $todaysCol+1;isholiday($nextdayCol) ; $nextdayCol++ ){}
        
        $nextday = $today.AddDays($nextdayCol - $todaysCol)
        
        for( $dayflag = $todayflag; $dayflag -le $nextdayflag; $dayflag++) {

            if($dayflag -eq $todayflag) {
                $dayscol = $todaysCol
            }
            else {
                $dayscol = $nextdayCol
            }

            if($dayflag -eq $todayflag) {
                $outtitle = "Today's OOF (" + $today.ToString("MM/dd") + ")  `r`n";
            }
            else {
                $outtitle = "Next day's OOF (" + $nextday.ToString("MM/dd") + ")  `r`n";
            }

            $outtext = ""

            for( $memberRaw = $memberRawStart ; $memberRaw -lt ($memberRawStart+$numMembers) ; $memberRaw++)
            {
                $state = $exSheet.Cells.Item($memberRaw,$dayscol ).Text
                if ( isOOF( $state ) )  { 
                    $ucname =   $exSheet.Cells.Item($memberRaw, $nameCol ).Text
                    #align name length to 8 chars to add double byte space chars
                    for( ; $ucname.Length -le 8; $ucname = $ucname+"　"){}
                    $ucname = $ucname + $exSheet.Cells.Item($memberRaw, $aliasCol ).Text
                    
                    $outtext = $outtext + $ucname + "    (" + $state + ")  `r`n"
                
                }
            }

            Out-String -inputobject $outtext

            $body = ConvertTo-Json @{
                title = $outtitle
                text = $outtext

            }
            $body = [Text.Encoding]::UTF8.GetBytes($body)
            Invoke-RestMethod -uri $uri -Method Post -Body $body -ContentType 'application/json'
            
        }
        
        # Create a table of this/next week schedule
        # Mon. and Tue. - make this week.
        # Wed.,Thu and Fri - make next week

        if( $today.DayOfWeek.value__ -le 2 )
        {
            $mondayCol = $todaysCol - ($today.DayOfWeek.value__ -1)
            $weektitle = "今週の予定"  
        }
        else
        {
            $mondayCol = $todaysCol+7 - ($today.DayOfWeek.value__ -1)
            $weektitle = "来週の予定"
        }
        #add from-to date in the title.
        $weektitle = $weektitle + "    (" +  $exSheet.Cells.Item($dayRaw, $mondayCol ).Text + " - " + $exSheet.Cells.Item($dayRaw, $mondayCol+4 ).Text + ")"
            
        $weektext = "月火水木金　            `r`n"

        for( $memberRaw = $memberRawStart ; $memberRaw -lt ($memberRawStart+$numMembers) ; $memberRaw++)
        {
            $ucname =   $exSheet.Cells.Item($memberRaw, $nameCol ).Text
            #align name length to 8 chars to add double byte space chars
            for( ; $ucname.Length -le 8; $ucname = $ucname+"　"){}

            for( $dow = $mondayCol ; $dow -le $mondayCol + 4 ; $dow++ )
            {
            $state = $exSheet.Cells.Item($memberRaw, $dow ).Text
            $state = formatState($state)
            $weektext = $weektext + $state
            }
            $weektext = $weektext + "　　" +$ucname+ "    `r`n"
        }

        Out-String -inputobject $weektext


        $body = ConvertTo-Json @{
            title = $weektitle
            text = $weektext
        }
        $body = [Text.Encoding]::UTF8.GetBytes($body)
        Invoke-RestMethod -uri $uri -Method Post -Body $body -ContentType 'application/json'
            
        
    }
}finally{
    $book.Close($false)
    $excel.Quit()
}
