
#todays off powershell script
# this script make a list of today's off members from excel file and post to the Teams
# via imcoming Webhook
# Version 10.6


$todayflag = 0
$nextdayflag = 1
$nameCol = 2　　#constant
$aliasCol = 3   #constant
$dayRaw = 5     #constant

#following valuables are count from excel sheet.
$memberRawStart = 1
$numMembers = 0

try {
    $config = Get-Content -Path .\config.json -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}catch{ 
    Write-Error( "config.json is not found")
    exit
}

if(($config.excelworkbook -eq "") -or ($config.url -eq ""))
{
    Out-String "config.json is not found or not configured"
    exit
}

$debugflag = $config.Debugflag



function isholiday($daycol)
{
    if( $exSheet.Cells.Item(6,$daycol ).Text -eq "土") {return $true}
    if( $exSheet.Cells.Item(6,$daycol ).Text -eq "日") {return $true}
    if( $exSheet.Cells.Item(2,$daycol ).Text -ne "")   {return $true}
    
    return $false
}

function getdaycol( $date )
{
    if( $date.Month -le 6 ) {
        $pastDays = $date.Date - (Get-Date -Year ($date.Year-1) -Month 7 -Day 1).Date
    }
    else {
        $pastDays = $date.Date - (Get-Date -Year ($date.Year) -Month 7 -Day 1).Date  
    }
    $todaysCol = $pastDays.Days + 4
    return $todaysCol
}

function getsheet( $date )
{
    if( $date.Month -le 6 ) {
        $sh = $exSheetFst
    }
    else {
        $sh = $exSheetNxt
    }
    return $sh
}

function isholiday2( $date )
{
    $sheet1 = getsheet( $date )
    

    if( $sheet1.Cells.Item(6, (getdaycol( $date )) ).Text -eq "土") {return $true}
    if( $sheet1.Cells.Item(6, (getdaycol( $date )) ).Text -eq "日") {return $true}
    if( $sheet1.Cells.Item(2, (getdaycol( $date )) ).Text -ne "")   {return $true}
    
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
    if( $term -eq "AM") { return $true}
    if( $term -eq "PM") { return $true}
    if( $term -eq "代休") { return $true}
    if( $term -eq "代")  { return $true}
    if( $term -eq "SD" ) { return $true}
    if( $term -eq "WB" ) { return $true}
    return $false
    
}


function formatState( $term )
{
    
    if($term -eq "T") { return "Ｔ"}
    if($term -eq "AM") { return "㏂"}
    if($term -eq "PM") { return "㏘"}
    if($term -eq "WH") { return "　" }
    if($term -eq "出") { return "出"}
    if($term -eq "代休") { return "代"}
    if($term -eq "代") { return "代"}
    if($term -eq "SD") { return "Ｓ"}
    if($term -eq "WB") { return "㏝"}
    if( isOFF($term) ) { return $term }

    return "　"
  
}

function SGTTeam( $targetdate )
{

    if($targetdate.Date -lt (Get-Date -Date "2020/7/1")) {
        retturn " "  
    }
    $timespan = ($targetdate.Date) - ((Get-Date -Date "2020/6/29").Date)
    $teamid = [Math]::Truncate((($timespan.Days / 7 ) % 3))
    if( $teamid -eq 0 )  {return "SGT: Red Team"}
    if( $teamid -eq 1 )  {return "SGT: Green Team"}
    if( $teamid -eq 2 )  {return "SGT: Blue Team"}
}


if($debugflag)  {
    # for testing purpose
    out-host -InputObject "Debug mode"
    $uri = $config.url
}
else {    
    #for お休みチャンネル
    Out-Host -InputObject "Release Mode"
    $uri = $config.url
}

$today = Get-Date
if($debugflag)  {
    #for debug purpose to override specify a date for today.
    $today = [DateTime]::ParseExact("20200714","yyyyMMdd", $null)
}


$thisyear = $today.Year - 2000
## $currentFY = 20
## $book
## $exSheet



if($today.Month -le 6 ) {
    $CurrentFY = $thisyear
    $past = Get-Date -Year ($today.Year-1) -Month 7 -Day 1
   
}
else {
    $CurrentFY = $thisyear+1
    $past = Get-Date -Year ($today.Year) -Month 7 -Day 1
}

$excelworkbookpath = $config.excelworkbookpath 
$excelnamebase = "Schedule_UC_FY"
$excelworkbook = $excelworkbookpath + $excelnamebase + ($CurrentFY) + ".xlsx"
$excelworksheet = "FY" + $CurrentFY



$endofthisFY = Get-Date -Year ($today.Year) -Month 6 -Day 30
$lastwednesdayinFY = $endofthisFY.AddDays( 1-( $endofthisFY.DayOfWeek.value__ +5))

if ( ($today.Date -lt $lastwednesdayinFY.Date ) -or ( $today.Date -ge ($endofthisFY.AddDays( 3-$endofthisFY.DayOfWeek.value__))) )  {

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $book = $excel.Workbooks.Open( $excelworkbook )
    $exSheet = $book.Worksheets.Item($excelworksheet)

    

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


        

        $pastDays = $today.Date - $past.Date
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
                    $sgtteam = SGTTeam( $today )
                    $outtitle = "Today's OOF (" + $today.ToString("MM/dd") + ") " + $sgtteam +  "`r`n";
                }
                else {
                    $sgtteam = SGTTeam( $nextday )
                    $outtitle = "Next day's OOF (" + $nextday.ToString("MM/dd") + ") " + $sgtteam + "`r`n";
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

                if( $outtext -eq "" ){
                    $outtext = "No one is OOF`r`n"
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

}
else {

    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $book = $excel.Workbooks.Open( $excelworkbook )
    $exSheet = $book.Worksheets.Item($excelworksheet)

    if( $today.Month -le 6 )    {
    
        $exSheetFst = $exSheet
        $bookFst = $book
        $excelworkbookNxt = $excelworkbookpath + $excelnamebase + ($CurrentFY+1) + ".xlsx"
        $excelworksheetNxt = "FY" + ($CurrentFY+1)
        $bookNxt = $excel.Workbooks.Open( $excelworkbookNxt )
        $exSheetNxt = $bookNxt.Worksheets.Item($excelworksheetNxt)
    
    }
    else {
        $exSheetNxt = $exSheet
        $bookNxt = $book
        $excelworkbookFst = $excelworkbookpath + $excelnamebase + ($CurrentFY-1) + ".xlsx"
        $excelworksheetFst = "FY" + ($CurrentFY-1)
        $bookFst = $excel.Workbooks.Open( $excelworkbookFst )
        $exSheetFst = $bookFst.Worksheets.Item($excelworksheetFst)
    }

    try{

        for( $memberRawStart = 1 ; $memberRawStart -le 60 ; $memberRawStart++ )
        {
            $sh_ = getsheet( $today )
            if( $sh_.Cells.Item( $memberRawStart, 1).Text -eq "UC" )
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


        if( !( isholiday2($today) ) ) {

            #Create a today and next day OOF members list and post to teams

            #skip holidays to find next business day.
            for($nextday = $today.AddDays(1) ;isholiday2($nextday) ; $nextday = $nextday.AddDays(1) ){}
            
            #$nextday = $today.AddDays($nextdayCol - $todaysCol)
            
            for( $dayflag = $todayflag; $dayflag -le $nextdayflag; $dayflag++) {

                if($dayflag -eq $todayflag) {
                    $day_ = $today
                }
                else {
                    $day_ = $nextday
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
                    $sh_ = getsheet( $day_ )
                    $state = $sh_.Cells.Item($memberRaw, (getdaycol( $day_ )) ).Text
                    if ( isOOF( $state ) )  { 
                        $ucname =   $sh_.Cells.Item($memberRaw, $nameCol ).Text
                        #align name length to 8 chars to add double byte space chars
                        for( ; $ucname.Length -le 8; $ucname = $ucname+"　"){}
                        $ucname = $ucname + $sh_.Cells.Item($memberRaw, $aliasCol ).Text
                        
                        $outtext = $outtext + $ucname + "    (" + $state + ")  `r`n"
                    
                    }
                }

                if( $outtext -eq "" ){
                    $outtext = "No one is OOF`r`n"
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

            $nextmonday = $today.AddDays( 8-$today.DayOfWeek.value__)
            $lastmonday = $today.AddDays( 1-$today.DayOfWeek.value__)
        
            if( $today.DayOfWeek.value__ -le 2 )
            {
                $monday = $lastmonday
                $weektitle = "今週の予定"  
            }
            else
            {
                $monday = $nextmonday
                $weektitle = "来週の予定"
            }
            #add from-to date in the title.
            $weektitle = $weektitle + "    (" +  $monday.Day + " - " + ($monday.AddDays(4)).Day + ")"
                
            $weektext = "月火水木金　            `r`n"

            for( $memberRaw = $memberRawStart ; $memberRaw -lt ($memberRawStart+$numMembers) ; $memberRaw++)
            {
                $ucname =   $exSheet.Cells.Item($memberRaw, $nameCol ).Text
                #align name length to 8 chars to add double byte space chars
                for( ; $ucname.Length -le 8; $ucname = $ucname+"　"){}

                for( $dow = $monday ; $dow.DayOfWeek.value__ -le 5 ; $dow = $dow.AddDays(1) )
                {
                    $sh_ = getsheet( $dow )
                    $state = $sh_.Cells.Item($memberRaw, (getdaycol( $dow ))).Text
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
        $bookNxt.Close($false)
        $bookFst.Close($false)
        $excel.Quit()
    }
    

}

