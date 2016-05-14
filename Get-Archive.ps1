# Add required types
Add-Type -Path .\HtmlAgilityPack.dll
Add-Type -AssemblyName System.Web

Function Get-Total ($closedorresolved, $bugorsuggestion){

	switch ($closedorresolved) {
		"Closed" { $status = '3' }
		"Resolved" { $status = '2' }
	}

	switch ($bugorsuggestion) {
		"Bug" { $tabbiew = '0' }
		"Suggestion" { $tabbiew = '1' }
	}

	$url = "https://connect.microsoft.com/SQLServer/SearchResults.aspx?FeedbackType=0&Status=$status&Scope=0&SortOrder=40&TabView=$tabbiew"
	$html = Invoke-WebRequest -UseBasicParsing -Uri $url
	$null = $doc.LoadHtml($html.Content) 
	
	if ($bugorsuggestion -eq "Bug") {
		$total = ($doc.DocumentNode.SelectNodes('//*[@id="ctl00_MasterBody_BugPagingControlTop_Summary"]').InnerText -Split "`n")[4].Trim()
	} else {
		$total = ($doc.DocumentNode.SelectNodes('//*[@id="ctl00_MasterBody_SuggestionPagingControlTop_Summary"]').InnerText -Split "`n")[4].Trim()
	}
	
	return $total

}

Function Parse-ConnectPages ($closedorresolved, $bugorsuggestion) {

	switch ($closedorresolved) {
		"Closed" { $status = '3' }
		"Resolved" { $status = '2' }
	}

	switch ($bugorsuggestion) {
		"Bug" { $tabbiew = '0' }
		"Suggestion" { $tabbiew = '1' }
	}
	
	$total = Get-Total $closedorresolved $bugorsuggestion
	$totalTimes = ($total / 10)
	
	$url = "https://connect.microsoft.com/SQLServer/SearchResults.aspx?FeedbackType=0&Status=$status&Scope=0&SortOrder=40&TabView=$tabbiew#&&PageIndex="
	
	Write-Output "Doing this a total of $totaltimes times for $total items in $closedorresolved $bugorsuggestion"
	1..$totalTimes | ForEach-Object {
		$currentpage = $_
		$currenturl = "$url" + "$currentpage"
		
		$ie.navigate($currenturl)
		While ($ie.Busy) { Start-Sleep 1 }
		
		Start-Sleep 3
		$body = $ie.Document.documentElement.outerHTML
		$null = $doc.LoadHtml($body) 
		
		switch ($bugorsuggestion) {
			"Bug" { Parse-BugPage $doc $closedorresolved  }
			"Suggestion" { Parse-SuggestionPage $doc $closedorresolved  }
		}
	}
}

Function Parse-BugPage ($doc, $status){
	$items =@()
	$count = $doc.DocumentNode.SelectNodes('//*[@id="ctl00_MasterBody_BugPagingControlTop_NumberOfItems"]').InnerText
	Write-Output "Number of items returned: $count"
			
	1..$count | ForEach-Object {

		if ($_ -lt 10) { $num = "0$_" } else { $num = "$_" }
		
		$basepath = "//*[@id='ctl00_MasterBody_BugsSearchResultsView_ctl$($num)_BugsResultModule"
		
		$xpath = "$($basepath)_FeedbackStatus_ResolutionLabel']"
		$node = $doc.DocumentNode.SelectNodes($xpath)
		$resolution = $node.InnerText

		if ($resolution -eq 'as Fixed') {
			
			# feedbackid
			$xpath = "$($basepath)_FeedbackItemDetailsUpdatePanel']/div/span[5]"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$feedbackid = $node.InnerText.TrimStart("feedback id: ")
			$link = "http://connect.microsoft.com/SQLServer/feedback/details/$feedbackid/"
			
			# Title 
			$xpath = "$($basepath)_FeedbackLink']"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$title = $node.InnerText
			$title = [System.Web.HttpUtility]::HtmlDecode($title)
			
			# Author
			$xpath = "$($basepath)_NewFeedbackAuthor']"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$author = $node.InnerText
			
			# Created and a bunch of other metadata
			$xpath = "$($basepath)_FeedbackItemDetailsUpdatePanel']/div"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			
			# created
			$xpath = "$($basepath)_FeedbackItemDetailsUpdatePanel']/div/span[1]"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$created = (($node.FirstChild.InnerText -Split "&nbsp;")[0]).TrimStart("Created on ")
			
			# closed
			try {
				$closed = $node.FirstChild.InnerText
				$closed = $closed -Split '\;\('
				$closed = $closed[1]
				$closed = $closed.Replace("ago)","")
				$closed = $closed.Replace("updated ","")
									
				if ($closed -match "weeks") {
					$daysago = ($closed.Replace(" weeks",""))
					$daysago = ($daysago.Replace(" week",""))
					$daysago = [int]$daysago * 7
				} elseif ($closed -match "days")  {
					$daysago = $closed.Replace(" days","")
					$daysago = $daysago.Replace(" day","")
					$daysago = [int]$daysago
				} else {
					$daysago = 0
				}
				
				$closed = $(Get-Date).AddDays(-$daysago).ToString('MM/dd/yyyy')
				
			} catch { $closed = $node.FirstChild.InnerText }
			
			# Votes
			$xpath = "$($basepath)_FeedbackItemVotingControlForVoting_SimpleVote_ctl02_text']"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$votes = $node.InnerText			
						
			# Create tweet
			$tweet = "$title - $author"
			if ($tweet.length -gt 119) {
				$tweet = $tweet.Substring(0,113) + "..."
			}
			
			$tweet = "$tweet $link"
			$tweet = $tweet.Replace("`&`#39;","")
			
			$exists = $items | Where-Object { $_.id -eq $feedbackid }
			
			if ($exists.count -eq 0) {
				# add item to collection
				$items += [PSCustomObject]@{
					Link = $link
					Title = $title
					ID = $feedbackid
					Author = $author
					Created = $created
					Closed = $closed
					Votes = $votes
					ItemType = "Bug"
					Status = $status
				} 
			}
		}
	}
	$items | Select-Object Link, Title, ID, Author, Created, Closed, Votes, ItemType, Status | Export-Csv -Path .\bug-$status.csv -Append -NoTypeInformation
}

Function Parse-SuggestionPage ($doc, $status) {
	$items = @()
	$count = $doc.DocumentNode.SelectNodes('//*[@id="ctl00_MasterBody_SuggestionPagingControlTop_NumberOfItems"]').InnerText
	Write-Output "Number of items returned: $count"

		1..$count | ForEach-Object {
		if ($_ -lt 10) { $num = "0$_" } else { $num = "$_" }
		
		$basepath = "//*[@id='ctl00_MasterBody_SuggestionsSearchResultsView_ctl$($num)_SuggestionsResultModule"
		
		$xpath = "$($basepath)_FeedbackStatus_ResolutionLabel']"
		$node = $doc.DocumentNode.SelectNodes($xpath)
		$resolution = $node.InnerText

		if ($resolution -eq 'as Fixed') {
		
			# feedbackid
			$xpath = "$($basepath)_FeedbackItemDetailsUpdatePanel']/div/span[4]"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$feedbackid = $node.InnerText.TrimStart("feedback id: ")
			$link = "http://connect.microsoft.com/SQLServer/feedback/details/$feedbackid/" #$link = $node.Attributes[1].Value
			
			# Title 
			$xpath = "$($basepath)_FeedbackLink']"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$title = $node.InnerText
			$title = [System.Web.HttpUtility]::HtmlDecode($title)
			
			# Author
			$xpath = "$($basepath)_NewFeedbackAuthor']"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$author = $node.InnerText
			
			switch ($author) {
				"AaronBertrand" { $author = "@AaronBertrand" }
				"Chrissy LeMaire" { $author = "@cl" }
				"SQLvariant" { $author = "@SQLvariant" }
			}
			
			# Created and a bunch of other metadata
			$xpath = "$($basepath)_FeedbackItemDetailsUpdatePanel']/div"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			
			# created
			$xpath = "$($basepath)_FeedbackItemDetailsUpdatePanel']/div/span[1]"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$created = (($node.FirstChild.InnerText -Split "&nbsp;")[0]).TrimStart("Created on ")
			
			# closed
			try {
				$closed = $node.FirstChild.InnerText
				$closed = $closed -Split '\;\('
				$closed = $closed[1]
				$closed = $closed.Replace("ago)","")
				$closed = $closed.Replace("updated ","")
									
				if ($closed -match "weeks") {
					$daysago = ($closed.Replace(" weeks",""))
					$daysago = ($daysago.Replace(" week",""))
					$daysago = [int]$daysago * 7
				} elseif ($closed -match "days")  {
					$daysago = $closed.Replace(" days","")
					$daysago = $daysago.Replace(" day","")
					$daysago = [int]$daysago
				} else {
					$daysago = 0
				}
				
				$closed = $(Get-Date).AddDays(-$daysago).ToString('MM/dd/yyyy')
				
			} catch { $closed = $node.FirstChild.InnerText }
		
			# Votes
			$xpath = "$($basepath)_FeedbackItemVotingControlForVoting_SimpleVote_ctl02_text']"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$votes = $node.InnerText
			
			$exists = $items | Where-Object { $_.id -eq $feedbackid }
			
			if ($exists.count -eq 0) {
			
				# Create tweet
				$tweet = "$title - $author"
				if ($tweet.length -gt 119) {
					$tweet = $tweet.Substring(0,113) + "..."
				}
				
				$tweet = "$tweet $link"
				$tweet = $tweet.Replace("`&`#39;","")
				
				# add item to collection
				$items += [PSCustomObject]@{
					Link = $link
					Title = $title
					ID = $feedbackid
					Author = $author
					Created = $created
					Closed = $closed
					Votes = $votes
					ItemType = "Suggestion"
					Status = $status
				}
			}
		}
	}
	
	$items | Select-Object Link, Title, ID, Author, Created, Closed, Votes, ItemType, Status | Export-Csv -Path .\suggestion-$status.csv -Append -NoTypeInformation
}

# Start Parsers
$doc = New-Object HtmlAgilityPack.HtmlDocument

# IE has to be used instead of Invoke-WebRequest because of some refresh/click requirements
$ie = New-Object -Com "InternetExplorer.Application"
$ie.visible = $true

# Prep CSV
Remove-Item *.csv -Force -ErrorAction Ignore

# Do it
Parse-ConnectPages "Resolved" "Suggestion"
Parse-ConnectPages "Resolved" "Bug"
Parse-ConnectPages "Closed" "Suggestion"
Parse-ConnectPages "Closed" "Bug"

# Clean up
Get-Process iexplore | Stop-Process -Force
