# from https://gallery.technet.microsoft.com/Send-Tweets-via-a-72b97964
 
workflow Send-Tweet {
    param (
    [Parameter(Mandatory=$true)][string]$Message
    )

    InlineScript {      
        [Reflection.Assembly]::LoadWithPartialName("System.Security")  
        [Reflection.Assembly]::LoadWithPartialName("System.Net")  
        
        $status = [System.Uri]::EscapeDataString($Using:Message);  
        $oauth_consumer_key = "xyz";  
        $oauth_consumer_secret = "xyz";  
        $oauth_token = "xyz";  
        $oauth_token_secret = "xyz";  
        $oauth_nonce = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes([System.DateTime]::Now.Ticks.ToString()));  
        $ts = [System.DateTime]::UtcNow - [System.DateTime]::ParseExact("01/01/1970", "dd/MM/yyyy", $null).ToUniversalTime();  
        $oauth_timestamp = [System.Convert]::ToInt64($ts.TotalSeconds).ToString();  
  
        $signature = "POST&";  
        $signature += [System.Uri]::EscapeDataString("https://api.twitter.com/1.1/statuses/update.json") + "&";  
        $signature += [System.Uri]::EscapeDataString("oauth_consumer_key=" + $oauth_consumer_key + "&");  
        $signature += [System.Uri]::EscapeDataString("oauth_nonce=" + $oauth_nonce + "&");   
        $signature += [System.Uri]::EscapeDataString("oauth_signature_method=HMAC-SHA1&");  
        $signature += [System.Uri]::EscapeDataString("oauth_timestamp=" + $oauth_timestamp + "&");  
        $signature += [System.Uri]::EscapeDataString("oauth_token=" + $oauth_token + "&");  
        $signature += [System.Uri]::EscapeDataString("oauth_version=1.0&");  
        $signature += [System.Uri]::EscapeDataString("status=" + $status);  
  
        $signature_key = [System.Uri]::EscapeDataString($oauth_consumer_secret) + "&" + [System.Uri]::EscapeDataString($oauth_token_secret);  
  
        $hmacsha1 = new-object System.Security.Cryptography.HMACSHA1;  
        $hmacsha1.Key = [System.Text.Encoding]::ASCII.GetBytes($signature_key);  
        $oauth_signature = [System.Convert]::ToBase64String($hmacsha1.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($signature)));  
  
        $oauth_authorization = 'OAuth ';  
        $oauth_authorization += 'oauth_consumer_key="' + [System.Uri]::EscapeDataString($oauth_consumer_key) + '",';  
        $oauth_authorization += 'oauth_nonce="' + [System.Uri]::EscapeDataString($oauth_nonce) + '",';  
        $oauth_authorization += 'oauth_signature="' + [System.Uri]::EscapeDataString($oauth_signature) + '",';  
        $oauth_authorization += 'oauth_signature_method="HMAC-SHA1",'  
        $oauth_authorization += 'oauth_timestamp="' + [System.Uri]::EscapeDataString($oauth_timestamp) + '",'  
        $oauth_authorization += 'oauth_token="' + [System.Uri]::EscapeDataString($oauth_token) + '",';  
        $oauth_authorization += 'oauth_version="1.0"';  
    
        $post_body = [System.Text.Encoding]::ASCII.GetBytes("status=" + $status);   
        [System.Net.HttpWebRequest] $request = [System.Net.WebRequest]::Create("https://api.twitter.com/1.1/statuses/update.json");  
        $request.Method = "POST";  
        $request.Headers.Add("Authorization", $oauth_authorization);  
        $request.ContentType = "application/x-www-form-urlencoded";  
        $body = $request.GetRequestStream();  
        $body.Write($post_body, 0, $post_body.length);  
        $body.flush();  
        $body.close();  
        $response = $request.GetResponse();
    }
 }

Start-Transcript -Path C:\Scripts\Scheduled\ClosedAsFixed\transcript.txt -Append
Add-Type -Path C:\Scripts\Scheduled\ClosedAsFixed\HtmlAgilityPack.dll
Add-Type -AssemblyName System.Web
$doc = New-Object HtmlAgilityPack.HtmlDocument 

$items = @()
$archive = Import-Csv C:\Scripts\Scheduled\ClosedAsFixed\archive.csv


2..3 | ForEach-Object {
	$statusnum = $_
	$url = "https://connect.microsoft.com/SQLServer/SearchResults.aspx?FeedbackType=0&Status=$statusnum&Scope=0&ChangedDays=2&SortOrder=40&TabView=0"
	
	#5,15,10,25,30,20,35
	$html = Invoke-WebRequest -UseBasicParsing -Uri $url

	$null = $doc.LoadHtml($html.Content) 

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
			
			switch ($author) {
				"Chrissy LeMaire" { $author = "@cl" }
				"SQLvariant" { $author = "@SQLvariant" }
				"Aaron Nelson" { $author = "@SQLvariant" }
				"AMtwo" { $author = "@AMtwo" }
			}
			
			# Created and a bunch of other metadata
			$xpath = "$($basepath)_FeedbackItemDetailsUpdatePanel']/div"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			
			# created
			$xpath = "$($basepath)_FeedbackItemDetailsUpdatePanel']/div/span[1]"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$created = (($node.FirstChild.InnerText -Split "&nbsp;")[0]).TrimStart("Created on ")
			
			
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
			
			$exists = $archive | Where-Object { $_.id -eq $feedbackid }
			
			if ($exists.count -eq 0) {
				# add item to collection
				$items += [PSCustomObject]@{
					Link = $link
					Title = $title
					ID = $feedbackid
					Author = $author
					Created = $created
					Closed = Get-Date -f "M/dd/yyyy"
					Votes = $votes
					ItemType = "Bug"
					Tweet = $tweet
				}
			}
		}
	}
}

2..3 | ForEach-Object {
	$statusnum = $_
	$url = "https://connect.microsoft.com/SQLServer/SearchResults.aspx?FeedbackType=0&Status=$statusnum&Scope=0&ChangedDays=2&SortOrder=40&TabView=1"
	
	$html = Invoke-WebRequest -UseBasicParsing -Uri $url
	$null = $doc.LoadHtml($html.Content) 

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
				"Chrissy LeMaire" { $author = "@cl" }
				"SQLvariant" { $author = "@SQLvariant" }
				"Aaron Nelson" { $author = "@SQLvariant" }
				"AMtwo" { $author = "@AMtwo" }
			}
			
			# Created and a bunch of other metadata
			$xpath = "$($basepath)_FeedbackItemDetailsUpdatePanel']/div"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			
			# created
			$xpath = "$($basepath)_FeedbackItemDetailsUpdatePanel']/div/span[1]"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$created = (($node.FirstChild.InnerText -Split "&nbsp;")[0]).TrimStart("Created on ")
		
			# Votes
			$xpath = "$($basepath)_FeedbackItemVotingControlForVoting_SimpleVote_ctl02_text']"
			$node = $doc.DocumentNode.SelectNodes($xpath)
			$votes = $node.InnerText
			
			$exists = $archive | Where-Object { $_.id -eq $feedbackid }
			
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
					Closed = Get-Date -f "M/dd/yyyy"
					Votes = $votes
					ItemType = "Suggestion"
					Tweet = $tweet
				}
			}
		}
	}
}

Write-Output "Number of Tweetable Items: $($items.count)"

$dedupe = $items | Where-Object { $_.Link -notin $archive.Link }
$post = $dedupe | Select -First 1

$post.Tweet

if ($post -ne $null) {
	try {
		$null = Send-Tweet $post.Tweet
		$post | Select-Object Link, Title, ID, Author, Created, Closed, Votes, ItemType | Export-Csv -Path C:\Scripts\Scheduled\ClosedAsFixed\archive.csv -Append
	} catch {
		Add-Content -Path C:\Scripts\Scheduled\ClosedAsFixed\errors.csv "$(Get-Date) $post"
	}
}

Stop-Transcript