############################################################### 
# Eric Ligmans Amazing Free Microsoft eBook Giveaway 
# https://blogs.msdn.microsoft.com/mssmallbiz/2017/07/11/largest-free-microsoft-ebook-giveaway-im-giving-away-millions-of-free-microsoft-ebooks-again-including-windows-10-office-365-office-2016-power-bi-azure-windows-8-1-office-2013-sharepo/
# Link to download list of eBooks 
# http://ligman.me/2sZVmcG 
# Thanks David Crosby for the template (https://social.technet.microsoft.com/profile/david%20crosby/)
#
###############################################################
#
# A little update and optimisation 14/07/2016
#
# Script now:
#			Checks if folder exists and creates directory if it doesnt already exist
#			Allows user to spesify file types they wish to download
# 			Uses BitsTransfer to download files a little faster than before (credit to Jourdan Templeton - https://blog.jourdant.me/post/3-ways-to-download-files-with-powershell)
# 			Handels errors & logs dead links 
#			
############################################################### 

# Create directory if it doesnt already exist 
$dest = "C:\Downloads\ebooks\" 
New-Item -ItemType Directory -Force -Path $dest 

Import-Module BitsTransfer

# Uncomment the filetypes required
$filetype = "PDF"
#$filetype += "EPUB"
#$filetype += "MOBI"

# Download the source list of books 
$downLoadList = "http://ligman.me/2sZVmcG" 
$bookList = Invoke-WebRequest $downLoadList 

# Convert the list to an array 
[string[]]$books = "" 
$books = $bookList.Content.Split("`n") 
# Remove the first line - it's not a book 
$books = $books[1..($books.Length -1)] 
#$books # Here's the list 
 
# Get the title and filetype infomation
foreach ($book in $books) { 
    try {$hdr = Invoke-WebRequest $book -Method Head }catch {$book | Out-File $dest"dead_links.log" -Append}
    $title = $hdr.BaseResponse.ResponseUri.Segments[-1] 

		# Download the books that match filetype 
		foreach ($dtd in $dtds| Where-Object{$filetype -match $title.Split('.')[1] }) {
		    $title = [uri]::UnescapeDataString($title) 
		    $saveTo = Join-Path $dest $title
			Start-BitsTransfer -Source $book -Destination $saveTo
		}
} 