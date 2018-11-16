# Clear-OrphanedRoomReservation
Removes meeting from conference rooms when the organizer account is deleted or disabled

## SYNOPSIS
Removes meeting from conference rooms where the organizer is a deleted mailbox.
		
## DESCRIPTION
The script will lookup the list of disconnected mailboxes on Exchange, then build a query for search-mailbox to find and remove all meetings from organizers matching the disconnected mailbox name.
        
## PARAMETERS
### SearchDays
Filter the input list to mailboxes that were disconnected in the last <SearchDays> 

## EXAMPLE
.\Clear-OrphanedRoomReservation -SearchDays 2 -Verbose
