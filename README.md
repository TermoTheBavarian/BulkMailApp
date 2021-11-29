# BulkMailApp
Bulk Mail App developed for ECOSEN in Python with a simple GUI.  The App scratchs an Excel Agenda of Company names and e-mail adress, adding the company name to the mail subject. 
It allows to add a PDF File (usually references), a personalized subject and a Word Text that will be converted to HTML as the mail message text. 
Then, it will display a draft with the first adress for some seconds; after closing this draft it will display the list of recipients and adress will be send; 
so the user can confirm. Finally, an individual mail will be send to each recipient each 1 minute in order avoid servers from blocking our campaign.  
Once the Campaign is finished a message will be dispalyed.
