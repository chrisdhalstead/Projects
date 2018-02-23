'on error resume next

    Dim lLocator
    Set lLocator = CreateObject("WbemScripting.SWbemLocator")
    Dim gService
    Set gService = lLocator.ConnectServer("TOWSMS03","root\sms\site_TOW")
	Set fsoObject = CreateObject("Scripting.FileSystemObject")
	
	sbuildtxt = ""
	sfilepath = ""
	scollection = ""
	saddedtxt = ""

	scollection = inputbox("Enter the name of the collection to modify","Enter Collection Name")

    bexists = false

	Set oCollectionSet = gService.ExecQuery("Select * From SMS_Collection")
 
     For Each oCollection In oCollectionSet

			If oCollection.Name = scollection Then
              
			   bexists = true
               
            End If
        
	Next

	set ocollectionset = nothing

	if bexists = False then
		
		wscript.echo("There is no Collection Named " & scollection)
		
		wscript.quit

	end if

	sfilepath = inputbox("Enter the Name of the file to read from","Enter File Name",".txt")

	if fsoObject.FileExists(sfilepath) Then


	else

	   wscript.echo("File Not Found!")

	   wscript.quit

	end if
	
	set file = fsoobject.opentextfile(sfilepath, 1)

do until file.AtEndOfStream

		stext = file.readline

		if stext = "" then
		

		else
		

	    Set Machines = gService.ExecQuery("Select * From SMS_R_System WHERE Name LIKE ""%" + stext + "%""")
	
		For Each PC In Machines
            
			If UCase(pc.name) = UCase(stext) Then
			
				 ResID = pc.ResourceID

				saddedtxt = saddedtxt & pc.name & vbcrlf
					
			else
				
				sbuildtxt = sbuildtxt & pc.name & vbcrlf

			end if
			  
        Next
	
      Dim CollectionRule
      Set CollectionRule = gService.Get("SMS_CollectionRuleDirect").SpawnInstance_()
      CollectionRule.ResourceClassName = "SMS_R_System"
      CollectionRule.RuleName = "Direct"
      CollectionRule.ResourceID = resid
      Dim oCollectionSet
      Dim oCollection
        
		Set oCollectionSet = gService.ExecQuery("Select * From SMS_Collection")
           
		For Each oCollection In oCollectionSet
			
			If oCollection.Name = scollection Then
              
				  oCollection.AddMembershipRule CollectionRule
               
            End If
        
		Next


		end if


loop

wscript.echo "Done!"

wscript.echo "Not in SMS: " & sbuildtxt

wscript.echo "Added: " & saddedtxt