$body= @{
    subject = "A nice subject"
    importance = "low"
    body = @{
        contentType ="HTML"
        content ="Body <b>message</b>!"
    }
    toRecipients =@(
		@{
            	emailAddress =@{
                		address = "<receivermail>"
            	}
		}
    )
}

 Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/users/<frommail>/messages" -ContentType application/json -Body $body 


