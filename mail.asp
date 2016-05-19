<%
Set Mail = CreateObject("CDO.Message")

'This section provides the configuration information for the remote SMTP server.

'Send the message using the network (SMTP over the network).
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="mail.aclaimrx.com"
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

'Use SSL for the connection (True or False)
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

'If your server requires outgoing authentication, uncomment the lines below and use a valid email address and password.
'Basic (clear-text) authentication
'Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 
'Your UserID on the SMTP server
'Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="ACL/RXUPDATES\rxupdates"
'Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="K2qhZXxz"

Mail.Configuration.Fields.Update

'End of remote SMTP server configuration section

Mail.Subject="Contact Request via aClaimRx.com"
Mail.From="rxupdates@aclaimrx.com"
Mail.To="rlee@aclaimrx.com"
Mail.TextBody= "Contact request from " & request.form("name") & ", E-Mail Address: " & request.form("email") & ", Telephone: " & request.form("phone") & ", Message: " & request.form("message") 
Mail.Send
Set Mail = Nothing
%>