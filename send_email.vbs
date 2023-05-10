
strUsername = "cctlkolkata@gmail.com"
strPassword = "JgbVhBytTrUOsq3x"
strSchema = "http://schemas.microsoft.com/cdo/configuration/"
Set objEmail = CreateObject("CDO.Message")
objEmail.From = strUsername
objEmail.To = "sprusty@stqc.gov.in; sanjayprusty@gmail.com"
objEmail.Subject = "email test as on" & vbcrlf & now()
objEmail.Textbody = "Hi, This is from Ani......."
'Fire the mail
objEmail.Configuration.Fields.Item (strSchema & "sendusing") = 2
objEmail.Configuration.Fields.Item (strSchema & "smtpserver") = "smtp-relay.sendinblue.com"
objEmail.Configuration.Fields.Item (strSchema & "smtpserverport") = 465
objEmail.Configuration.Fields.Item (strSchema & "smtpusessl") = 1
objEmail.Configuration.Fields.Item (strSchema & "smtpauthenticate") = 1
objEmail.Configuration.Fields.Item (strSchema & "sendusername") = strUsername
objEmail.Configuration.Fields.Item (strSchema & "sendpassword") = strPassword
objEmail.Configuration.Fields.Update
objEmail.Send