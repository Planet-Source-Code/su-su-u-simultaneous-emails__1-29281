Hello - 

First I would like to give credit where credit is due...
Gregg Housh - did the MX Lookup Control
Pik Soft Inc - did the routine that sends the email

And Second - forgive me if you think this demonstration is pathetic...
This is really intended for beginners, and I know there are some out there
asking how to do this.

If you use any of this code, please give them credit.

This is just a simple program that shows how to send an email
without having an email server at your disposal.

To use, simply enter the Mail From, From Name, Mail To, and Subject, and Message.

How does this work?
Simple - it first takes the domain from the email you are sending to...
So if you are sending to a@b.com the domain is b.com

It then does an MX Lookup to find the mail server for that domain (mail.b.com)
Then it sends the email to the server.

Bryan Cairns