'*********************************************************************
' WARNING: ANY USE BY YOU IS AT YOUR OWN RISK. I provide this code
' "as is" without warranty of any kind, either express or implied,
' including but not limited to the implied warranties of
' merchantability and/or fitness for a particular purpose.
'*********************************************************************

A friend of mine asked me about how to let excel periodically update
some worksheets containing stock prices. So the first thing I did was
searching the web for a solution in Visual Basic for Applications.
But the only one I found was to use a 3rd party ActiveX timer control
which you have to place onto a userform. Then I found the VBA command
"Application.onTime" and I had an idea. 

Here is my solution only utilizing the VBA onTime command. No need for
ActiveX control or any other 3rd party module.

If you find this piece of code useful, please vote for it at Planet-Source-Code.com:
http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=34409&lngWId=1


Kind regards,
Sebastian Thomschke										  			05/03/2002
http://www.sebthom.de