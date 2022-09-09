<%
 '
 ' Copyright (c) 2014 by PontaMedia & Zarego
 ' Author: Carlos A. Bellucci, Daniel G. Gomez 
 '
 ' This is a ASP library that handles calling BrandCaptcha.
 '    - Documentation and latest version
 '          http://www.pontamedia.com/
 '

 ' This code is based on code from,
 ' and copied, modified and distributed with permission in accordance with its terms: https://developers.google.com/recaptcha/docs/asp
 
 ' Licensed under the Apache License, Version 2.0 (the "License");
 ' you may not use this file except in compliance with the License.
 ' You may obtain a copy of the License at

 '  http://www.apache.org/licenses/LICENSE-2.0

 ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 ' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 ' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 ' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 ' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 ' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 ' THE SOFTWARE.
 



' returns the HTML for the widget
function brandcaptcha_challenge_writer(public_key)
    brandcaptcha_challenge_writer = _
    "<script type=""text/javascript"" src=""//api.ponta.co/challenge.php?k=" & public_key & """></script>"
end function

' returns "" if correct, otherwise it returns the error response
function brandcaptcha_confirm(rechallenge,reresponse, private_key)

    Dim VarString
    VarString = _
            "privatekey=" & private_key & _
            "&remoteip=" & Request.ServerVariables("REMOTE_ADDR") & _
            "&challenge=" & rechallenge & _
            "&response=" & reresponse

    Dim objXmlHttp
    Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
    objXmlHttp.open "POST", "http://api.ponta.co/verify.php", False
    objXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXmlHttp.send VarString

    Dim ResponseString
    ResponseString = split(objXmlHttp.responseText, vblf)
    Set objXmlHttp = Nothing
  
  
    if ResponseString(0) = "true" then
    'They answered correctly
        brandcaptcha_confirm = ""
    else
    'They answered incorrectly
        brandcaptcha_confirm = ResponseString(1)
    end if

end function
%>
