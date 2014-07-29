<!--#include file="brandcaptchalib.asp"-->
<%
brandcaptcha_public_key       = "your_public_key" ' your public key
brandcaptcha_private_key      = "your_private_key" ' your private key
%>
<%

brandcaptcha_challenge_field  = Request.Form("brand_cap_challenge")
brandcaptcha_response_field   = Request.Form("brand_cap_answer")
    
if (brandcaptcha_challenge_field <> "" or brandcaptcha_response_field <> "") then
    server_response = brandcaptcha_confirm(brandcaptcha_challenge_field, brandcaptcha_response_field, brandcaptcha_private_key)

    if server_response <> "" then
        'An error occurred 
        Response.Write ("Wrong")
    else
        'The solution was correct
        Response.Write("Correct")
    end if

end if
%>
  
<html>
    <body>
        <!-- Generating the form -->
        <form action="example.asp" method="post">
            <%=brandcaptcha_challenge_writer(brandcaptcha_public_key)%>
        </form>
    </body>
</html>