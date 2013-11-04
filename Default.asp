<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="Default.CodeBehind.asp"-->
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
    </head>
    <body>
        <form method="post" id="frmMain" name="frmMain">
            <table width="100%">
                <tr>
                    <td>
                        <% lblLiteral %>
                        <% txtEdit %>
                    </td>
                </tr>
            </table>
            <%
                btnSend
            %>
            <%
                btnTest
            %>
        </form>
    </body>
</html>
