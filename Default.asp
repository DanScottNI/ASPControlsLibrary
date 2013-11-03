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
                    <td><% lblLiteral.Render() %>
                        <% txtEdit.Render() %>
                    </td>
                </tr>
            </table>
            <%
                btnSend.Render()
            %>
            <%
                btnTest.Render()
            %>
        </form>
    </body>
</html>
