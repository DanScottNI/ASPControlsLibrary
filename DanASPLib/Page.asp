<%
    Function IsPostBack()
      IsPostBack = Request.ServerVariables("REQUEST_METHOD")= "POST"
    End Function
%>