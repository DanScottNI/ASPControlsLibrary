<!--#include file="DanASPLib/Page.asp"-->
<!--#include file="DanASPLib/Button.asp"-->
<!--#include file="DanASPLib/Literal.asp"-->
<%
    ' Submit button for the form.
    dim btnSend
    dim lblLiteral
    
    ' Initialises the controls on the form.
    Sub Page_Init()
      set btnSend = new Button
      btnSend.New_Button("btnSend")

      set lblLiteral = new Literal
      Call lblLiteral.New_Literal()
    End Sub

    ' Initialises the page's state.
    Sub Page_Load()
        if Not IsPostBack() then
            lblLiteral.Text = "Page_Load()"
        else
            lblLiteral.Text ="Page_Load() on Postback"
        end if
    End Sub

    ' Handles the Click "event" of the btnSend control.
    Sub btnSend_Click()
        Response.Write "btnSend_Click()"
    End Sub

    ' Page Setup.
    Call Page_Init()
    Call Page_Load()

    ' Check the form that the btnSend control has posted a value.
    if isEmpty(Request(btnSend.ID)) = false then
        Call btnSend_Click()
    end if
%>