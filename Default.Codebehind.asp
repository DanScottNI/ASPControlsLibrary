<!--#include file="DanASPLib/Page.asp"-->
<!--#include file="DanASPLib/Button.asp"-->
<!--#include file="DanASPLib/Literal.asp"-->
<!--#include file="DanASPLib/TextBox.asp"-->
<%
    ' Submit button for the form.
    dim btnSend
    dim btnTest
    dim lblLiteral
    dim txtEdit
    
    ' Initialises the controls on the form.
    Sub Page_Init()
      set btnSend = new Button
      btnSend.New_Button("btnSend")

      set btnTest = new Button
      btnTest.New_Button("btnTest")

      set lblLiteral = new Literal
      Call lblLiteral.New_Literal()

      set txtEdit = new TextBox
      txtEdit.New_TextBox("txtEdit")
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

    Sub btnTest_Click()
        Response.Write "btnTest_Click()"
    End Sub

    Call Page.Load()
%>