<!--#include file="DanASPLib/Page.asp"-->
<!--#include file="DanASPLib/Button.asp"-->
<!--#include file="DanASPLib/Literal.asp"-->
<!--#include file="DanASPLib/TextBox.asp"-->
<!--#include file="DanASPLib/DropDownList.asp"-->
<!--#include file="DanASPLib/RadioButtonList.asp"-->
<!--#include file="DanASPLib/CheckBox.asp"-->
<%
    ' Submit button for the form.
    dim btnSend : set btnSend = nothing
    dim btnTest : set btnTest = nothing
    dim lblLiteral : set lblLiteral = nothing
    dim txtEdit : set txtEdit = nothing
    dim cbTest : set cbTest = nothing
    dim rblTest : set rblTest = nothing
    dim chkTest : set chkTest = nothing

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

      set cbTest = new DropDownList
      cbTest.New_DropDownList("cbTest")

      Call cbTest.AddListItem("Test", "1")
      Call cbTest.AddListItem("Test2", "2")
      Call cbTest.AddListItemWithColour("Test3", "3", "red")

      set rblTest = new RadioButtonList
      rblTest.New_RadioButtonList("rblTest")

      Call rblTest.AddListItem("Test1", "1")
      Call rblTest.AddListItem("Test2", "2")
      Call rblTest.AddListItem("Test3", "3")

      'rblTest.SelectedValue = "2"

      set chktest = new CheckBox
      chkTest.New_CheckBox("chkTest")
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