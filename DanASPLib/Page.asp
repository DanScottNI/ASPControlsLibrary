<%    
    Class BasePageClass
        dim postBackDictionary

        public sub Class_Initialize()
            set postBackDictionary = Server.CreateObject("Scripting.Dictionary")
            redim postBackHandlers(-1)
        end sub

        public Sub NewPostBackHandler(pID, pName)
            if (postBackDictionary.Exists(pID)) then
                postBackDictionary(pID) = pName
            else
                Call postBackDictionary.Add(pID, pName)
            end if
        end sub

        public Sub RemovePostBackHandler(pID)
            if (postBackDictionary.Exists(pID)) then
                Call postBackDictionary.Remove(pID)
            end if
        end sub

        public sub ExecutePostBacks()
            
            dim a
            dim i
            a=postBackDictionary.Keys

            for i = 0 to postBackDictionary.Count - 1
            
                'if isEmpty(Request(btnTest.ID)) = false then
                if isEmpty(Request.Form(a(i))) = false then
                    dim postBackEvt

                    ' Attempt to find the correct event
                    set postBackEvt = GetMethodRef(postBackDictionary(a(i)))

                    if not postBackEvt is nothing then
                        Call postBackEvt()
                    end if
                end if
            next
        end sub

        public sub Load()
            dim pageInit
            dim pageLoad

            ' Attempt to find Page_Init
            set pageInit = GetMethodRef("Page_Init")

            if not pageInit is nothing then
                Call pageInit()
            end if

            ' Attempt to find Page_Load
            set pageLoad = GetMethodRef("Page_Load")

            if not pageLoad is nothing then
                Call pageLoad()
            end if

            if IsPostBack() then
                ExecutePostBacks()
            end if
        end sub
    End Class

    Function IsPostBack()
      IsPostBack = Request.ServerVariables("REQUEST_METHOD")= "POST"
    End Function

    Public Function GetMethodRef(pMethodName)
		On Error Resume Next
			Set GetMethodRef = Nothing
			Set GetMethodRef = GetRef(pMethodName)		
		On error Goto 0
	End Function

    dim Page
    set Page = nothing
    set Page = new BasePageClass
%>