<% 
    public const TextModeSingleLine = 0
    public const TextModeMultiLine = 1
    
    Class TextBox
        private m_ID
        private m_Text
        private m_CSSClass
        private m_TextMode
        private m_ReadOnly
        private m_Rows
        private m_Columns
        private m_Disabled

        public sub Class_Initialize()
            m_Text= ""
            m_TextMode = TextModeSingleLine
            m_Rows = -1
            m_Columns = -1
            m_ReadOnly = false
            m_Disabled = false
        end sub

        private property Let ID(pID)
            m_ID = pID
        end property

        public property Get ID
            ID = m_ID
        end property

        public property Let Text(pText)
            m_Text = pText
        end property

        public property Get Text
            Text = m_Text
        end property

        public property Let CSSClass(pCSSClass)
            m_CSSClass = pCSSClass
        end property

        public property Get CSSClass
            CSSClass = m_CSSClass
        end property

        public property Let TextMode(pTextMode)
            m_TextMode = pTextMode
        end property

        public property Get TextMode
            TextMode = m_TextMode
        end property

        public property Let Rows(pRows)
            m_Rows = pRows
        end property

        public property Get Rows
            Rows = m_Rows
        end property

        public property Let Columns(pColumns)
            m_Columns = pColumns
        end property

        public property Get Columns
            Columns = m_Columns
        end property

        public property Let ReadOnly(pReadOnly)
            m_ReadOnly = pReadOnly
        end property

        public property Get ReadOnly
            ReadOnly = m_ReadOnly
        end property

        public property Let Disabled(pDisabled)
            m_Disabled = pDisabled
        end property

        public property Get Disabled
            Disabled = m_Disabled
        end property

        public sub New_TextBox(pID)
            ID = pID

            if (isEmpty(Request.Form(ID)) = false) then
                Text = Request.Form(ID)
            end if
        end Sub

        public default sub Render()
            if TextMode = TextModeSingleLine then
                Response.Write "<input type=""text"" value=""" & Text & """ name=""" & ID & """"

                if CSSClass <> "" then
                  Response.Write " class=""" & CSSClass & """"
                end if

                if Disabled = true then
                  Response.Write " disabled"
                end if

                if ReadOnly = true then
                  Response.Write " readonly"
                end if

                Response.Write ">"
            elseif TextMode = TextModeMultiLine then
                Response.Write "<textarea name=""" & ID & """>"

                if CSSClass <> "" then
                  Response.Write " class=""" & CSSClass & """"
                end if

                if Rows > -1 then
                  Response.Write " rows=""" & Rows & """"
                end if

                if Columns > -1 then
                  Response.Write " cols=""" & Columns & """"
                end if

                if Disabled = true then
                  Response.Write " disabled"
                end if

                if ReadOnly = true then
                  Response.Write " readonly"
                end if

                Response.Write Text
                Response.Write "</textarea>"
            end if
        end sub
    End Class
%>