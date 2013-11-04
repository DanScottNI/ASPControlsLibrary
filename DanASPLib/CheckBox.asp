<%
    Class CheckBox
        private m_ID
        private m_Text
        private m_Checked

        public sub Class_Initialize()
            m_Text= ""
            m_Checked = false
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

        public property Let Checked(pChecked)
            m_Checked = pChecked
        end property

        public property Get Checked
            Checked = m_Checked
        end property

        public sub New_CheckBox(pID)
            ID = pID

            if (isEmpty(Request.Form(ID)) = false) then
                Checked = true
            end if
        end Sub

        public default sub Render()
            Response.Write "<input type=""checkbox"" value=""" & Text & """ id=""" & ID & """ name=""" & ID & """"
            
            if Checked then
                Response.Write " checked"
            end if

            Response.Write ">"
        end sub
    End Class
%>