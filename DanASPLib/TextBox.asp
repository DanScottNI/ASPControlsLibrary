<%   Class TextBox
        private m_ID

        public sub Class_Initialize()
            
            m_Text= ""    
        end sub

        public property Let ID(pID)
            m_ID = pID
        end property

        public property Get ID
            ID = m_ID
        end property

        private m_Text

        public property Let Text(pText)
            m_Text = pText
        end property

        public property Get Text
            Text = m_Text
        end property

        public Sub New_TextBox(pID)
            ID = pID

            if (isEmpty(Request.Form(pID)) = false) then
                Text = Request.Form(pID)
            end if
        end Sub

        public default sub Render()
            Response.Write "<input type=""text"" value=""" & Text & """ name=""" & ID & """>"
        end sub
    End Class
%>