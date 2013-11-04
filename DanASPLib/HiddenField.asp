<% 
    Class HiddenField
        private m_ID
        private m_Text

        public sub Class_Initialize()
            m_Text= ""
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

        public sub New_HiddenField(pID)
            ID = pID

            if (isEmpty(Request.Form(ID)) = false) then
                Text = Request.Form(ID)
            end if
        end Sub

        public default sub Render()
            Response.Write "<input type=""hidden"" value=""" & Text & """ id=""" & ID & """ name=""" & ID & """>"
        end sub
    End Class
%>