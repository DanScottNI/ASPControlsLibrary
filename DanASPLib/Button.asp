<%
    ' Button class for HTML buttons
    Class Button
        private m_ID
        private m_Text

        public sub Class_Initialize()
            m_Text= "Submit"
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

        public Sub New_Button(pID)
            ID = pID
            Call Page.NewPostBackHandler(pID, pID & "_Click")
        end Sub

        public default sub Render()
            Response.Write "<input type=""submit"" value=""" & Text & """ name=""" & ID & """>"
        end sub
    End Class
%>