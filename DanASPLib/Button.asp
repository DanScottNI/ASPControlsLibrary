<%
    Class Button
        private m_ID

        public sub Class_Initialize()
            m_Text= "Submit that beotch"    
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

        public Sub New_Button(pID)
            ID = pID
        end Sub

        public sub Render()
            Response.Write "<input type=""submit"" value=""" & Text & """ name=""" & ID & """>"
        end sub
    End Class
%>