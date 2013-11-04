<%
    Class Literal
        private m_Text

        public sub Class_Initialize()
            Text= ""
        end sub

        public property Let Text(pText)
            m_Text = pText
        end property

        public property Get Text
            Text = m_Text
        end property

        public Sub New_Literal()
        end Sub

        public default sub Render()
            Response.Write Text
        end sub
    End Class
%>