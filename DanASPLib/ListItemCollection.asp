<%
    class ListItemCollection
        dim col

        private sub Class_Initialize()
            set col = server.createObject("System.Collections.ArrayList")
        end sub

        Private Sub Class_Terminate()  
            Set col = Nothing
        End Sub  

        public sub AddItem(obj)
            col.Add(obj)
        end sub

        public function Item(pIndex)
           set Item = col.Item(CINT(pIndex))
        end function

        public function Count()
            Count = col.Count
        end function
    end class

    class ListItem
        private m_Text
        private m_Value
        private m_Selected

        public property Let Text(pText)
            m_Text = pText
        end property

        public property Get Text
            Text = m_Text
        end property

        public property Let Value(pValue)
            m_Value = pValue
        end property

        public property Get Value
            Value = m_Value
        end property

        public property Let Selected(pSelected)
            m_Selected = pSelected
        end property

        public property Get Selected
            Selected = m_Selected
        end property
    end class
%>