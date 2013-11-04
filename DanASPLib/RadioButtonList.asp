<%
    class RadioButtonListItemCollection
        dim col 
    
        private sub Class_Initialize()
            set col = nothing
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
    
    class RadioButtonListItem
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
    
    ' RadioButtonList Class for HTML buttons
    Class RadioButtonList
        private m_ID
        private m_SelectedValue
        private m_Items

        private sub Class_Initialize()
            set m_Items = nothing
            set m_Items = new RadioButtonListItemCollection
        end sub
    
        private sub class_terminate()
            set m_Items = nothing
        end sub

        private property Let ID(pID)
            m_ID = pID
        end property
    
        public property Get ID
            ID = m_ID
        end property
    
        public property Let SelectedValue(pSelectedValue)
            m_SelectedValue = pSelectedValue
        end property
    
        public property Get SelectedValue
            SelectedValue = m_SelectedValue
        end property
        
        public sub AddListItem(pText, pValue)
            dim obj
            set obj = nothing
            set obj = new DropDownListItem
            
            obj.Text = pText
            obj.Value = pValue

            m_Items.AddItem(obj)

            set obj = nothing
        end sub

        public Sub New_RadioButtonList(pID)
            ID = pID

            if (isEmpty(Request.Form(ID)) = false) then
                SelectedValue = Request.Form(ID)
            end if
        end Sub
    
        public default sub Render()
            dim i

            for i = 0 to m_Items.Count() - 1
                dim o : set o = nothing
                set o = m_Items.Item(i)

                Response.Write "<input id=""" & ID & "$" & i & """ type=""radio"" name=""" & ID & """ value=""" & o.Value & """"

                if SelectedValue = o.Value then
                    Response.Write " checked=""checked"""
                end if

                Response.Write " /><label for=""" & ID & "$" & i & """>" & o.Text & "</label>" & vbCrLf

                set o = nothing
            next
        end sub
    End Class
%>