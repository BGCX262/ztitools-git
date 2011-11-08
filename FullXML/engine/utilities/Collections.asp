<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: A simple class to simulate Collection with asp ::
	'::::::::::::::::::::::::::::::::::::::::::::::::::::
	Class Collection
		private m_arrCollection
		
		
		private sub class_initialize
			m_arrCollection = array()
		End sub
		
		
		private sub class_terminate
			clear
			m_arrCollection = empty
		end sub
		
		
		'-- returns the number od items in the array
		public property get Count
			Count = UBound(m_arrCollection) - LBound(m_arrCollection)
		end property
				
		
		'-- add an element at the end of the array
		public function Add(Value)
			redim preserve m_arrCollection(ubound(m_arrCollection)+1)	
			m_arrCollection(ubound(m_arrCollection)) = Value
		end function
		
		
		'-- get an item by index
		public Property Get Item(Index)
			If (Index >= LBound(m_arrCollection)) and Index <= UBound(m_arrCollection) Then
				Item = m_arrCollection(Index)
			Else
				Item = empty
			End If
		end Property
					
		
		'-- clear the array
		public function Clear()
			dim i
			for i = LBound(m_arrCollection) to UBound(m_arrCollection)
				'set m_arrCollection(i) = Nothing
				m_arrCollection = Empty
			next
			redim  m_arrCollection(0)
		End function
		
		
		'-- return the array
		public Function ToArray()
			ToArray = m_arrCollection
		End Function
		
		'-- return the array as serialzed string
		public Function ToString()
			ToString = join(m_arrCollection, ";")
		End Function
		
	End Class
%>