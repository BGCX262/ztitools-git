<%
	'----------------------------------------------------------
	'-- This class is used to do better string concatenation --
	'----------------------------------------------------------
	Class FastString
		Dim stringArray,growthRate,numItems
		
		Private Sub Class_Initialize()
			growthRate = 50: numItems = 0
			ReDim stringArray(growthRate)
		End Sub
		
		Private Sub Class_Terminate()
			Erase stringArray
		End Sub
		
		Public Sub Append(ByVal strValue)
			' next line prevents type mismatch error if strValue is null. Performance hit is negligible.
			strValue=strValue & ""
			If numItems > UBound(stringArray) Then 
				ReDim Preserve stringArray(UBound(stringArray) + growthRate)
			End If
			stringArray(numItems) = strValue
			numItems = numItems + 1
		End Sub
		
		Public Sub Reset
			Erase stringArray
			Class_Initialize
		End Sub
		
		Public Function concat() 
			Redim Preserve stringArray(numItems) 
			concat = Join(stringArray, "")
		End Function
		
	End Class 
%>