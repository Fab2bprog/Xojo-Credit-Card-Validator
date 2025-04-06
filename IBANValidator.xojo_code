#tag Class
Protected Class IBANValidator
	#tag Method, Flags = &h0
		Function CleanIBAN(iban As String) As String
		  iban = ReplaceAll(iban, " ", "")
		  Return iban.Uppercase
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsAlpha(s As String) As Boolean
		  // Returns True if a string contains only letters A–Z
		  For i As Integer = 0 To s.Length - 1
		    Var ch As String = s.Mid(i + 1, 1)
		    If Not (ch >= "A" And ch <= "Z") Then
		      Return False
		    End If
		  Next
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsNumeric(s As String) As Boolean
		  For i As Integer = 0 To s.Length - 1
		    Var ch As String = s.Mid(i + 1, 1)
		    If Not (ch >= "0" And ch <= "9") Then
		      Return False
		    End If
		  Next
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsValid(iban As String) As Boolean
		  // Step 1: Clean the IBAN input (remove spaces and convert to uppercase)
		  iban = CleanIBAN(iban)
		  
		  // Step 2: Check length
		  If iban.Length < 15 Or iban.Length > 34 Then
		    Return False
		  End If
		  
		  // Step 3: Check if it starts with 2 letters (country code) followed by digits
		  If Not IsAlpha(iban.Left(2)) Then Return False
		  If Not IsNumeric(iban.Mid(5)) Then Return False // the last part must be mostly numeric after rearrangement
		  
		  // Step 4: Move the first 4 characters to the end of the string
		  Var rearranged As String = iban.Mid(5) + iban.Left(4)
		  
		  // Step 5: Replace each letter with digits (A=10, B=11, ..., Z=35)
		  Var numericString As String = ""
		  For i As Integer = 0 To rearranged.Length - 1
		    Var ch As String = rearranged.Mid(i + 1, 1)
		    If IsNumeric(ch) Then
		      numericString = numericString + ch
		    ElseIf IsAlpha(ch) Then
		      Var value As Integer = Asc(ch.Uppercase) - Asc("A") + 10
		      numericString = numericString + Str(value)
		    Else
		      Return False // Invalid character
		    End If
		  Next
		  
		  // Step 6: Perform the modulo 97 check
		  Try
		    Var remainder As Integer = Mod97(numericString)
		    Return (remainder = 1)
		  Catch e As RuntimeException
		    Return False
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Mod97(numberString As String) As Integer
		  // Performs mod 97 on a string representing a large integer
		  Var chunk As String
		  Var temp As Integer = 0
		  Var i As Integer = 0
		  
		  While i < numberString.Length
		    chunk = Str(temp) + numberString.Mid(i + 1, 9 - Str(temp).Length)
		    temp = Val(chunk) Mod 97
		    i = i + (9 - Str(temp).Length)
		  Wend
		  
		  Return temp
		End Function
	#tag EndMethod


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
