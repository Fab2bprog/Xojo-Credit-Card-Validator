#tag Class
Protected Class USBankAccountValidator
	#tag Method, Flags = &h0
		Sub Constructor(routingNumberParam As String, accountNumberParam As String)
		  RoutingNumber = routingNumberParam.Trim
		  AccountNumber = accountNumberParam.Trim
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetBankInfo() As Dictionary
		  Dim info As New Dictionary
		  
		  ' In a real implementation, this could call an API
		  ' or a database to get information about the bank
		  
		  ' Demo code only
		  Dim firstDigit As Integer = Val(RoutingNumber.Left(1))
		  
		  Select Case firstDigit
		  Case 0, 1, 2, 3
		    info.Value("region") = "Eastern United States"
		  Case 4, 5, 6
		    info.Value("region") = "Central United States"
		  Case 7, 8
		    info.Value("region") = "Western United States"
		  Case 9
		    info.Value("region") = "Reserved for electronic transfers"
		  End Select
		  
		  Return info
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsValid() As Boolean
		  Return IsValidRoutingNumber() And IsValidAccountNumber()
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsValidAccountNumber() As Boolean
		  ' Check length (varies by bank)
		  If AccountNumber.Length < MIN_ACCOUNT_LENGTH Or AccountNumber.Length > MAX_ACCOUNT_LENGTH Then
		    Return False
		  End If
		  
		  ' Check that it contains only digits
		  If Not IsNumeric(AccountNumber) Then
		    Return False
		  End If
		  
		  ' Note: Bank-specific validations could be added here
		  ' but there is no single standard for account number validation
		  
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsValidRoutingNumber() As Boolean
		  ' Check length
		  
		  If RoutingNumber.Length <> ROUTING_NUMBER_LENGTH Then
		    Return False
		  end if
		  
		  ' Check that it contains only digits
		  If Not IsNumeric(RoutingNumber) Then
		    Return False
		  End If
		  
		  ' Check the ABA validation algorithm
		  ' The routing number uses a specific algorithm:
		  ' [3(d1 + d4 + d7) + 7(d2 + d5 + d8) + (d3 + d6 + d9)] mod 10 = 0
		  Dim sum As Integer = 0
		  sum = sum + 3 * Val(RoutingNumber.Mid(1, 1))
		  sum = sum + 7 * Val(RoutingNumber.Mid(2, 1))
		  sum = sum + 1 * Val(RoutingNumber.Mid(3, 1))
		  sum = sum + 3 * Val(RoutingNumber.Mid(4, 1))
		  sum = sum + 7 * Val(RoutingNumber.Mid(5, 1))
		  sum = sum + 1 * Val(RoutingNumber.Mid(6, 1))
		  sum = sum + 3 * Val(RoutingNumber.Mid(7, 1))
		  sum = sum + 7 * Val(RoutingNumber.Mid(8, 1))
		  sum = sum + 1 * Val(RoutingNumber.Mid(9, 1))
		  
		  Return (sum Mod 10) = 0
		End Function
	#tag EndMethod


	#tag Note, Name = Licence MIT
		
		MIT License
		
		Copyright (c) 2025 Fabrice Garcia,  20290 Borgo,  Corsica Island , France, Europe.
		
		Permission is hereby granted, free of charge, to any person obtaining a copy
		of this software and associated documentation files (the "Software"), to deal
		in the Software without restriction, including without limitation the rights
		to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
		copies of the Software, and to permit persons to whom the Software is
		furnished to do so, subject to the following conditions:
		
		The above copyright notice and this permission notice shall be included in all
		copies or substantial portions of the Software.
		
		THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
		IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
		FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
		AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
		LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
		OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
		SOFTWARE.
		
		
	#tag EndNote


	#tag Property, Flags = &h0
		AccountNumber As string
	#tag EndProperty

	#tag Property, Flags = &h0
		RoutingNumber As string
	#tag EndProperty


	#tag Constant, Name = MAX_ACCOUNT_LENGTH, Type = Double, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"17"
	#tag EndConstant

	#tag Constant, Name = MIN_ACCOUNT_LENGTH, Type = Double, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"4"
	#tag EndConstant

	#tag Constant, Name = ROUTING_NUMBER_LENGTH, Type = Double, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"9"
	#tag EndConstant


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
		#tag ViewProperty
			Name="AccountNumber"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="RoutingNumber"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
