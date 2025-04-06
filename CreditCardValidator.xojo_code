#tag Class
Protected Class CreditCardValidator
	#tag Method, Flags = &h0
		Function CheckLuhn(number As String) As Boolean
		  // Luhn algorithm to check if a card number is mathematically valid
		  
		  Var sum As Integer = 0
		  Var shouldDouble As Boolean = False
		  
		  For i As Integer = number.Length - 1 DownTo 0
		    Var digit As Integer = Val(Mid(number, i + 1, 1))
		    
		    If shouldDouble Then
		      digit = digit * 2
		      If digit > 9 Then digit = digit - 9
		    End If
		    
		    sum = sum + digit
		    shouldDouble = Not shouldDouble
		  Next
		  
		  Return (sum Mod 10 = 0)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CleanNumber(number As String) As String
		  // Utility method to clean the input (remove spaces and dashes)
		  
		  number = ReplaceAll(number, " ", "")
		  number = ReplaceAll(number, "-", "")
		  Return number
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetCardType(cardNumber As String) As CardType
		  // Detects the type of credit card based on prefix and length
		  
		  cardNumber = CleanNumber(cardNumber)
		  
		  If cardNumber = "" Or Not IsNumeric(cardNumber) Then Return CardType.Unknown
		  
		  Var length As Integer = cardNumber.Length
		  
		  // Visa: Starts with 4 and has 13, 16 or 19 digits
		  If cardNumber.Left(1) = "4" And (length = 13 Or length = 16 Or length = 19) Then
		    Return CardType.Visa
		  End If
		  
		  // MasterCard: Starts with 51–55 or 2221–2720 and has 16 digits
		  Var prefix2 As Integer = Val(cardNumber.Left(2))
		  Var prefix4 As Integer = Val(cardNumber.Left(4))
		  If (prefix2 >= 51 And prefix2 <= 55 Or (prefix4 >= 2221 And prefix4 <= 2720)) And length = 16 Then
		    Return CardType.MasterCard
		  End If
		  
		  // American Express: Starts with 34 or 37 and has 15 digits
		  If (cardNumber.Left(2) = "34" Or cardNumber.Left(2) = "37") And length = 15 Then
		    Return CardType.AmericanExpress
		  End If
		  
		  // Discover: Starts with 6011, 65, or 644–649 and has 16 digits
		  If (cardNumber.Left(4) = "6011" Or cardNumber.Left(2) = "65" Or _
		    (Val(cardNumber.Left(3)) >= 644 And Val(cardNumber.Left(3)) <= 649)) And length = 16 Then
		    Return CardType.Discover
		  End If
		  
		  // Diners Club: Starts with 300–305, 36, 38, or 39 and has 14 digits
		  If (Val(cardNumber.Left(3)) >= 300 And Val(cardNumber.Left(3)) <= 305 Or _
		    cardNumber.Left(2) = "36" Or cardNumber.Left(2) = "38" Or cardNumber.Left(2) = "39") And length = 14 Then
		    Return CardType.DinersClub
		  End If
		  
		  Return CardType.Unknown
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsValid(cardNumber As String, expectedType As CardType) As Boolean
		  // Returns True if the card number is structurally valid and matches the expected card type
		  
		  cardNumber = CleanNumber(cardNumber)
		  
		  If Not IsNumeric(cardNumber) Then Return False
		  
		  Var detectedType As CardType = GetCardType(cardNumber)
		  
		  If detectedType <> expectedType Then
		    Return False // The number does not match the selected card type
		  End If
		  
		  If Not CheckLuhn(cardNumber) Then
		    Return False
		  End If
		  
		  Return True
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


	#tag Enum, Name = CardType, Type = Integer, Flags = &h0
		Unknown
		  Visa
		  MasterCard
		  AmericanExpress
		  Discover
		DinersClub
	#tag EndEnum


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
