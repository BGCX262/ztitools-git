<%
	
	Class CRc4
		private sbox(255)
		private key(255)


		'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
		':::  This routine called by EnDeCrypt function. Initializes the
		':::  sbox and the key array)
		'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
		private Sub RC4Initialize(strPwd)
		   dim tempSwap
		   dim a
		   dim b
		   dim intLength

		   intLength = len(strPwd)
		   For a = 0 To 255
		      key(a) = asc(mid(strpwd, (a mod intLength)+1, 1))
		      sbox(a) = a
		   next

		   b = 0
		   For a = 0 To 255
		      b = (b + sbox(a) + key(a)) Mod 256
		      tempSwap = sbox(a)
		      sbox(a) = sbox(b)
		      sbox(b) = tempSwap
		   Next
		End Sub


		public default Function EnDeCrypt(plaintxt, psw)
		   dim temp
		   dim a
		   dim i
		   dim j
		   dim k
		   dim cipherby
		   dim cipher

		   i = 0
		   j = 0

		   RC4Initialize psw

		   For a = 1 To Len(plaintxt)
		      i = (i + 1) Mod 256
		      j = (j + sbox(i)) Mod 256
		      temp = sbox(i)
		      sbox(i) = sbox(j)
		      sbox(j) = temp

		      k = sbox((sbox(i) + sbox(j)) Mod 256)

		      cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
		      cipher = cipher & Chr(cipherby)
		   Next

		   EnDeCrypt = cipher
		End Function
		
	End Class
%>
