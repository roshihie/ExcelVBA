
Public Type typProperty
    strTYPE    As String
    intSIZE    As inte
    intOCCRS   As inte
    strUSAGE   As String
End Type





Private Sub ProcParse(StrReadArea()  As String, _
					  intLength      As Integer)
					  
	
	Const conPERIODPtn  = ".+\.[ ]"
	
	
	Dim ItemPty       As typProperty
	
	Dim intLEVNO      As Integer
	Dim intOCCRSLEVNO As Integer
	
	
	Call ProcInitPty(ItemPty, isPIC)

	For intIdx  =  0  To  Ubound(strReadArea)
	
		LEVNOPtn の 検索
		ヒットしたとき
			intLEVNO  =  設定
			If intLEVNO = intOCCRSLEVNO Then
		
		●OCCRS行のみの場合に配列数を中間加算エリアに乗算する
		
		PICPtn   の 検索
		ヒットしたとき
			isPIC            =  True
			intPICLEVNO      =  intLEVNO
			ItemPty.strTYPE  =  設定値
			ItemPty.intSIZE  =  設定値
			
		USAGEPtn の 検索
		ヒットしたとき
			ItemPty.strUSAGE =  設定値
			
		OCCRSPtn の 検索
		ヒットしたとき
			intDefOCCRS =  設定値
			intOCCRSLEVNO = intLEVNO
			
		PERIODPtnの 検索
		ヒットしたとき
			If isPIC    And _
			   intPICLEVNO  =  intOCCRSLEVNO  Then
			   
			    ItemPty.intOCCRS  =  intDefOCCRS
				intWorkLen = FuncLenCal(ItemPty)
				
				CallProcInitPty(ItemPty)
			End If

















Private Sub ProcInitPty(ItemPty As typProperty, _
						isPIC   As Boolean )

	ItemPty.strTYPE  =  ""
	ItemPty.intSIZE  =  0
	ItemPty.intOCCRS =  1
	ItemPty.strUSAGE =  ""
	
	isPIC            =  False
	
End Sub
	
