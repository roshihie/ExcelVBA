
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
	
		LEVNOPtn �� ����
		�q�b�g�����Ƃ�
			intLEVNO  =  �ݒ�
			If intLEVNO = intOCCRSLEVNO Then
		
		��OCCRS�s�݂̂̏ꍇ�ɔz�񐔂𒆊ԉ��Z�G���A�ɏ�Z����
		
		PICPtn   �� ����
		�q�b�g�����Ƃ�
			isPIC            =  True
			intPICLEVNO      =  intLEVNO
			ItemPty.strTYPE  =  �ݒ�l
			ItemPty.intSIZE  =  �ݒ�l
			
		USAGEPtn �� ����
		�q�b�g�����Ƃ�
			ItemPty.strUSAGE =  �ݒ�l
			
		OCCRSPtn �� ����
		�q�b�g�����Ƃ�
			intDefOCCRS =  �ݒ�l
			intOCCRSLEVNO = intLEVNO
			
		PERIODPtn�� ����
		�q�b�g�����Ƃ�
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
	
