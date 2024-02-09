Option Explicit
																								'Function List
'************************************************************************************************************************************************************************************************************
'1. Fn_WebMyTc_CreateVendor()
'2. Fn_WebMyTc_CreateBidPackage()
'3. Fn_WebMyTc_CreateCommercialPart()
'4. Fn_WebMyTc_CreateCompanyContact()
'5. Fn_WebMyTc_CreateVendorPart()
'6. Fn_WebMyTc_VendorOperations()
'7. Fn_Web_ImportSpecification()
'8. Fn_WebMyTc_ManageGlobalAlternateOperations()
'9. Fn_WebMyTc_CreateVendorPartSearchReults()
'10. Fn_WebMyTc_ChangeVendorOperations
'************************************************************************************************************************************************************************************************************
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebMyTc_CreateVendor
'@@
'@@    Description				 :	Function Used To Create New Vendor
'@@
'@@    Parameters			   :	1.dicVendor: Vendor Full Information Dictionary Object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	dicVendor("Name")="User1"
'@@												dicVendor("Description")="Spare Part Vendor"
'@@												dicVendor("CreateAlternateID")="Off"
'@@												dicVendor("Contact")="9822755994"
'@@												dicVendor("Address")="Pune
'@@												Call Fn_WebMyTc_CreateVendor(dicVendor)
'@@												
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									20-Apr-2011						1.0																								Sunny Ruparel
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Ganesh Bhosale										12-Mar-2014						1.1								Added Code check existence of THWebElement object to select type if it exist
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_WebMyTc_CreateVendor(dicVendor)
	GBL_FAILED_FUNCTION_NAME="Fn_WebMyTc_CreateVendor"
 	'Variable Declaration
	Dim ObjVendor,ObjVendorInfo,ObjVendorAddInfo,iCounter,objMDR
	Dim dicItems,dicKeys,strWEBMenuPath,strMenu, crrType
    Dim ObjMyTeamCenter
	Set objMDR = CreateObject("Mercury.DeviceReplay")

	Fn_WebMyTc_CreateVendor=False
	'Creating Objects Of "Vendor" Tables
	Set ObjVendor=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Vendor")
	Set ObjVendorAddInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Vendor").WebTable("AdditionalVendorInfo")
	Set ObjVendorInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Vendor").WebTable("VendorInfo")
    Set ObjMyTeamCenter = Browser("TeamcenterWeb").Page("MyTeamCenter")
	'Checking Existance Of "New Vendor" Dialog
	If Not ObjVendor.Exist(5) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("WebMyTc_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewVendor")
		Call Fn_Web_MenuOperation("Select",strMenu)
		Wait 5
	End If
	dicItems=dicVendor.Items
	dicKeys=dicVendor.Keys

	If dicVendor("Type")="" Then
		dicVendor("Type")="Vendor"
	End If
	If dicVendor("Type")<>"" Then
        If ObjMyTeamCenter.WebElement("THWebElement").Exist(1) Then
			'Do Nothing
		Else
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateVendor",ObjMyTeamCenter.WebElement("THWebElement"),"innertext","Type\*:") ''used code to handle FormType  object
		End If
		'Setting Item Type
		If ObjMyTeamCenter.WebElement("THWebElement").Exist(1) Then
			Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendor",ObjVendor,"ItemType")
			Wait 1
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateVendor",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",dicVendor("Type"))
			Wait 0,200
			If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Exist(2) = False then
				Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendor",ObjVendor,"ItemType")
				Wait 1
			End If
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
			wait(3)
			Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendor",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
			wait(3)
		End If
	End If
	wait(2)

	'Setting Vendor ID
	If dicVendor("ID")<>"" Then
		ObjVendorInfo.WebEdit("ID").Object.focus
		objMDR.SendString dicVendor("ID")
		wait(2)
	End If
    'Setting Revision
	If dicVendor("Revision")<>"" Then
		ObjVendorInfo.WebEdit("Revision").Object.focus
		objMDR.SendString dicVendor("Revision")
		wait(2)
	End If
	'Setting Vendor Name
	If dicVendor("Name")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendor", ObjVendorInfo, "Name",dicVendor("Name"))
		wait(2)
	End If
	'Setting Vendor Description
	If dicVendor("Description")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendor", ObjVendorInfo, "Description", dicVendor("Description"))
		wait(2)
	End If
	'UOM
	If dicVendor("UOM")<>"" Then
	End If
	'Setting Create Alternate ID Option
	If dicVendor("CreateAlternateID")<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_WebMyTc_CreateVendor", ObjVendorInfo, "CreateAltID", dicVendor("CreateAlternateID"))
		wait(2)
	End If
	'Setting Check Out On Create Option
	If dicVendor("CheckOut")<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_WebMyTc_CreateVendor", ObjVendorInfo, "CheckOutOnCrt", dicVendor("CheckOut"))
		wait(2)
	End If
    'Clicking On Next Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendor",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
	wait(2)
	'Setting Vendor Contact
	If dicVendor("Contact")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendor", ObjVendorAddInfo, "Contact", dicVendor("Contact"))
		wait(2)
	End If
	'Setting Vendor Address
	If dicVendor("Address")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendor", ObjVendorAddInfo,"Address", dicVendor("Address"))
		wait(2)
	End If
	'Setting Vendor Web Site Address
	If dicVendor("WebSite")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendor", ObjVendorAddInfo,"WebSite", dicVendor("WebSite"))
		wait(2)
	End If
	'Setting Vendor Phone
	If dicVendor("Phone")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendor", ObjVendorAddInfo,"Phone", dicVendor("Phone"))
		wait(2)
	End If
	'Setting Vendor Email
	If dicVendor("Email")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendor", ObjVendorAddInfo,"Email", dicVendor("Email"))
		wait(2)
	End If
	'Clicking On Finish Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendor",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Finish")

	For iCounter=0 to 5
		If ObjVendor.Exist Then
			wait(10)
		Else
			Exit For
		End If
	Next

	'Function Returns True
	Fn_WebMyTc_CreateVendor=True
	'Releasing Object
	Set ObjVendor=Nothing
	Set ObjVendorInfo=Nothing
	Set ObjVendorAddInfo=Nothing
	Set objMDR =Nothing

End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebMyTc_CreateBidPackage
'@@
'@@    Description				 :	Function Used To Create New Bid Package
'@@
'@@    Parameters			   :	1.dicBidPackage: Bid Package Full Information Dictionary Object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	dicBidPackage("Name")="Package1"
'@@												dicBidPackage("Description")="Demo Package"
'@@												dicBidPackage("CheckOut")="Off"
'@@												dicBidPackage("RequiredPurpose")="For Test"
'@@												Call Fn_WebMyTc_CreateBidPackage(dicBidPackage)
'@@												
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									20-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_WebMyTc_CreateBidPackage(dicBidPackage)
	GBL_FAILED_FUNCTION_NAME="Fn_WebMyTc_CreateBidPackage"
   'Variable Declaration
	Dim ObjPackage,ObjPackageInfo,ObjPackageRevInfo
	Dim strWEBMenuPath,strMenu,dicItems


	Fn_WebMyTc_CreateBidPackage=False
	Set ObjPackage=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BidPackage")
	Set ObjPackageInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BidPackage").WebTable("BidPackageInfo")
	Set ObjPackageRevInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BidPackage").WebTable("BidPackageRevInfo")

	If Not ObjPackage.Exist(5) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("WebMyTc_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewBidPackage")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	Wait 2

	Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateBidPackage",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType"),"innertext","Bid Package Information")
	Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType").Click 1,1,micLeftBtn
	
	dicItems=dicBidPackage.Items
	'Setting Bid Package ID
	If dicItems(0)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateBidPackage", ObjPackageInfo, "ID", dicItems(0))
	End If
	'Setting Bid Package Revision
	If dicItems(1)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateBidPackage", ObjPackageInfo, "Revision", dicItems(1))
	End If
	'Setting Bid Package Name
	If dicItems(2)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateBidPackage", ObjPackageInfo, "Name", dicItems(2))
	End If
	'Setting Bid Package Description
	If dicItems(3)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateBidPackage", ObjPackageInfo, "Description", dicItems(3))
	End If
	'Setting UOM
	If dicItems(4)<>"" Then
	End If
	'Setting Create Alternate ID Option
	If dicItems(5)<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_WebMyTc_CreateBidPackage", ObjPackageInfo, "CreateAltID", dicItems(5))
	End If
	'Setting Check Out On Create Option
	If dicItems(6)<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_WebMyTc_CreateBidPackage", ObjPackageInfo, "CheckOutOnCreate", dicItems(6))
	End If
	'Clicking On Next Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateBidPackage",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
	'Setting Bid Package Revision Required Purpose
	If dicItems(8)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateBidPackage", ObjPackageRevInfo, "RequiredPurpose", dicItems(8))
	End If
	'Clicking On Finish Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateBidPackage",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Finish")
	For iCounter=0 to 5
		If ObjPackage.Exist Then
			wait(10)
		Else
			Exit For
		End If
	Next
	'Function Returns True
	Fn_WebMyTc_CreateBidPackage=True
	'Releasing All Objects Of Tables
	Set ObjPackage=Nothing
	Set ObjPackageInfo=Nothing
	Set ObjPackageRevInfo=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebMyTc_CreateCommercialPart
'@@
'@@    Description				 :	Function Used To Create Commercial Part
'@@
'@@    Parameters			   :	1.dicCommercialPart: Commercial Part Full Information Dictionary Object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	dicCommercialPart("Name")="CmPart1"
'@@												dicCommercialPart("Description")="Test Commercial Part"
'@@												dicCommercialPart("CreateAlternateID")="Off"
'@@												dicCommercialPart("DesignRequired")="On"
'@@												Call Fn_WebMyTc_CreateCommercialPart(dicCommercialPart)
'@@												
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									22-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_WebMyTc_CreateCommercialPart(dicCommercialPart)
	GBL_FAILED_FUNCTION_NAME="Fn_WebMyTc_CreateCommercialPart"
   'Variable Declaration
   Dim ObjCmPart,ObjCmPartAddInfo,ObjCmPartInfo,ObjCmPartRevInfo
   Dim strWEBMenuPath,strMenu,dicItems,crrInfo
   Dim ObjMyTeamCenter

	Fn_WebMyTc_CreateCommercialPart=False
	'Creating Objects Of All Commercial Part Tables
	Set ObjCmPart=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CommercialPart")
	Set ObjCmPartAddInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CommercialPart").WebTable("CommercialPartAddInfo")
	Set ObjCmPartInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CommercialPart").WebTable("CommercialPartInfo")
	Set ObjCmPartRevInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CommercialPart").WebTable("CommercialPartRevInfo")
	Set ObjMyTeamCenter = Browser("TeamcenterWeb").Page("MyTeamCenter")
	'Checking Existance Of "CommercialPart" Dialog
	If Not ObjCmPart.Exist(5) Then
		'Calling "New->Commercial Part..." Menu Option
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("WebMyTc_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewCommercialPart")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
'Added By Ujwal Nagre  [05-June-2014] for MFK template..
	dicItems=dicCommercialPart.Items
	dicKeys=dicCommercialPart.Keys

	If dicCommercialPart("Type")="" Then
		dicCommercialPart("Type")="Commercial Part"
	End If
	If dicCommercialPart("Type")<>"" Then
         Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateCommercialPart",ObjMyTeamCenter.WebElement("THWebElement"),"innertext","Type\*:") 	''Added Code check existence of THWebElement object to select type if it exist
			'Setting Item Type
			If  ObjMyTeamCenter.WebElement("THWebElement").Exist(1) Then
				Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCommercialPart",ObjCmPart,"ItemType")
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateCommercialPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",dicCommercialPart("Type"))
				If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Exist(2) = False then
					Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCommercialPart",ObjCmPart,"ItemType")
				End If
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
				wait(1)
				Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCommercialPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
			End If
	End If
	wait(2)
	'Taking all values From Dictionory Object
	dicItems=dicCommercialPart.Items
	'Setting Commercial Part ID
	If dicItems(0)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCommercialPart", ObjCmPartInfo, "ID", dicItems(0))
	End If
	'Setting Commercial Part Revision
	If dicItems(1)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCommercialPart", ObjCmPartInfo, "Revision", dicItems(1))
	End If
	'Setting Commercial Part Name
	If dicItems(2)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCommercialPart", ObjCmPartInfo, "Name", dicItems(2))
	End If
	'Setting Commercial Part Description
	If dicItems(3)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCommercialPart", ObjCmPartInfo, "Description", dicItems(3))
	End If
	'Setting UOM
	If dicItems(4)<>"" Then
	End If
	'Setting Create Alternate ID Option
	If dicItems(5)<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_WebMyTc_CreateCommercialPart", ObjCmPartInfo, "CreateAltID", dicItems(5))
	End If
	'Setting Check Out on Create Option
	If dicItems(6)<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_WebMyTc_CreateCommercialPart", ObjCmPartInfo, "CheckOutOnCrt", dicItems(6))
	End If
	'Clicking On Next Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCommercialPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
	'Setting Design Required Option
	If dicItems(7)<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_WebMyTc_CreateCommercialPart", ObjCmPartAddInfo, "DesignReq", dicItems(7))
	End If
	'Clicking On Next Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCommercialPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
	'Setting Cm Part Revision Option
	If dicItems(8)<>"" Then
		crrInfo=ObjCmPartRevInfo.WebEdit("MakeOrBuy").GetROProperty("value")
		If Trim(crrInfo)<>Trim(dicItems(8)) Then
			Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCommercialPart",ObjCmPartRevInfo,"MakeOrBuy")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateCommercialPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType"),"innertext",dicItems(8))
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType").Click 1,1,micLeftBtn
		End If
	End If
	'Clicking On Finish Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCommercialPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Finish")
	For iCounter=0 to 5
		If ObjCmPart.Exist Then
			wait(10)
		Else
			Exit For
		End If
	Next
	Fn_WebMyTc_CreateCommercialPart=True
	'Creating Objects Of All Commercial Part Tables
	Set ObjCmPart=Nothing
	Set ObjCmPartAddInfo=Nothing
	Set ObjCmPartInfo=Nothing
	Set ObjCmPartRevInfo=Nothing
	Set ObjMyTeamCenter=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebMyTc_CreateCompanyContact
'@@
'@@    Description				 :	Function Used To Create New Company Contact
'@@
'@@    Parameters			   :	1.dicCompanyContact: Company Contact Full Information Dictionary Object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	dicCompanyContact("Title")="Mr.,Mr."
'@@												dicCompanyContact("FirstName")="Amol"
'@@												dicCompanyContact("LastName")="Lanke"
'@@												dicCompanyContact("Suffix")="W"
'@@												dicCompanyContact("Mobile")="9960606060"
'@@												dicCompanyContact("Description")="Client From Holand"
'@@												Call Fn_WebMyTc_CreateCompanyContact(dicCompanyContact)
'@@												
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									22-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_WebMyTc_CreateCompanyContact(dicCompanyContact)
	GBL_FAILED_FUNCTION_NAME="Fn_WebMyTc_CreateCompanyContact"
 	'Variable Declaration
	Dim ObjBusiness,ObjObjectType,ObjContactInfo
	Dim strWEBMenuPath,strMenu,dicItems,crrType,crrTitle
 	'Creating Object Related Company Contact Tables
	Set ObjBusiness=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BussinessObject")
	Set ObjObjectType=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BussinessObject").WebTable("BussinessObjectType")
	Set ObjContactInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BussinessObject").WebTable("CompanyContactInfo")
	Fn_WebMyTc_CreateCompanyContact=False
	If Not ObjBusiness.Exist(5) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewOther")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	dicItems=dicCompanyContact.Items

    crrType=ObjObjectType.WebEdit("Type").GetROProperty("value")
	If Trim(crrType)<>"Company Contact" Then
		Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCompanyContact",ObjObjectType,"TypeButton")
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateCompanyContact",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType"),"innertext","Company Contact")
		Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType").Click 1,1,micLeftBtn
		wait(2)
	End If
	
	'Clicking On Next Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCompanyContact",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
	If dicItems(0)<>"" Then
		crrTitle=ObjContactInfo.WebEdit("Title").GetROProperty("value")
		If Trim(crrTitle)<>Trim(dicItems(0)) Then
			Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCompanyContact",ObjContactInfo,"TitleButton")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateCompanyContact",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType"),"innertext",dicItems(0))
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType").Click 1,1,micLeftBtn
			wait(2)
		End If
	End If
	'Setting First Name 
	If dicItems(1)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyContact", ObjContactInfo, "FirstName", dicItems(1))
    End If
	'Setting Last Name 
	If dicItems(2)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyContact", ObjContactInfo, "LastName", dicItems(2))
    End If
	'Setting Suffix 
	If dicItems(3)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyContact", ObjContactInfo, "Suffix", dicItems(3))
    End If
	'Setting Business Phone 
	If dicItems(4)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyContact", ObjContactInfo, "BusinessPhone", dicItems(4))
    End If
	'Setting Home Phone 
	If dicItems(5)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyContact", ObjContactInfo, "HomePhone", dicItems(5))
    End If
	'Setting Mobile Number
	If dicItems(6)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyContact", ObjContactInfo, "Mobile", dicItems(6))
    End If
	'Setting Pager Number
	If dicItems(7)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyContact", ObjContactInfo, "Pager", dicItems(7))
    End If
	'Setting Email
	If dicItems(8)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyContact", ObjContactInfo, "Email", dicItems(8))
    End If
	'Setting Email
	If dicItems(9)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyContact", ObjContactInfo, "Description", dicItems(9))
    End If
	'Clicking On Finish Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCompanyContact",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Finish")
	For iCounter=0 to 5
		If ObjBusiness.Exist Then
			wait(10)
		Else
			Exit For
		End If
	Next
	Fn_WebMyTc_CreateCompanyContact=True
	'Releasing Object Related Company Contact Tables
	Set ObjBusiness=Nothing
	Set ObjObjectType=Nothing
	Set ObjContactInfo=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebMyTc_CreateCompanyLocation
'@@
'@@    Description				 :	Function Used To Create New Company Location
'@@
'@@    Parameters			   :	1.dicCompanyContact: Company Location Full Information Dictionary Object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	dicCompanyLocation("LocationCode")="12"
'@@												dicCompanyLocation("LocationType")="GLN,Global Location Number"
'@@												dicCompanyLocation("Street")="Phase1"
'@@												dicCompanyLocation("City")="Pune"
'@@												dicCompanyLocation("State")="Maharashtra"
'@@												dicCompanyLocation("PostalCode")="411057"
'@@												dicCompanyLocation("Country")="INDIA"
'@@												dicCompanyLocation("Region")="North"
'@@												dicCompanyLocation("URL")="www.ugs.com"
'@@												dicCompanyLocation("Description")="PLM Domain"
'@@												Call Fn_WebMyTc_CreateCompanyLocation(dicCompanyLocation)
'@@												
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									22-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_WebMyTc_CreateCompanyLocation(dicCompanyLocation)
	GBL_FAILED_FUNCTION_NAME="Fn_WebMyTc_CreateCompanyLocation"
	Dim ObjBusiness,ObjObjectType,ObjLocInfo
	Dim strWEBMenuPath,strMenu,dicItems,crrType,currLocType
	 'Creating Object Related Company Contact Tables
	Set ObjBusiness=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BussinessObject")
	Set ObjObjectType=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BussinessObject").WebTable("BussinessObjectType")
	Set ObjLocInfo=ObjBusiness.WebTable("CompanyLocationInfo")
	Fn_WebMyTc_CreateCompanyLocation=False
	'Checking Existance of New Business Object Dialog
	If Not ObjBusiness.Exist(5) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewOther")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	dicItems=dicCompanyLocation.Items

    crrType=ObjObjectType.WebEdit("Type").GetROProperty("value")
	If Trim(crrType)<>"Company Location" Then
		Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCompanyLocation",ObjObjectType,"TypeButton")
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateCompanyLocation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType"),"innertext","Company Location")
		Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType").Click 1,1,micLeftBtn
		wait(2)
	End If
	'Clicking On Next Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCompanyLocation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
	'Setting Company Name
	If dicItems(0)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyLocation", ObjLocInfo, "CompanyName", dicItems(0))
	End If
	'Setting Location Code
	If dicItems(1)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyLocation", ObjLocInfo, "LocationCode", dicItems(1))
	End If
	'Setting Location Type
	If dicItems(2)<>"" Then
		currLocType=ObjLocInfo.WebEdit("LocationType").GetROProperty("value")
		If Trim(currLocType)<>Trim(dicItems(2)) Then
			Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCompanyLocation",ObjLocInfo,"LocationType")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateCompanyLocation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType"),"innertext",dicItems(2))
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType").Click 1,1,micLeftBtn
			wait(2)
		End If
	End If
	'Setting Street Name
	If dicItems(3)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyLocation", ObjLocInfo, "Street", dicItems(3))
	End If
	'Setting City Name
	If dicItems(4)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyLocation", ObjLocInfo, "City", dicItems(4))
	End If
	'Setting State Name
	If dicItems(5)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyLocation", ObjLocInfo, "State", dicItems(5))
	End If
	'Setting Postal Code
	If dicItems(6)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyLocation", ObjLocInfo, "PostalCode", dicItems(6))
	End If
	'Setting Country Name
	If dicItems(7)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyLocation", ObjLocInfo, "Country", dicItems(7))
	End If
	'Setting Region
	If dicItems(8)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyLocation", ObjLocInfo, "Region", dicItems(8))
	End If
	'Setting URL
	If dicItems(9)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyLocation", ObjLocInfo, "URL", dicItems(9))
	End If
	'Setting Description
	If dicItems(10)<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateCompanyLocation", ObjLocInfo, "Description", dicItems(10))
	End If
	'Clicking On Finish Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateCompanyLocation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Finish")
	For iCounter=0 to 5
		If ObjBusiness.Exist Then
			wait(10)
		Else
			Exit For
		End If
	Next

	Fn_WebMyTc_CreateCompanyLocation=True
	 'Releasing Object Related Company Location Tables
	Set ObjBusiness=Nothing
	Set ObjObjectType=Nothing
	Set ObjLocInfo=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebMyTc_CreateVendorPart
'@@
'@@    Description				 :	Function Used To Create New Vendor Part
'@@
'@@    Parameters			   :	1.dicVendorPart: Vendor Part Full Information Dictionary Object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	dicVendorPart("PartNumber")="123456"
'@@												dicVendorPart("PartName")="VendorPart"
'@@												dicVendorPart("ID")="2486"
'@@												dicVendorPart("Description")="Test vendor Part"
'@@												dicVendorPart("Type")="Vendor Part"
'@@												dicVendorPart("DesignRequired")="On"
'@@												dicVendorPart("MakeOrBuy")="1,Make"
'@@
'@@												Call Fn_WebMyTc_CreateVendorPart(dicVendorPart)
'@@												
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									25-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_WebMyTc_CreateVendorPart(dicVendorPart)
	GBL_FAILED_FUNCTION_NAME="Fn_WebMyTc_CreateVendorPart"
	'Variable Declaration
	Dim ObjVendorPart,ObjPartInfo,ObjPartMasterInfo,ObjPartRevInfo
	Dim CurrValue,strWEBMenuPath,strMenu,dicItems
	Fn_WebMyTc_CreateVendorPart=False
	'Creating Objects Of Vendor Part Related Tables
	Set ObjVendorPart=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("VendorPart")
	Set ObjPartInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("VendorPart").WebTable("VendorPartInfo")
	Set ObjPartMasterInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("VendorPart").WebTable("VendorPartMasterInfo")
	Set ObjPartRevInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("VendorPart").WebTable("VendorPartRevInfo")

	'Checking Existance of "New Vendor Part" Dialog
	If Not ObjVendorPart.Exist(5) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("WebMyTc_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewVendorPart")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	'Retrieving All values from 'dicVendorPart' Dictionary Object
	dicItems=dicVendorPart.Items
	'Entering Vendor Part Info
	If dicItems(0)<>"" Then
		'Setting Part Number
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendorPart", ObjPartInfo, "PartNumber", dicItems(0))
	End If
	If dicItems(1)<>"" Then
		'Setting Part Name
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendorPart", ObjPartInfo, "PartName", dicItems(1))
	End If
	If dicItems(2)<>"" Then
		'Setting Vendor ID
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendorPart", ObjPartInfo, "VendorID", dicItems(2))
	End If
	If dicItems(3)<>"" Then
		'Setting Vendor Name
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendorPart", ObjPartInfo, "VendorName", dicItems(3))
	End If
	If dicItems(4)<>"" Then
		'Setting Location
		CurrValue=ObjPartInfo.WebEdit("Location").GetROProperty("value")
		If Trim(CurrValue)<>Trim(dicItems(4)) Then
			Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendorPart",ObjPartInfo,"Location")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateVendorPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType"),"innertext",dicItems(4))
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType").Click 1,1,micLeftBtn
		End If
		CurrValue=""
	End If
	If dicItems(5)<>"" Then
		'Setting Vendor Description
		Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_CreateVendorPart", ObjPartInfo, "Description", dicItems(5))
	End If
	If dicItems(6)<>"" Then
		'Setting Type
		CurrValue=ObjPartInfo.WebEdit("Type").GetROProperty("value")
		If Trim(CurrValue)<>Trim(dicItems(6)) Then
			Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendorPart",ObjPartInfo,"Type")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateVendorPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType"),"innertext",dicItems(6))
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType").Click 1,1,micLeftBtn
		End If
		CurrValue=""
	End If
	If dicItems(7)<>"" Then
		'Setting Unit Of Measure
		CurrValue=ObjPartInfo.WebEdit("UOM").GetROProperty("value")
		If Trim(CurrValue)<>Trim(dicItems(7)) Then
			Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendorPart",ObjPartInfo,"UOM")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateVendorPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType"),"innertext",dicItems(7))
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType").Click 1,1,micLeftBtn
		End If
		CurrValue=""
	End If
	If dicItems(8)<>"" Then
		'Setting Check Out Item Revision On Create Option
		Call Fn_Web_UI_CheckBox_Set("Fn_WebMyTc_CreateVendorPart", ObjPartInfo, "CheckOutOnCreate", dicItems(8))
	End If
	If dicItems(9)<>"" Then
		'Setting Create Alternate ID Option
		Call Fn_Web_UI_CheckBox_Set("Fn_WebMyTc_CreateVendorPart", ObjPartInfo, "AlternateID", dicItems(9))
	End If
	'Clicking On Next Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendorPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
	If dicItems(10)<>"" Then
		'Setting Create Alternate ID Option
		Call Fn_Web_UI_CheckBox_Set("Fn_WebMyTc_CreateVendorPart", ObjPartMasterInfo, "DesignRequired", dicItems(10))
	End If
	'Clicking On Next Button
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendorPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
	If dicItems(11)<>"" Then
		'Setting Unit Of Measure
		CurrValue=ObjPartRevInfo.WebEdit("MakeOrBuy").GetROProperty("value")
		If Trim(CurrValue)<>Trim(dicItems(11)) Then
			Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendorPart",ObjPartRevInfo,"MakeOrBuy")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_WebMyTc_CreateVendorPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType"),"innertext",dicItems(11))
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ObjectType").Click 1,1,micLeftBtn
		End If
		CurrValue=""
	End If
	'Clicking On Finish Button
	Wait(5)
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendorPart",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Finish")
	Wait(5)
	For iCounter=0 to 5
		If ObjVendorPart.Exist Then
			wait(10)
		Else
			Exit For
		End If
	Next
	'Function Returns True
	Fn_WebMyTc_CreateVendorPart=True
	'Releasing Objects Of Vendor Part Related Tables
	Set ObjVendorPart=Nothing
	Set ObjPartInfo=Nothing
	Set ObjPartMasterInfo=Nothing
	Set ObjPartRevInfo=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebMyTc_VendorOperations
'@@
'@@    Description				 :	Function Used To Add Vendor Role and Remove Vendor Role
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client	and vendor revision should be selected.
'@@
'@@    Examples					:	
'@@    Case : "AddVendorRole"	  dicVendor("sAction")="AddVendorRole"
'@@												dicVendor("VendorRole")="Suppiler"
'@@												dicVendor("VendorStatus")="pending"
'@@												dicVendor("CertificationStatus")="1"
'@@												Call Fn_WebMyTc_VendorOperations(dicVendor)
'@@
'@@	  Case : "Remove"  				  dicVendor("sAction")="RemoveVendorRole"
'@@												dicVendor("VendorRole")="Suppiler"
'@@												Call Fn_WebMyTc_VendorOperations(dicVendor)
'@@												
'@@	   History					 	:	
'@@													Developer Name								Date						Rev. No.						Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@													Ketan Raje									03-May-2011					1.0																					Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_WebMyTc_VendorOperations(dicVendor)
	GBL_FAILED_FUNCTION_NAME="Fn_WebMyTc_VendorOperations"
 	'Variable Declaration
	Dim ObjAddVendorRole, ObjRemoveVendorRole,iCounter
	Fn_WebMyTc_VendorOperations=False	
	Set ObjAddVendorRole = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AddVendorRole")
	Select Case dicVendor("sAction")
	Case "AddVendorRole"
				'Checking Existance Of "Add Vendor Role" Dialog
				If Not ObjAddVendorRole.Exist(5) Then
					strWEBMenuPath=Fn_LogUtil_GetXMLPath("WebMyTc_Menu")
					strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "AddVendorRole")
					Call Fn_Web_MenuOperation("Select",strMenu)
				End If
				'Setting Vendor Role
				If dicVendor("VendorRole") <> "" Then
					Call Fn_Web_UI_Button_Click("Fn_WebMyTc_VendorOperations",ObjAddVendorRole,"VendorRole")
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",dicVendor("VendorRole"))
					Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click ,,micLeftBtn
					wait(2)					
					'Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_VendorOperations", ObjAddVendorRole, "VendorRole", dicVendor("VendorRole"))
				End If
				'Setting Vendor Status
				If dicVendor("VendorStatus") <> "" Then
					For iCounter=0 to 5
						If ObjAddVendorRole.WebEdit("VendorStatus").Exist Then
							Exit For
						Else
							wait(10)
						End If
					Next
					Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_VendorOperations", ObjAddVendorRole, "VendorStatus", dicVendor("VendorStatus"))
				End If	
				'Setting Certification Status
				If dicVendor("CertificationStatus") <> "" Then
					Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_VendorOperations", ObjAddVendorRole, "CertificationStatus", dicVendor("CertificationStatus"))
				End If	
	Case "RemoveVendorRole"
				'Checking Existance Of "Remove Vendor Role" Dialog
				If Not ObjAddVendorRole.Exist(5) Then
					strWEBMenuPath=Fn_LogUtil_GetXMLPath("WebMyTc_Menu")
					strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "RemoveVendorRole")
					Call Fn_Web_MenuOperation("Select",strMenu)
				End If
				'Setting Vendor Role
				If dicVendor("VendorRole") <> "" Then
					Call Fn_Web_UI_Button_Click("Fn_WebMyTc_VendorOperations",ObjAddVendorRole,"VendorRole")
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",dicVendor("VendorRole"))
					Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
					wait(2)					
				End If
	End Select
	'Clicking On OKButton
	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_VendorOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"OK")

	For iCounter=0 to 5
		If ObjAddVendorRole.Exist Then
			wait(10)
		Else
			Exit For
		End If
	Next
	'Function Returns True
	Fn_WebMyTc_VendorOperations=True
	'Releasing Object
	Set ObjAddVendorRole = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_Web_ImportSpecification
'@@
'@@    Description				:	Function to import specification from Web client.
'@@
'@@    Return Value		   	    : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client	and folder should be selected.
'@@
'@@    Examples					:	Call  Fn_Web_ImportSpecification("ImportSpecification", "c:\op.txt", "MArketing Brief", "Custom Note")
'@@												
'@@	   History					:	
'@@					Developer Name				Date		    Rev. No.	 Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Koustubh Watwe			  25-Nov-2011		1.0				Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function Fn_Web_ImportSpecification(sAction, sFileName, sSpecificationType, sImportAs)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ImportSpecification"
	Dim objImpSpec, objWebChild, iCnt, sMenu, bFlag, sTitle, iCreationTime
	Fn_Web_ImportSpecification = False
	Set objImpSpec = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("InportSpecification")
	If objImpSpec.Exist(10) = false Then
		'performing menu operation
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WebMyTc_Menu"), "ToolsImportSpecification")
		If sMenu <> "False" Then
			bFlag = Fn_Web_MenuOperation("Select", sMenu)
			If NOT(bFlag) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_AddComponent : Failed to perform menu operation [ " & sMenu & " ].")
				Exit function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_AddComponent : Successfully performed menu operation [ " & sMenu & " ].")
			End If
		else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_AddComponent : Failed to get menu from XML.")
				Exit function
		End If
	End If
	Select Case sAction
		Case "ImportSpecification"
				For iCnt = 0 to cInt(objImpSpec.RowCount)
						Select case trim(lcase(objImpSpec.GetCellData(iCnt, 1)))
							' - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - - 
							Case "filename"
								If sFileName <> ""  Then		
									Set objWebChild = objImpSpec.ChildItem(iCnt,2,"WebButton",0)
									If objWebChild.exist(5) Then
										objWebChild.Click 1,1,micleftBtn
									End If
									If JavaDialog("UploadFile").exist(10) then
										JavaDialog("UploadFile").JavaEdit("FileName").Set sFileName
										JavaDialog("UploadFile").JavaButton("Open").Click
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_Web_ImportSpecification : Successfully opend file [ " & sFileName & " ] ")
									else
										' fail to identify open dialog
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Web_ImportSpecification : Failed to open file [ " & sFileName & " ] ")
											Set objImpSpec = nothing
											Set objWebChild = nothing
											Exit function
									End If
								End If
								'Close the 2nd tab - Added by Sunny R
'								wait(10)
'								If Browser("TeamcenterWeb").GetROProperty("number of tabs") > 1 Then
'									sTitle = Browser("TeamcenterWeb").GetTOProperty("title")
'									iCreationTime = Browser("TeamcenterWeb").GetTOProperty("CreationTime")
'									Browser("TeamcenterWeb").SetTOProperty "title",".*"
'									Browser("TeamcenterWeb").SetTOProperty "CreationTime","1"
'									Browser("TeamcenterWeb").Close
'									Browser("TeamcenterWeb").SetTOProperty "title",sTitle
'									Browser("TeamcenterWeb").SetTOProperty "CreationTime",iCreationTime
'								End If

							'sSpecificationType
							Case "specification type"
									If sSpecificationType <> ""  Then		
										Set objWebChild = objImpSpec.ChildItem(iCnt,2,"WebEdit",0)
										If objWebChild.exist(5) Then
											objWebChild.Set sSpecificationType
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_Web_ImportSpecification : Successfully set Specification Type [ " & sSpecificationType & " ] ")
										End If
									End If

							'sImportAs
							Case "import as"
								If sImportAs <> ""  Then		
									Set objWebChild = objImpSpec.ChildItem(iCnt,2,"WebEdit",0)
									If objWebChild.exist(5) Then
										objWebChild.Set sImportAs
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_Web_ImportSpecification : Successfully set Import As [ " & sImportAs & " ] ")
									End If
								End If
						End Select
				Next
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("OK").Click
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_Web_ImportSpecification : Successfully clicked on [ " & "OK" & " ] ")
				' TC112-2015070100-21_07_2015-Porting-VivekA-Changed as per design change
				wait(10)
				If Browser("TeamcenterWeb").GetROProperty("number of tabs") > 1 Then
					sTitle = Browser("TeamcenterWeb").GetTOProperty("title")
					iCreationTime = Browser("TeamcenterWeb").GetTOProperty("CreationTime")
					Browser("TeamcenterWeb").SetTOProperty "title",".*"
					Browser("TeamcenterWeb").SetTOProperty "CreationTime","1"
					Browser("TeamcenterWeb").Close
					Browser("TeamcenterWeb").SetTOProperty "title",sTitle
					Browser("TeamcenterWeb").SetTOProperty "CreationTime",iCreationTime
				End If

				Fn_Web_ImportSpecification = True
		' - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - - 
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_Web_ImportSpecification : Invalid case [ " & sAction & " ] ")
		' - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - -  - - - - - 
	End Select
	If Fn_Web_ImportSpecification = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_Web_ImportSpecification : Executed successfully with Case [ " & sAction & " ] ")
	End If
	Set objImpSpec = nothing
	Set objWebChild = nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebMyTc_ManageGlobalAlternateOperations
'@@
'@@    Description				 :	Function perform operations on Global Alternates 
'@@
'@@    Parameters			   :	1. sAction : Action to be performed
'@@									2. sOpenDialogBy : Open dialog by method Menu / BOMNodeIcon
'@@									3. sModuleName : Name of the perspective MyTeamcenter / StructureManager
'@@									4. sPath : Nav tree node Path / BOM Table node path
'@@									5. sAlternateItem : Global Alternate Item name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	My Teamcenter / Structure Manager perspective shuld be activated.
'@@
'@@    Examples					:	
'@@									Call Fn_WebMyTc_ManageGlobalAlternateOperations("VerifyCheckMark", "Menu", "MyTeamcenter", "000186/A;1-Assm1 (View):000187/A;1-Comp1", "000151-SportsBody")
'@@									Call Fn_WebMyTc_ManageGlobalAlternateOperations("Verify", "Menu", "StructureManager", "000186/A;1-Assm1 (View):000187/A;1-Comp1", "000151-CoupeBody")
'@@									Call Fn_WebMyTc_ManageGlobalAlternateOperations("Remove", "BOMNodeIcon", "StructureManager", "000186/A;1-Assm1 (View):000187/A;1-Comp1", "000150-CoupeBody")
'@@									Call Fn_WebMyTc_ManageGlobalAlternateOperations("Prefer", "BOMNodeIcon", "StructureManager", "000186/A;1-Assm1 (View):000187/A;1-Comp1", "000151-SportsBody")
'@@	   History					 	:	
'@@				Developer Name					Date			Rev. No.		Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe				04-Dec-2011		 		1.0			Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_WebMyTc_ManageGlobalAlternateOperations(sAction, sOpenDialogBy, sModuleName, sPath, sAlternateItem)
	GBL_FAILED_FUNCTION_NAME="Fn_WebMyTc_ManageGlobalAlternateOperations"
	Dim objMngGlotAlt, iCnt, bReturn, sMenu
	Dim objDialog, iRowCnt, objImg, arrGlobAltItems
	Dim iCount, bFlag
	bFlag = False
	Fn_WebMyTc_ManageGlobalAlternateOperations = False

	Set objMngGlotAlt = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ListGlobalAlternates")
	' checking existence of dialog
	If objMngGlotAlt.Exist(5) = False Then
		Select Case sModuleName
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			' opening from MyTeamcenter module
			Case "MyTeamcenter"
				' selecting item from NavTree
				If sPath <> "" Then
					bReturn =  Fn_Web_NavTreeOperation("Select",sPath)
					If NOT(bReturn) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebMyTc_ManageGlobalAlternateOperations : Failed to select node [ " & sPath & " ].")
						Exit function
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebMyTc_ManageGlobalAlternateOperations : Successfully select node [ " & sPath & " ].")
				End If

				Select Case sOpenDialogBy
					' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					' opening by performing menu operation
					Case "Menu", ""
							sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_Menu"), "ToolsGlobalAlternates")
							If sMenu = "False" Then		
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebMyTc_ManageGlobalAlternateOperations : Failed to get menu from XML.")
								Exit function
							End If
					
							'performing menu operation
							bReturn = Fn_Web_MenuOperation("Select", sMenu)
							If NOT(bReturn) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebMyTc_ManageGlobalAlternateOperations : Failed to perform menu operation [ " & sMenu & " ].")
								Exit function
							End If
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebMyTc_ManageGlobalAlternateOperations : Successfully performed menu operation [ " & sMenu & " ].")
				End Select
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			' opening from Structure Manager module
			Case "StructureManager"
					Select Case sOpenDialogBy
						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						' opening by performing menu operation
						Case "Menu", ""
								' selecting BOM node.
								If sPath <> "" Then
									bReturn =  Fn_WebPSE_BOMTableOperations("Select",sPath,"","")
									If NOT(bReturn) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebMyTc_ManageGlobalAlternateOperations : Failed to select node [ " & sPath & " ].")
										Exit function
									End If
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebMyTc_ManageGlobalAlternateOperations : Successfully select node [ " & sPath & " ].")
								End If
								sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "EditRemovePreferGlobalAlternate")
								If sMenu = "False" Then		
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebMyTc_ManageGlobalAlternateOperations : Failed to get menu from XML.")
									Exit function
								End If
						
								'performing menu operation
								bReturn = Fn_Web_MenuOperation("Select", sMenu)
								If NOT(bReturn) Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebMyTc_ManageGlobalAlternateOperations : Failed to perform menu operation [ " & sMenu & " ].")
									Exit function
								End If
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebMyTc_ManageGlobalAlternateOperations : Successfully performed menu operation [ " & sMenu & " ].")
						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "BOMNodeIcon"
								Set objDialog = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("BOMTable")
								iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sPath, "")
								If iRowCnt <> -1 Then
										Set objImg = objDialog.ChildItem(iRowCnt, 2, "Image", 1)
										If TypeName(objImg) <> "Nothing" Then
											If objImg.GetROProperty("file name") = "global_alternate_16.png" Then
													objImg.Click 1,1, micLeftBtn
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebMyTc_ManageGlobalAlternateOperations : Successfully clicked on [ Global Alternate ] icon of node [ " & sPath & " ].")
											End If
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebMyTc_ManageGlobalAlternateOperations : Failed to find [ Global Alternate ] icon for node [ " & sPath & " ].")
											Exit function
										End IF
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebMyTc_ManageGlobalAlternateOperations : Failed to find node [ " & sPath & " ].")
										Exit function
								End IF
				End Select
		End Select
	End If

	' checking existence of dialog.
	If objMngGlotAlt.Exist(5) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebMyTc_ManageGlobalAlternateOperations : Failed to open [ Manage Global Alternates ] Window.")
		Set objMngGlotAlt = Nothing
		Exit Function
	End If

	iRowCnt = cInt(objMngGlotAlt.RowCount)

	Select Case sAction
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Remove", "Prefer"
						'selecting alternate
						For iCnt = 1 to iRowCnt 
							If trim(objMngGlotAlt.GetCellData(iCnt, 1)) = sAlternateItem then
								objMngGlotAlt.Object.rows(iCnt - 1).click 1,1,"LEFT"
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebMyTc_ManageGlobalAlternateOperations : Successfully selected [ " & sAlternateItem  & " ].")
								bFlag = True
								Exit for
							End If
						Next
						Fn_WebMyTc_ManageGlobalAlternateOperations = 	bFlag
						If bFlag Then
							If sAction = "Prefer" Then
								' clicking on prefer button
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("SetPreferred").Click 1,1,micLeftBtn
							Else
								'clicking on remove button
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("Remove").Click 1,1,micLeftBtn
								'handling delete confirmation box
								If Browser("TeamcenterWeb").Dialog("Dialog").Exist(5) then
									Browser("TeamcenterWeb").Dialog("Dialog").WinButton("OK").Click 1,1,micLeftBtn
								End If
							End If
						End If
						' closing dialog
						Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("Close").Click 1,1,micLeftBtn
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Verify", "VerifyCheckMark"
					arrGlobAltItems = split(sAlternateItem,"~")
					For iCount = 0 to UBound(arrGlobAltItems)
						bFlag = False
						For iCnt = 1 to iRowCnt 
							If trim(objMngGlotAlt.GetCellData(iCnt, 1)) = arrGlobAltItems(iCount) then
								bFlag = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebMyTc_ManageGlobalAlternateOperations : Successfully verified existence of [ " & arrGlobAltItems(iCount)  & " ].")
								If sAction = "VerifyCheckMark" Then
									bFlag = False
									Set objImg = objMngGlotAlt.ChildItem(iCnt, 3, "Image", 0)
									If TypeName(objImg) <> "Nothing" Then
										If objImg.GetROProperty("file name") = "checkmark.png" Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebMyTc_ManageGlobalAlternateOperations : Successfully verified check mark against [ " & arrGlobAltItems(iCount)  & " ].")
											bFlag = True
										End If
									End IF
								End If
								Exit for
							End If
						Next
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebMyTc_ManageGlobalAlternateOperations : [ " & arrGlobAltItems(iCount)  & " ] does not exists.")
							Exit for
						End If
					Next
					' closing dialog
					Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("Close").Click 1,1,micLeftBtn
					Fn_WebMyTc_ManageGlobalAlternateOperations =  bFlag
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebMyTc_ManageGlobalAlternateOperations : Invalid case [ " & sAction  & " ].")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		End Select

	If Fn_WebMyTc_ManageGlobalAlternateOperations Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebMyTc_ManageGlobalAlternateOperations : executed successfully with case [ " & sAction  & " ].")
	End If
	Set objMngGlotAlt = Nothing
	Set objImg = Nothing
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebMyTc_CreateVendorPartSearchReults
'@@
'@@    Description				 :	Function Used To Select Vendor From Vendor Part Search Reults
'@@
'@@    Parameters			   :	1.sAction :Action To Perform
'@@											2. dicVendorPart: Vendor Part Full Information Dictionary Object
'@@											3. sButton :Button Name To Click On
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	
'@@												dicVendor("Name")="000016-VendorPart~000016-Vendor"
'@@												Call Fn_WebMyTc_CreateVendorPartSearchReults("Select", dicVendor, "OK")
'@@												
'@@	   History					 	:	
'@@													Developer Name						Date				Rev. No.						Changes Done			Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Pranav Ingle							19-Dec-2013				1.0																	Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_WebMyTc_CreateVendorPartSearchReults(sAction, dicVendor, sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_WebMyTc_CreateVendorPartSearchReults"
	'Variable Declaration
	Dim ObjVendorPart,ObjPartInfo
	Dim CurrValue,strWEBMenuPath,strMenu,dicItems
	Dim iCount, iRowCounter, arrVendorPart, iCounter, bFlag

	Fn_WebMyTc_CreateVendorPartSearchReults=False
	'Creating Objects Of Vendor Part Related Tables
	Set ObjVendorPart=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchResults")
	'Set ObjPartInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchResults") '' used this hiearachy after discussion with Akshay J.(by Ujwal N-26-Jun-2014)
	Set ObjPartInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchResults").WebTable("ResultsTable") ' used this hiearachy for Tc11.1 build0723(by Saurabh Khandekar)
	'Checking Existance of "New Vendor Part" Dialog
	If Not ObjVendorPart.Exist(5) Then
		Exit Function
	End If
	'arrVendorPart = Split(dicVendor("PartName"), "~")
	arrVendorPart = Split(dicVendor("Name"), "~")
	iRowCounter = ObjPartInfo.GetROProperty("Rows")

	Select Case sAction
		Case "Select", "Verify"
			For iCounter = 0 To UBound(arrVendorPart)
				bFlag = False
				For iCount = 3 To iRowCounter
					' Taking row Count from 3 as values starting from 3rd Row and Col Count from 2 for Name
					CurrValue = ObjPartInfo.GetCellData(iCount, 2) 
					If  CurrValue = arrVendorPart(iCounter) Then
                        bFlag = True
						If sAction = "Select" Then
							ObjPartInfo.WebRadioGroup("RadioGroup").Select "#"&iCount-3
						End If
					End If
        		Next
				If bFlag = False Then
					Exit Function
				End If
			Next
	End Select

    'Clicking On Finish Button
	Wait(1)
'	Call Fn_Web_UI_Button_Click("Fn_WebMyTc_CreateVendorPartSearchReults",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("2ndLavelButtunPanel"),sButton)
	Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("2ndLavelButtunPanel").WebButton(sButton).Click
	Wait(1)
	For iCounter=0 to 5
		If ObjVendorPart.Exist Then
			wait(10)
		Else
			Exit For
		End If
	Next
	'Function Returns True
	Fn_WebMyTc_CreateVendorPartSearchReults=True
	'Releasing Objects Of Vendor Part Related Tables
	Set ObjVendorPart=Nothing
	Set ObjPartInfo=Nothing
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebMyTc_ChangeVendorOperations
'@@
'@@    Description				 :	Function Used To Select Vendor From Vendor Part Search Reults
'@@
'@@    Parameters			   :	1.sAction :Action To Perform
'@@											2. dicVendorPart: Vendor Part Full Information Dictionary Object
'@@											3. sButton :Button Name To Click On
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	
'@@												Case "Insert"
'@@												dicVendor("New Vendor")="000016-VendorPart~000016-Vendor"
'@@												Call Fn_WebMyTc_ChangeVendorOperations("Insert", dicVendor, "OK")
'@@												
'@@												Case "Search"
'@@												dicVendor("New Vendor")="000016-VendorPart~000016-Vendor"
'@@												Call Fn_WebMyTc_ChangeVendorOperations("Insert", dicVendor, "OK")
'@@												
'@@	   History					 	:	
'@@													Developer Name						Date				Rev. No.						Changes Done			Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Pranav Ingle							27-Dec-2013				1.0																	Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_WebMyTc_ChangeVendorOperations(sAction, dicVendor, sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_WebMyTc_ChangeVendorOperations"
	'Variable Declaration
	Dim objChangeVendor, dicItem, dicKeys
	Dim CurrValue,strWEBMenuPath,strMenu,dicItems
	Dim iCount, iRowCounter, arrVendorPart, iCounter, bFlag

	Fn_WebMyTc_ChangeVendorOperations=False
	'Creating Objects Of Vendor Part Related Tables
	Set objChangeVendor=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ChangeVendor")

	'Checking Existance of "New Vendor Part" Dialog
	If Not objChangeVendor.Exist(5) Then
		Exit Function
	End If


	Select Case sAction
		Case "Insert"
			Call Fn_Web_UI_WebEdit_Set("Fn_WebMyTc_ChangeVendorOperations", objChangeVendor, "WebEdit", dicVendor("New Vendor"))
			 'Clicking On Finish Button
			Call Fn_Web_UI_Button_Click("Fn_WebMyTc_ChangeVendorOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),sButton)
			Wait(1)
		Case "Search"
			' In Progress
			'	Set dicItem = dicVendor.Items
			'	Set dicIKeys = dicVendor.Keys
			
			'	arrVendorPart = Split(dicVendor("Name"), "~")
			'	iRowCounter = objChangeVendor.GetROProperty("Rows")

	End Select

'	Wait(1)
'	For iCounter=0 to 5
'		If ObjVendorPart.Exist Then
'			wait(10)
'		Else
'			Exit For
'		End If
'	Next
	'Function Returns True
	Fn_WebMyTc_ChangeVendorOperations=True
	'Releasing Objects Of Vendor Part Related Tables
	Set objChangeVendor=Nothing
End Function
