Option Explicit

'######################################################################################################################################################################################################################################################
'#	Function Library Name						:	 SearchDictionary.vbs
'#
'#
'#	Function Teamcenter Action Association		:	Teamcenter Search
'#
'#		Teamcenter UI State Pre Condition		:	My Teamcenter Application - [Search Criteria] Tab loaded
'#
'#		Teamcenter UI State Post Condition		:	Preferred action performed under [Search Criteria] view in My Teamcenter Context
'#
'#
'#	Function UI Control Types Exercised			:	
'#
'#		Java									:	Java Edit, Java Date, Java button, Java Combo
'#		Eclipse									:	-None-
'#		Web										:	-None-
'#		Windows									:	-None-
'#
'#
'#
'#	Function Logical Implementation Description	:	Function is designed to store test arguements for search criteria which are to be exercised
'#													under My Teamcenter Application Context
'#
'#
'#  	Function Parameter Details				:	Dictionary Key Value Pair
'#
'#
'#
'#	Function Return Value/Type					:	
'#	   General Case								:	Boolean
'#	   Specific (None)							:	
'#
'#
'#
'#	Function Dependancy Matrix					:	
'#		Parent Functions						:	
'#		Child Functions							:
'#
'#
'#
'#	Function Unit Test And Publication Tc Build	:	Teamcenter 9.0.20101103
'#
'#
'#	Function Usage Example						:
'#
'#
'#	Function Change History						:
'#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	Created By					Date			Change Version		Function Change Unit Tests TcBuild (Build ID)			Change Review/Approval			Change Description
'#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	Mallikarjun					11/21/2010		0.1			Teamcenter 9.0 (20101103)					Mallikarjun				New Creation and Publication
'######################################################################################################################################################################################################################################################
'********************************************************************************************
' Search Dictionary Definitions:
'********************************************************************************************

' Defining Dictionary For Specific Search Criteria:
Dim dicSearchCriteria

'Instanciating Search Dictionary Object
Set dicSearchCriteria = CreateObject( "Scripting.Dictionary" )

'Declaration Search Dictionary Object Structure:
With dicSearchCriteria  
'Java Edit Box
			.Add "SrchPersonName", ""
			.Add "AllSeqOwningUser", ""
			.Add "CurrentTask",""	
			.Add "DatasetID", ""	                
			.Add "Description", ""	
			 .Add "DatasetType", ""	
			.Add "ItmItemID", ""	 
			.Add "ItmName", ""	 
			.Add "ItmOwningOrgID", ""	 
			.Add "ItmRevAliasID", ""	 
            .Add "ItmRevAliasIDCntxtNm", ""	
			.Add "ItmRevAltID", ""	 
			.Add "ItmRevAltRev", ""
			.Add "ItmRevAltRevIDCntxtNm", ""	 
			.Add "ItmRevCurrentTask", ""	 		
			.Add "ItmRevDesc", ""  
			.Add "ItmRevItemID", ""
			.Add "ItmRevName", ""
			.Add "ItmRevRelStatus", ""
			.Add "ItmRevRevision", ""
			.Add "Keywords", ""
			.Add "Name", ""
			.Add "ReleaseStatus", ""
			.Add "Revision", ""
			.Add "PersonNm", ""
			.Add "ObjectID", ""
			.Add "ProjectID", ""
			.Add "UserData1", ""
			.Add "UserData2", ""
			.Add "ExcludeStatus", ""
            .Add "project_name", ""
'Java Check Box ->> Date Button
			.Add "CreatedAfterDt",""	
			.Add "CreatedBeforeDt",""	
			.Add "ModifiedAfterDt",""	
			.Add "ModifiedBeforeDt",""	
			.Add "ReleasedAfterDt",""	
			.Add "ReleasedBeforeDt",""	
' Java Button(DropDown)
			.Add "DsOwnGrpDrpDwn",""	
			.Add "DsOwnUsrDrpDwn",""	
			.Add "DsTypDrpDwn", ""	
			.Add "ItmRevAliasTypDrpDwn", ""	
			.Add "ItmRevAltRevTypDrpDwn", ""
			.Add "ItmRevOwningGrpDrpDwn",""	
			.Add "ItmRevOwningUsrDrpDwn",""	
			.Add "ItmRevTypDrpDwn", ""	
			.Add "OwningGrpDrpDwn", ""	
			.Add "OwningUsrDrpDwn", ""
			.Add "GenOwningUsrDrpDwn", ""			
			.Add "GenOwningGrpDrpDwn", ""
			.Add "DsKwSrchDatasetTypeDrpDwn", ""
			.Add "DsKwSrchOwnGrpDrpDwn", ""
			.Add "DsKwSrchOwnUsrDrpDwn", ""
			.Add "UsrID", ""
			.Add "Requestor", ""
End with
