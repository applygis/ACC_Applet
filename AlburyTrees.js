/**************************************

Client: Albury City Council
Project: Albury Trees
Author: Adam Goodfellowne

Copyright: RIA Mobile GIS 2008

**************************************/

/**************************************
Constants
**************************************/
function FormMode(){}
FormMode.Closed = -1;
FormMode.Unbound = 0;
FormMode.Identify = 1;
FormMode.Edit = 2;
FormMode.Create = 3;

function RecordsetMode(){}
RecordsetMode.Read = 1;
RecordsetMode.Write = 2;

function FileFormat(){}
FileFormat.Default = -2;
FileFormat.Unicode = -1;
FileFormat.ASCII = 0;

function FileMode(){}
FileMode.Read = 1;
FileMode.Write = 2;
FileMode.Append = 8;

function ShapeType(){}
ShapeType.Null = 0;
ShapeType.Point = 1;
ShapeType.Line = 3;
ShapeType.Polygon = 5;

function FieldType(){}
FieldType.Numeric = 5;
FieldType.Date = 9;
FieldType.Boolean = 11;
FieldType.Char = 129;

function MessageBoxType(){}
MessageBoxType.YesNo = 4;
MessageBoxType.Critical = 16;
MessageBoxType.Question = 32;
MessageBoxType.Exclamation = 48;
MessageBoxType.Information = 64;

function MessageBoxResponse(){}
MessageBoxResponse.OK = 1;
MessageBoxResponse.Cancel = 2;
MessageBoxResponse.Abort = 3;
MessageBoxResponse.Retry = 4;
MessageBoxResponse.Ignore = 5;
MessageBoxResponse.Yes = 6;
MessageBoxResponse.No = 7;

var APPLET_NAME = "RIA_Applet";
var TOOLBAR_NAME = "RiaToolbar01";

/**************************************
Glogal Vars
**************************************/
var g_sDataPath = Preferences.Properties( "DataPath" );
var g_sAppletPath = Preferences.Properties( "AppletsPath" );
var g_sLayerName = "AlburyTrees";
var g_sParcelLayerName = "parcels";
var g_sRoadLayerName = "RoadsCentreline";
var g_sZLLayerName = "City_Tree_Zones";
var g_sAXFFileName;
var g_sDBFFileName;
var g_sLogFileName = "\\AlburyTrees.log";
var g_sRisk_Fail_RS_Location = "\\LookUpTables\\RISK_FAI.DBF";
var g_sRisk_Prob_RS_Location = "\\LookUpTables\\RISK_PRO.DBF";
var g_sRisk_Targ_RS_Location = "\\LookUpTables\\RISK_TAR.DBF";
var g_sZone_LUT_Location = "\\LookUpTables\\LUT_ZONE.DBF";
var g_sLedg_LUT_Location = "\\LookUpTables\\LUT_LEDG.DBF";
var g_sLUT_Works = "\\LookUpTables\\LUT_WORK.DBF";
var g_sLUT_Works_Datafiles = "LUT_WORK.DBF";
var g_sLUT_Opp = "\\LookUpTables\\LUT_OPER.DBF";
var g_sLUT_Opp_Datafiles = "LUT_OPER.DBF";
var g_sLUT_Nature = "\\LookUpTables\\LUT_Natu.DBF";
var g_sLUT_Works = "\\LookUpTables\\LUTWORKC.DBF";
var g_sLUT_GS = "\\LookUpTables\\GenusSpecies.DBF";
var g_sLUT_Defects = "\\LookUpTables\\LUTDEFEC.DBF";
var g_sLUT_Power = "\\LookUpTables\\POWERLIN.DBF";

var g_oRiskFailRS = Application.CreateAppObject("RecordSet");
var g_oRiskProbRS = Application.CreateAppObject("RecordSet");
var g_oRiskTargRS = Application.CreateAppObject("RecordSet");

var g_bRiskFailisOpen = false;
var g_bRiskProbisOpen = false;
var g_bRiskTargisOpen = false;

var g_oCurrentPoint= Application.CreateAppObject("point");
var g_bAuditSession = false;
var g_bChangeAudit = false;
var g_bNewPlantingSession = false;
var g_bELBTools = false;
var g_iChangessToSave = 0;
var g_lAssetID = 0;
var g_iBotIndex = 0;

var g_sStreetP;
var g_sStreetP_1;
var g_sStreetN;
var g_sStreetNum;
var g_sSub;
var g_sZL;
var g_bLoading = false;
var g_Vacant;
var g_strTreeIma_Name = "";
var g_Redundant = false;
var g_LoadedStatusValue;
var g_MadeTreeCurrent = false;

var g_fp;
var g_fs;
var g_ChangedToVacant = false;

var g_intAssetID;
		

Application.UserProperties("MaintFormClose") = false;

/**************************************
AppStartup
**************************************/
function appStartup() {
	var oFile = Application.CreateAppObject("file");
	if ( !oFile.Exists( g_sDataPath + "\\" + g_sLUT_Opp_Datafiles) ){
		oFile.Copy(g_sAppletPath + "\\" + g_sLUT_Opp, g_sDataPath + "\\" + g_sLUT_Opp_Datafiles);
	}
	if ( !oFile.Exists( g_sDataPath + "\\" + g_sLUT_Works_Datafiles) ){
		oFile.Copy(g_sAppletPath + "\\" + g_sLUT_Works, g_sDataPath + "\\" + g_sLUT_Works_Datafiles);
	}
	
}

function btnNewTree_OnPointerDown( oButton ){
	Application.UserProperties("MaintFormClose")  = true;
	
	g_oCurrentPoint.X = Map.PointerX;
	g_oCurrentPoint.Y = Map.PointerY;

	var dt = new Date();
	//g_intAssetID = parseInt (pad20( dt.getHours().toString() )+ pad20(dt.getMinutes().toString() )+ pad20(dt.getYear().toString())+ pad20(dt.getDate().toString()) + pad20((dt.getMonth() + 1).toString()));  
	g_intAssetID = parseInt (pad20(dt.getYear().toString())+ pad20(dt.getDate().toString()) + pad20((dt.getMonth() + 1).toString()) + pad20( dt.getHours().toString())+ pad20(dt.getMinutes().toString()));  
	
	var ods = Map.Layers(g_sLayerName).DataSource;

	if (ods.isOpen){
		Console.print (g_intAssetID);
		if (g_intAssetID >= 2147483647){ //2336140810
			Console.print ("need a new id");
			return;
		}
		var insertSQL = "INSERT INTO [ALBURYTREES] (ASSET_ID, AXF_STATUS, SHAPE_X, SHAPE_Y) VALUES (" + g_intAssetID + ", 1, " + Map.PointerX + ", " + Map.PointerY + ");" 
		Console.print (insertSQL);
		ods.Execute( insertSQL );
	}

	Map.Refresh();
	ods.Close();

	if ( !getBackgroundData( g_oCurrentPoint.X, g_oCurrentPoint.Y ) ){
		return;
	}
	processZL( g_oCurrentPoint.X, g_oCurrentPoint.Y );
	Application.Applets("AlburyTrees").Forms("NewTree").Show();	
}

function btnNewBeetle_OnPointerDown( oButton ){
	Application.UserProperties("MaintFormClose")  = true;

	if(!Map.EditLayer){
		Map.Layers("ALBURYTREES").Editable = true;
	}
	var result = Map.SelectXY(Map.PointerX, Map.PointerY);
	if (result = 'True'){
		var objLayer = Map.EditLayer;
		var objRS = objLayer.Records;
		objRS.BookMark
		objRS.BookMark = Map.SelectionBookMark;

		g_lAssetID = objRS.fields("ASSET_ID").Value;
		Application.Applets("AlburyTrees").Forms("NewBeetle").Show();
	}
}

function btnNewRootPrune_OnPointerDown( oButton ){
	Application.UserProperties("MaintFormClose")  = true;
	if(!Map.EditLayer){
		Map.Layers("ALBURYTREES").Editable = true;
	}
	Map.SelectXY(Map.PointerX, Map.PointerY);
	var objLayer = Map.EditLayer;
	var objRS = objLayer.Records;

	objRS.BookMark = Map.SelectionBookMark;

	g_lAssetID = objRS.fields("ASSET_ID").Value;
	Application.Applets("AlburyTrees").Forms("NewRootPrune").Show();
}

function btnBeetleWorks_OnPointerDown( oButton ){
	Application.UserProperties("MaintFormClose")  = true;
	if(!Map.EditLayer){
		Map.Layers("ALBURYTREES").Editable = true;
	}
	var result = Map.SelectXY(Map.PointerX, Map.PointerY);
	//Console.print ("result: " + result);
	var objLayer = Map.EditLayer;
	var objRS = objLayer.Records;

	objRS.BookMark = Map.SelectionBookMark;

	g_lAssetID = objRS.fields("ASSET_ID").Value;
	Application.Applets("AlburyTrees").Forms("BeetleWorks").Show();
}

function btnRootPruneWorks_OnPointerDown( oButton ){
	Application.UserProperties("MaintFormClose")  = true;
	if(!Map.EditLayer){
		Map.Layers("ALBURYTREES").Editable = true;
	}
	Map.SelectXY(Map.PointerX, Map.PointerY);
	var objLayer = Map.EditLayer;
	var objRS = objLayer.Records;

	objRS.BookMark = Map.SelectionBookMark;

	g_lAssetID = objRS.fields("ASSET_ID").Value;
	Application.Applets("AlburyTrees").Forms("RootPruneWorks").Show();
}

function btn_GPSNew( oToolButton ){
	if ( GPS.IsValidFix ){
		g_oCurrentPoint.X = GPS.X;
		g_oCurrentPoint.Y = GPS.Y;

	}else{
		MessageBox("Not a valid GPS Fix, unable to use GPS." );
		return;
	}
	
	if ( !getBackgroundData( g_oCurrentPoint.X, g_oCurrentPoint.Y ) ){
		return;
	}
	processZL( g_oCurrentPoint.X, g_oCurrentPoint.Y );
	Application.Applets("AlburyTrees").Forms("NewTree").Show();
}
function getBackgroundData(dX, dY){
	//Get Layers and record sets

	var oRoadRS, oParcelRS;
	var oRoadLayer, oParcelLayer;
	oRoadLayer = Map.Layers( g_sRoadLayerName );
	oParcelLayer = Map.Layers( g_sParcelLayerName );

	if ( oRoadLayer == null ){
		MessageBox( "Unable to locate road layer. Check that there is a road centerline layer loaded called: " + g_sRoadLayerName + ".");
		return false;		
	}
	if ( oParcelLayer == null ){
		MessageBox( "Unable to locate parcel layer. Check that there is a parcel layer loaded called: " + g_sParcelLayerName + "." );
		return false;
	}

	oRoadRS = oRoadLayer.Records;
	oParcelRS = oParcelLayer.Records;

	if ( oRoadRS == null || oParcelRS == null ){
		return false;
	}


	//Roads
	//Now we know we have to 2 record sets continue
	var dRoadBK = oRoadRS.FindNearestXY( dX, dY, 30 );

	if ( dRoadBK > 0 ){
		oRoadRS.Bookmark = dRoadBK;
		g_sStreetP = oRoadRS.Fields("ROAD_NAME").Value;
	}else{
		g_sStreetP = "";
	}
	//Parcels
	var dParcelBK = oParcelRS.FindNearestXY( dX, dY, 30 );

	if ( dParcelBK > 0 ){
		oParcelRS.Bookmark = dParcelBK;
		g_sStreetN = oParcelRS.Fields("STREET_NAM").Value;
		
		g_sStreetNum = oParcelRS.Fields("HOUSE_NUMB").Value;

		//g_sStreetP_1 = g_sStreetNum + " " + g_sStreetP;

	}else{
		g_sStreetN = "";
		g_sStreetNum = 0;
	}
	return true;
}
function processZL( dX, dY ){
	//Zone Ledger
//Console.print ("processZL");
	var oZLLayer,oZLRS;
	oZLLayer = Map.Layers( g_sZLLayerName );
	if ( oZLLayer == null ){
		g_sZL = "";
		return;
	}
	oZLRS = oZLLayer.Records;
	var dZLBM = oZLRS.FindNearestXY( dX, dY )
	if ( dZLBM > 0 ){
		oZLRS.Bookmark = dZLBM;
		g_sZL = oZLRS.Fields( "Zone_ID" ).Value + ":6.0" + (pad20(oZLRS.Fields( "Code" ).Value)).toString() ;
		g_sStreetP_1 = g_sStreetP + " " + oZLRS.Fields( "Zone_ID" ).Value

		if (g_bELBTools){
			g_lAssetID = oZLRS.Fields( "Asset_ID" ).Value;
		}
	}else{
		g_sZL = "";
		g_lAssetID = "";
	}
	Application.UserProperties( "ZL" ) = g_sZL;	
}
function btn_AuditTree_onPointerDown( oButton ){
	g_bAuditSession = true;
	Application.UserProperties("NewPlanting") =  false;
	Application.UserProperties("MaintFormClose") = false;
	g_bNewPlantingSession = false;

	var oLayer = Map.Layers( g_sLayerName );
	var oForm = oLayer.Forms("EDITFORM");
	if ( oLayer == null ){
		MessageBox( "Unable to locate " + g_sLayerName + " layer. Please enuse layer is loaded with the corect name and try again." );
		return;
	}
	oLayer.Editable  = true;
	Application.ExecuteCommand ( "modeselect" );
}

function btn_GPSAudit( oToolButton ){

}

function btn_ChangeDetails_onPointerDown( oButton ){
	g_bChangeAudit = true;
	var oLayer = Map.Layers( g_sLayerName );
	if ( oLayer == null ){
		MessageBox( "Unable to locate " + g_sLayerName + " layer. Please enuse layer is loaded with the corect name and try again." );
		return;
	}
	oLayer.Editable  = true;
	Application.ExecuteCommand ( "modeselect" );
}

function btn_NewPlanting_onPointerDown( oButton ){

	if ( ! Map.SelectXY(Map.PointerX, Map.PointerY) ){
		return;
	}
	
	Application.UserProperties("MaintFormClose") = false;
	g_bAuditSession = false;
	Application.UserProperties("NewPlanting") = true;
	g_bNewPlantingSession = true;
	var oLayer = Map.Layers( g_sLayerName );
	if ( oLayer == null ){
		MessageBox( "Unable to locate " + g_sLayerName + " layer. Please enuse layer is loaded with the corect name and try again." );
		return;
	}
	oLayer.Editable  = true;
	Application.ExecuteCommand ( "modeselect" );
}

function NewTree_onLoad( oForm ){

	WaitCursor ( 1 );
	g_bLoading = true;
	resetForm( oForm );
	setFileNameAXF();

	OpenRiskRS();
	var oDS = OpenAXF(g_sAXFFileName);
	var oRS;
	fso = Application.CreateAppObject("file");

	if ( !fso.Exists( g_sAXFFileName ) ){
		Application.MessageBox ( "Required File not Found: " + g_sAXFFileName );
		oDS.Close();
		oForm.Close();
		return;
	}
	if ( !fso.Exists( g_sAppletPath + g_sLUT_Works ) ){
		Application.MessageBox ( "Required File not Found: " + g_sAppletPath + g_sLUT_Works );
		oDS.Close();
		oForm.Close();
		return;
	}

	var oPage1C = oForm.Pages("PAGE1").Controls;
	var oPage2C = oForm.Pages("PAGE2").Controls;
	var oPage3C = oForm.Pages("PAGE3").Controls;
	var oPage4C = oForm.Pages("PAGE4").Controls;
	var oPage5C = oForm.Pages("PAGE5").Controls;

	//Load the Combo Box with values Page 1
/*	oRS = oDS.Execute ( "SELECT Code as Value, Description as Text FROM [CVD_AUDITS_ROAD_NAME] where is_hidden = 0 order by Text;" );
	LoadCombobox( oRS, ( oPage1C( "cb_streetPlanted" ) ) );
	oRS = oDS.Execute ( "SELECT Code as Value, Description as Text FROM [CVD_AUDITS_ST_NAME] where is_hidden = 0 order by Text;" );
	LoadCombobox( oRS, ( oPage1C( "cb_StreetN" ) ) );
*/
	//Page 1
	//genus_OnLoad( oPage1C ); 

	//Page 2
	//Tree Age
	oRS = oDS.Execute ( "SELECT Code as Value, Description as Text FROM [CVD_AUDITS_LUT_AGE] where is_hidden = 0;" );
	LoadCombobox( oRS, ( oPage2C( "cbx_TA" ) ) );
	//Tree Health
	oRS = oDS.Execute ( "SELECT Code as Value, Description as Text FROM [CVD_AUDITS_LUT_HEALTH] where is_hidden = 0;" );
	LoadCombobox( oRS, ( oPage2C( "cbx_TH" ) ) );
	//ULE
	oRS = oDS.Execute ( "SELECT Description as Value,Description as Text FROM [CVD_AUDITS_LUT_ULE] where is_hidden = 0 order by code;" );
	LoadCombobox( oRS, ( oPage2C( "cbx_ule" ) ) );
	//Tree Structure
	oRS = oDS.Execute ( "SELECT Code as Value, Description as Text FROM [CVD_AUDITS_LUT_STRUCTUR] where is_hidden = 0;" );
	LoadCombobox( oRS, ( oPage2C( "cbx_ts" ) ) );

	oRS = oDS.Execute ( "SELECT code as Value,description as Text FROM [CVD_AUDITS_CVD_PROBABILITYOFFAILURE] where is_hidden = 0;" );
	LoadCombobox( oRS, ( oPage2C( "cbx_fp" ) ) );
	oRS = oDS.Execute ( "SELECT Code as Value,Description as Text FROM [CVD_AUDITS_CVD_FAILURE] where is_hidden = 0;" );
	LoadCombobox( oRS, ( oPage2C( "cbx_fs" ) ) );
	oRS = oDS.Execute ( "SELECT Code as Value, Description as Text FROM [CVD_AUDITS_CVD_TARGET_RATING] where is_hidden = 0;" );
	LoadCombobox( oRS, ( oPage2C( "cbx_to" ) ) );

	oRS = oDS.Execute ( "SELECT Code as Value, Description as Text FROM [CVD_AUDITS_LUT_WORKS_PRIORITY] where is_hidden = 0;" );
	LoadCombobox( oRS, ( oPage2C( "cbx_p" ) ) );

	oRS = oDS.Execute ( "SELECT Code as Value, Description as Text FROM [CVD_AUDITS_LUT_NATURE_STRIP_WIDTH] where is_hidden = 0 order by code;" );
	LoadCombobox( oRS, ( oPage2C( "cb_NatureStrip" ) ) );

	//Generate the Temporay Asset ID from Date
	dt = new Date();
	if ( g_bChangeAudit ) {
		//Console.print ( "in form load: " + g_Vacant );
		g_LoadedStatusValue = oForm.Pages( "PAGE1" ).Controls( "cb_CURRENT_ST" ).Value

		oForm.Pages( "PAGE4" ).Activate();
	
		LoadListBoxAudits( oPage4C("lb_Audits"), g_intAssetID, oForm );
		oPage1C("cb_streetPlanted").Enabled = false;
		oPage1C("tb_HouseNum").Enabled = false;
		oPage1C("cb_StreetN").Enabled = false;
		oPage1C("cb_street_p").Enabled = false;
		
		//oPage2C( "tbx_works" ).Value = "No Works";
		oForm.Caption = "Edit Asset Information";

		showStreetText( true, oPage1C );

		//g_Vacant = oPage1C("chkVacant").Value;
		//Console.print ( "in form load 2: " + g_Vacant );

	/*	if (oPage1C( "cb_CURRENT_ST" ) == "Proposed" ){
			oPage1C("chkVacant").Enabled = true;
			oPage1C("chkVacant").Value = true;
		}
		else {
			oPage1C("chkVacant").Enabled = false;
			oPage1C("chkVacant").Value = false;
		}*/
		
		WaitCursor ( -1 );
	}else {
//		g_intAssetID = parseInt (pad20( dt.getHours().toString() )+ pad20(dt.getMinutes().toString() )+ pad20(dt.getYear().toString())+ pad20(dt.getDate().toString()) + pad20((dt.getMonth() + 1).toString()));  
		//Console.print ("on load: " + g_intAssetID );
		oForm.Pages( "PAGE1" ).Controls( "tbx_id" ).Value = g_intAssetID; // pad20( dt.getHours().toString() )+ pad20(dt.getMinutes().toString() )+ pad20(dt.getYear().toString())+ pad20(dt.getDate().toString()) + pad20((dt.getMonth() + 1).toString());
		oForm.Caption = "Create New Site/Asset";
		oForm.Pages( "PAGE1" ).Controls("cb_CURRENT_ST").ListIndex = 2;

			//Page3
		oForm.Pages( "PAGE3" ).Activate();
		oForm.Pages( "PAGE3" ).Controls("cb_Defects").ListIndex = 0;
		
		oForm.Pages( "PAGE1" ).Activate();
		showStreetText( false, oPage1C );
		var oParcelLayer = Layers( g_sParcelLayerName );
		if ( oParcelLayer == null ){
			MessageBox( "Unable to locate parcel layer, property details will not be populated.\n Please manually select property address" );
		}else{
			oPage1C( "cb_StreetN" ).AddItemsFromTable ( g_sDataPath + "\\" + oParcelLayer.FilePath.FileParts().Filename + ".dbf", "STREET_NAM","STREET_NAM" );
			
		}
		var oRoadLayer = Layers( g_sRoadLayerName );
		if ( oRoadLayer == null ){
			MessageBox( "Unable to locate road layer, street planted details will not be populated.\n Please manually select street planted" );
		}else{
			oPage1C( "cb_streetPlanted" ).AddItemsFromTable ( g_sDataPath + "\\" + oRoadLayer.FilePath.FileParts().Filename + ".dbf", "ROAD_NAME","ROAD_NAME" );	
		}
		
		oPage1C( "cb_Botanical" ).Clear();
		oPage1C( "cb_Botanical" ).Enabled = false;
		setCBOReg( oPage1C("cb_streetPlanted"), "/*" + removeStsuff( g_sStreetP ) + "/*" );
		setCBO( oPage1C("cb_streetPlanted"), g_sStreetP );
		setCBO( oPage1C("cb_StreetN"), g_sStreetN );
		oPage1C("tb_HouseNum").Text = g_sStreetNum;
		oPage1C("tb_HouseNum").Value = g_sStreetNum;
		oPage1C("cb_Zone").Value = g_sZL;
		setCBO( oPage2C( "cbx_works" ), "No works" );
		setTB( oPage2C( "tbx_works" ), "No works" );
		oPage3C( "dp_DateVisited" ).Value = new Date().getVarDate();

		WaitCursor ( -1 );
	}

	g_bLoading = false;

	g_strTreeIma_Name = "";

	oRS.Close();
	oDS.Close();

	WaitCursor ( -1 );
}

function checkForAudit( oPage ){ 
	var oPage1C = oPage;

	if (oPage1C("cb_CURRENT_ST").value != "Proposed"){
		oPage1C("chkVacant").value = false;
		g_Vacant = false;
		chk_Vacant_onClick(oPage1C("chkVacant"));
	}
}

function NewBeetle_onLoad( oForm ){
	
	WaitCursor ( 1 );
	g_bLoading = true;
	resetForm( oForm );
	setFileNameAXF();

	OpenRiskRS();

	var oDS = OpenAXF(g_sAXFFileName);
	var oRS;

	fso = Application.CreateAppObject("file");
	if ( !fso.Exists( g_sAXFFileName ) ){
		Application.MessageBox ( "Required File not Found: " + g_sAXFFileName );
		oDS.Close();
		oForm.Close();
		return;
	}
	if ( !fso.Exists( g_sAppletPath + g_sLUT_Works ) ){
		Application.MessageBox ( "Required File not Found: " + g_sAppletPath + g_sLUT_Works );
		oDS.Close();
		oForm.Close();
		return;
	}
	if(oForm.Name == "NewBeetle"){
		var oPage1 = oForm.Pages("tp_ELB").Controls;
	}
	else{
		var oPage1 = oForm.Pages("tp_ELBWorks").Controls;
		g_bELBTools= true;
	}		

	oRS = Application.CreateAppObject("recordset");
	oRS.Open(g_sAppletPath + g_sLUT_Opp);

	oRS.MoveFirst();
	var counter = 1
	while (counter <= oRS.RecordCount){
		oPage1("cb_inspector").AddItem ( oRS.Fields(1).Value, oRS.Fields(2).Value );
		oRS.MoveNext();
		counter++
	}

	oRS.Close();
	WaitCursor ( -1 );
}


function NewRootPrune_onLoad( oForm ){

	WaitCursor ( 1 );
	g_bLoading = true;
	resetForm( oForm );
	setFileNameAXF();

	//OpenRiskRS();

	var oDS = OpenAXF(g_sAXFFileName);
	var oRS;

	fso = Application.CreateAppObject("file");
	if ( !fso.Exists( g_sAXFFileName ) ){
		Application.MessageBox ( "Required File not Found: " + g_sAXFFileName );
		oDS.Close();
		oForm.Close();
		return;
	}
	if ( !fso.Exists( g_sAppletPath + g_sLUT_Works ) ){
		Application.MessageBox ( "Required File not Found: " + g_sAppletPath + g_sLUT_Works );
		oDS.Close();
		oForm.Close();
		return;
	}

	if(oForm.Name == "NewRootPrune"){
		var oPage1 = oForm.Pages("tp_RootPrune").Controls;
	}
	else{
		var oPage1 = oForm.Pages("tp_RootPruneWorks").Controls;
		g_bELBTools =true;
	}

	oRS = Application.CreateAppObject("recordset");
	oRS.Open(g_sAppletPath + g_sLUT_Opp);

	oRS.MoveFirst();
	//Console.print (oRS.RecordCount);
	var counter = 1
	while (counter <= oRS.RecordCount){
		//Console.Print (oRS.Fields(1).Value);
		oPage1("cb_inspector").AddItem ( oRS.Fields(1).Value, oRS.Fields(2).Value );
		oRS.MoveNext();
		counter++
	}

	oRS.Close();
	WaitCursor ( -1 );
}

function NewBeetle_onOK( oForm ){
	var oLayerRS = null
	var oLayer = Map.Layers( g_sLayerName );
	if ( oLayer == null ){
		MessageBox("Error finding " + g_sLayerName + " Layer.");
		return;
	}
	oLayerRS = oLayer.Records;
	if ( oLayerRS == null ){
		MessageBox("Error finding " + g_sLayerName + " Layer.");
		return;
	}
	g_oCurrentPoint.CoordinateSystem = Map.CoordinateSystem;
	var oDS = Map.Layers( g_sLayerName ).DataSource;
	var sSQL
	try{
		var sSearcTree = oForm.Pages("PAGE1").Controls("cb_Botanical").Value.replace(/'/g, "''");
		
		sSQL = "SELECT code FROM [CVD_AUDITS_CVD_BOTANICA] where DESCRIPTION = '" + sSearcTree + "';";
		var sBot = new VBArray(oDS.Execute(sSQL).ToArray(false)).toArray();
	
	}catch( ex ){
		MessageBox( "Error saving record, unable to locate tree details" );		
	}
	if ( !g_bChangeAudit ){
	//debugger
	//New Tree/Asset
		//Add point to Layer and set attributes
		oLayerRS.AddNew(g_oCurrentPoint);
		oLayerRS.Fields("ASSET_ID").Value = oForm.Pages("PAGE1").Controls("tbx_id").Value;
		var tempDate = new Date();
		oLayerRS.Update();

		
		var lAXFID = oLayerRS.Fields("AXF_OBJECTID").Value;

		//Console.Print( "a: " +  new VBArray(oDS.Execute("SELECT AXF_STATUS FROM [ALBURYTREES] where AXF_OBJECTID = " + lAXFID + ";").ToArray(false)).toArray()[0]);

		//See location and force record to be seen as new (AP8 bug workaround)
		sSQL = "Update [ALBURYTREES] set shape_X= " + g_oCurrentPoint.X;
		sSQL += ",shape_Y=" + g_oCurrentPoint.Y + ", AXF_STATUS = 1 where AXF_OBJECTID = " + lAXFID + ";";
		oDS.Execute (sSQL);
    
    //Console.Print( "b: " +  new VBArray(oDS.Execute("SELECT AXF_STATUS FROM [ALBURYTREES] where AXF_OBJECTID = " + lAXFID + ";").ToArray(false)).toArray()[0]);
		
		
		Map.Refresh( true );
		var oReturn;
		//Create Audit Entry
		try{
			sSQL = "INSERT INTO AUDITS (STREET_P,STREET_P_1,ASSET_ID,GENUS,COMMON_N,ORIGIN,HEIGHT,DBH,AGE,HEALTH," +
			"STRUCTUR,ULE,WORKS_RE,PRIORITY,DEFECTS,COMMENTS,WORKS_CA,RISK_SCO,FAILURE,PROBABIL," +
			"TARGET_R,NATURE_S,HOUSE_NU,STREET_N,SUBURB,POSTCODE,PROPERTY,POWERLIN, CURRENT_ST, INSPECTO," +
			"INS_DATE,UTC_TIME,ZONE_LEDGER,WIDTH,HERITAGE_TREE,BOTANICA," +
			"AXF_TIMESTAMP,AXF_STATUS,INSPECTOR_TYPE,TREE_IMA)";

			var oPage1C = oForm.Pages( "PAGE1" ).Controls;
			var oPage2C = oForm.Pages( "PAGE2" ).Controls;
			var oPage3C = oForm.Pages( "PAGE3" ).Controls;
			//sSQL += "VALUES( 'Street Tree','" ;
			//sSQL += oPage1C("cb_street_p").Value + "','" ;
			sSQL += "VALUES( '" + oPage1C("cb_street_p").Value + "','" ;
			sSQL += oPage1C("cb_streetPlanted").Text + "'," ;
			sSQL += oPage1C("tbx_id").Text + ",'" ;
			sSQL += oPage1C("cb_genus_spec").Text + "','" ;
			sSQL += oPage1C("tb_CommonN").Text + "','" ;
			sSQL += oPage1C("tb_Origin").Text + "'," ;
			sSQL += oPage2C("tbx_height").Value + "," ;
			sSQL += oPage2C("txt_dbh").Value + ",'" ;
			sSQL += oPage2C("cbx_TA").Text + "','" ;
			sSQL += oPage2C("cbx_TH").Text + "','" ;
			sSQL += oPage2C("cbx_ts").Text + "','" ;
			sSQL += oPage2C("cbx_ule").Text + "','" ;
			sSQL += oPage2C("tbx_works").Value + "','" ;
			sSQL += oPage2C("cbx_p").Text + "','" ;
			sSQL += oPage3C("cb_Defects").Text + "','" ;
			sSQL += oPage3C("tb_Comments").Text + "','" ;
			sSQL += oPage2C("cb_WorkC").Text + "',";
			sSQL += oPage2C("tbx_rs").Value + ",'" ;
			sSQL += oPage2C("cbx_fs").Value + "','" ;
			sSQL += oPage2C("cbx_fp").Value + "','" ;
			sSQL += oPage2C("cbx_to").Value + "','" ;
			sSQL += oPage2C("cb_NatureStrip").Value + "','" ;
			sSQL += oPage1C("tb_HouseNum").Text + "','" ;
			sSQL += oPage1C("cb_StreetN").Value + "','" ;
			sSQL += oPage1C("cb_Zone").Value + "','" ;		
			sSQL += "2640" + "','" + oPage1C("tb_HouseNum").Text + " " +oPage1C("cb_StreetN").Value + "','" ;			sSQL += oPage1C("cb_PowerLine").Value + "','" ;
			sSQL += oPage1C("cb_CURRENT_ST").Value + "','" ;
			sSQL += oPage1C("cb_inspector").Value + "','" ;
			var dateTemp = new Date( oPage3C("dp_DateVisited").Value );

			sSQL += formatDate( dateTemp, "Full" ) + "'" ;
			sSQL += ", getDate(),'" ;
			sSQL += oPage1C("cb_Zone").Value + "'," ;
			sSQL += oPage2C("tb_Width").Value + ",'" ;
			sSQL += oPage1C("cb_Heritage").Value + "'," ;
			sSQL += sBot[0] + "," ;
			sSQL += "getDate(), 1,'";
			sSQL += oPage1C("tb_InspectorType").Text + "','";
			sSQL += g_strTreeIma_Name + "');" ;

			oReturn = oDS.Execute (sSQL);
		}catch(ex){
			MessageBox("Error Adding details, check here: " + ex.Message);
			AddToLog( "NewTree_onOK_NewTree with sSQL = ", sSQL );
		}

	}else {
		//Change, still added an entry to Audits table.
		var oPage1C = oForm.Pages( "PAGE1" ).Controls;
		var oPage2C = oForm.Pages( "PAGE2" ).Controls;
		var oPage3C = oForm.Pages( "PAGE3" ).Controls;
		var oPage4C = oForm.Pages( "PAGE4" ).Controls;
		
		sSQL = "INSERT INTO AUDITS (STREET_P,STREET_P_1,ASSET_ID,GENUS,COMMON_N,ORIGIN,HEIGHT,DBH,AGE,HEALTH," +
		"STRUCTUR,ULE,WORKS_RE,PRIORITY,DEFECTS,COMMENTS,WORKS_CA,RISK_SCO,FAILURE,PROBABIL," +
		"TARGET_R,NATURE_S,HOUSE_NU,STREET_N,SUBURB,POSTCODE,PROPERTY,POWERLIN, CURRENT_ST, INSPECTO," +
		"INS_DATE,UTC_TIME,ZONE_LEDGER,WIDTH,HERITAGE_TREE,BOTANICA," +
		"AXF_TIMESTAMP,AXF_STATUS,INSPECTOR_TYPE,TREE_IMA) ";

		var oPage1C = oForm.Pages( "PAGE1" ).Controls;
		var oPage2C = oForm.Pages( "PAGE2" ).Controls;
		var oPage3C = oForm.Pages( "PAGE3" ).Controls;

		//sSQL += "VALUES( 'Street Tree','" ;
		//sSQL += oPage1C("cb_street_p").Value + "','" ;
		sSQL += "VALUES ( '" + oPage1C("cb_street_p").Value + "','" ;
		sSQL += oPage1C("tb_streetPlanted").Text + "'," ;
		sSQL += oPage1C("tbx_id").Text + ",'" ;
		sSQL += oPage1C("cb_genus_spec").Text + "','" ;
		sSQL += oPage1C("tb_CommonN").Text + "','" ;
		sSQL += oPage1C("tb_Origin").Text + "'," ;
		sSQL += oPage2C("tbx_height").Value + "," ;
		sSQL += oPage2C("txt_dbh").Value + ",'" ;
		sSQL += oPage2C("cbx_TA").Text + "','" ;
		sSQL += oPage2C("cbx_TH").Text + "','" ;
		sSQL += oPage2C("cbx_ts").Text + "','" ;
		sSQL += oPage2C("cbx_ule").Text + "','" ;
		sSQL += oPage2C("tbx_works").Value + "','" ;
		sSQL += oPage2C("cbx_p").Text + "','" ;
		sSQL += oPage3C("cb_Defects").Text + "','" ;
		sSQL += oPage3C("tb_Comments").Text + "','" ;
		sSQL += oPage2C("cb_WorkC").Text + "',";
		sSQL += oPage2C("tbx_rs").Value + ",'" ;
		sSQL += oPage2C("cbx_fs").Value + "','" ;
		sSQL += oPage2C("cbx_fp").Value + "','" ;
		sSQL += oPage2C("cbx_to").Value + "','" ;
		sSQL += oPage2C("cb_NatureStrip").Value + "','" ;
		sSQL += oPage1C("tb_HouseNum").Text + "','" ;
		sSQL += oPage1C("tb_StreetN").Value + "','" ;
		sSQL += oPage1C("cb_Zone").Value + "','" ;
		sSQL += "2640" + "','" + oPage1C("tb_HouseNum").Text + " " +oPage1C("tb_StreetN").Value + "','" ;
		sSQL += oPage1C("cb_PowerLine").Value + "','" ;
		sSQL += oPage1C("cb_CURRENT_ST").Value + "','" ;
		sSQL += oPage1C("cb_inspector").Value + "','" ;
		var dateTemp = new Date( oPage3C("dp_DateVisited").Value );
		var oDate = new Date();
		dateTemp.setHours( oDate.getHours(), oDate.getMinutes(), oDate.getSeconds());
		sSQL += formatDate( dateTemp, "Full" ) + "'" ;
		sSQL += ", getDate(),'" ;
		sSQL += oPage1C("cb_Zone").Value + "'," ;
		sSQL += oPage2C("tb_Width").Value + ",'" ;
		sSQL += oPage1C("cb_Heritage").Value + "'," ;
		sSQL += "'" + oPage1C("cb_Botanical").Value + "',"; //sBot[0] + "," ;
		sSQL += "getDate(), 1,'";
		sSQL += oPage1C("tb_InspectorType").Text + "', '";
		sSQL += g_strTreeIma_Name + "');";

		try{
			var sMessage = "";
			if ( oDS.IsOpen )
			{
				sMessage = "Error adding to Audits table 754: ";
				oReturn = oDS.Execute ( sSQL );			
				oReturn = null;
				sMessage = "Error updating Albury Trees: ";
				sSQL = "Update ALBURYTREES Set AXF_STATUS = " + getAXFStatusFromAssetID( g_lAssetID, oDS ) + ", AXF_TIMESTAMP = getDate() where Asset_ID = " + g_lAssetID + ";";
				oReturn = oDS.Execute ( sSQL );				
			}
		}catch( ex) {
			MessageBox(sMessage + ex.Message);
			AddToLog( sMessage + "NewTree_onOK_Update with sSQL = ", sSQL );
		}
	}
	oReturn = null;
	//oDS.Close();
	Map.Refresh( true );
	g_bChangeAudit = false;
}

function showStreetText( bShowTxt, oPage ){
		oPage( "cb_streetPlanted" ).Visible = !bShowTxt;
		oPage( "cb_StreetN" ).Visible = !bShowTxt;
		oPage( "tb_streetPlanted" ).Visible = bShowTxt;
		oPage( "tb_StreetN" ).Visible = bShowTxt;
		
		oPage( "tb_streetPlanted" ).Enabled = false;
		oPage( "tb_StreetN" ).Enabled = false;
}

function LoadListBoxAudits( oListBox, lAssestID, oForm ){
	var oDS = OpenAXF( g_sAXFFileName );
	var sSQL = "SELECT AXF_OBJECTID, INS_DATE AS UTC_TIME FROM [AUDITS] where Asset_ID = " + lAssestID + " order by INS_DATE;";

	//var oRS= Application.CreateAppObject("RecordSet");
	var oRS = oDS.Execute (sSQL);

	oRS.MoveFirst();
	var oObject;
	var sString;
	var outString;
	oListBox.Clear();
	while ( !oRS.EOF ){
		oObject = new Date(oRS.Fields( "UTC_TIME" ).Value);
		sString = oRS.Fields( "AXF_OBJECTID" ).Value;
		outString = sString + ", " + formatDate(oObject, "Short" );
		oRS.MoveNext();
	}

	if ( ! oRS.Fields( "AXF_OBJECTID" ).Value ){

		Application.MessageBox (g_lAssetID + " The last audit record was wrong or is corrupt.", apOkOnly, "Warning")
		oForm.Close();
		return;
	}

	oListBox.AddItem(sString,outString);
	oListBox.ListIndex = oListBox.ListCount - 1;
	oDS.Close();
	oRS = null;
	WaitCursor ( -1 );

}

function cbo_AssetStatus_onSelChange ( oCombo ){
	//if (g_LoadedStatusValue == "Proposed" && oCombo.Value == "Current"){
	if (oCombo.Value == "Current"){
		g_MadeTreeCurrent = true;
		//Console.print ( "onselchange: " + ( g_MadeTreeCurrent ) );
	}
	
	if (oCombo.value == "Redundant"){
		if ( Application.MessageBox ("Are you sure?", apYesNo, "Make Asset redundant.") == apYes){
			g_Redundant = true;
		}
	}
}
	

function btn_Change_onClick( oButton ){
	g_bLoading = true;
	if ( oButton.Parent.Controls("lb_Audits").ListIndex == -1 ) {
		return 0;
	}
	oButton.Parent.Parent.Pages("PAGE1").Activate();
	LoadData(oButton.Parent.Controls("lb_Audits").Value, oButton.Parent.Parent );
	g_bLoading = false;

	var oPage = oButton.Parent.Parent.Pages("PAGE1");
	//chk_Vacant_onClick (oPage.Controls("chkVacant"));
	checkForAudit( oPage );
}

function chk_Vacant_onClick ( oCheck ){
	var oForm = oCheck.Parent.Parent;
	var oPage1C = oForm.Pages( "PAGE1" ).Controls;
	var oPage2C = oForm.Pages( "PAGE2" ).Controls;
	var oPage3C = oForm.Pages( "PAGE3" ).Controls;
//Console.print ("ocheck " + oCheck.Value);

	if ( oCheck.Value ){

		oPage1C("cb_Genus_Spec").enabled = false;
		oPage1C("cb_Botanical").enabled = false;

		oPage2C("cbx_TA").enabled = false;
		oPage2C("cbx_TH").enabled = false;
		oPage2C("cbx_ule").enabled = false;
		oPage2C("cbx_TS").enabled = false;

		oForm.Pages( "PAGE2" ).Activate();

		oPage2C("cbx_fs").listIndex = 5;
		oPage2C("cbx_fp").listIndex = 5;

		oForm.Pages( "PAGE1" ).Activate();
		/*oPage2C("cbx_fp").enabled = false;
		cenabled = false;
		g_fp = "Negligible";
		g_fs = "None";
		g_ChangedToVacant = true;*/

		oPage1C("cb_CURRENT_ST").listIndex = 2;
				
		//Console.print (oPage2C("cbx_fp").Value); 

		//Console.print ( 

		g_Vacant = true;
	}
	else {
		oPage1C("cb_CURRENT_ST").listIndex = 0;

		oPage1C("cb_Genus_Spec").enabled = true;
		oPage1C("cb_Botanical").enabled = true;

		oForm.Pages( "PAGE2" ).Activate();
		oPage2C("tb_Width").Text = 1;
		oPage2C("txt_dbh").Text = 40;
		oPage2C("tbx_height").Text = 2;

		oPage2C("cbx_TA").enabled = true;
		oPage2C("cbx_TH").enabled = true;
		oPage2C("cbx_ule").enabled = true;
		oPage2C("cbx_TS").enabled = true;	

		oPage2C("cbx_TA").listIndex = 5;
		oPage2C("cbx_TH").listIndex = 2;
		oPage2C("cbx_ule").listIndex = 2;
		oPage2C("cbx_TS").listIndex = 1;

		oForm.Pages( "PAGE1" ).Activate();
		g_fp = "";
		g_fs = "";
		g_ChangedToVacant = false;
		g_Vacant = false;
	}
}

function NewTree_onUnLoad( oForm ){
	CloseRiskRS();
	g_bChangeAudit = false;
	g_oLatestAuditRS = null;
	g_Redundant = false;

}


function NewBeetle_onUnLoad( oForm ){
	g_bELBTools = false;
	g_lAssetID = 0;
}
function BeetleRPWorks_onUnLoad( oForm ){
	g_bELBTools = false;
	g_lAssetID = 0;
}

function NewTree_Page1_KillActive( oPage ){
	if ( oPage.Controls("chkVacant").Value){
		g_Vacant = true;	
	}
}
function NewTree_Page2_SetActive( oPage ){
	
	if ( g_bChangeAudit ){
		//Console.print ("not going to load the new values...");
		return;
	}
	oPage.Controls("cb_WorkC").ListIndex = 5;
		//Console.print ( "gvacant: " + g_Vacant);

	if (g_Vacant == true){
		oPage.Controls("tbx_height").value = 0;
		oPage.Controls("txt_dbh").value = 0;
		oPage.Controls("tb_width").value = 0;
		//oPage.Controls("cb_WorkC").ListIndex = 5;
		oPage.Controls("cbx_fp").ListIndex = 5;
		oPage.Controls("cbx_fs").ListIndex = 5;
		//oPage.Controls("cbx_to").ListIndex = 21;
		oPage.Controls("cbx_p").ListIndex = 4;
		oPage.Controls("tbx_rs").text = 0;

		oPage.Controls("cbx_TA").enabled = false;
		oPage.Controls("cbx_TH").enabled = false;
		oPage.Controls("cbx_ule").enabled = false;
		oPage.Controls("cbx_TS").enabled = false;
		oPage.Controls("cbx_fs").enabled = false;
		oPage.Controls("cbx_fp").enabled = false;
		oPage.Controls("cbx_fs").enabled = false;
		oPage.Controls("cbx_p").enabled = false;

		oPage.Controls("tbx_height").enabled = false;
		oPage.Controls("txt_dbh").enabled = false;
		oPage.Controls("tb_width").enabled = false;
		oPage.Controls("cb_WorkC").enabled = false;
		oPage.Controls("cbx_works").enabled = false;
		oPage.Controls("tbx_works").enabled = false;


		//Console.print ( oPage.Controls("cbx_fs").ListIndex );
	}
	else{
		oPage.Controls("cbx_TA").ListIndex = 4;
		oPage.Controls("cbx_TH").ListIndex = 2;
		oPage.Controls("cbx_ule").ListIndex = 2; // 0;
		oPage.Controls("cbx_TS").ListIndex = 1;
		oPage.Controls("cbx_fp").ListIndex = 5;
		oPage.Controls("cbx_fs").ListIndex = 5;
		oPage.Controls("cbx_p").ListIndex = 5;

		oPage.Controls("cbx_TA").enabled = true;
		oPage.Controls("cbx_TH").enabled = true;
		oPage.Controls("cbx_ule").enabled = true;
		oPage.Controls("cbx_TS").enabled = true;
		oPage.Controls("cbx_fs").enabled = true;
		oPage.Controls("cbx_fp").enabled = true;
		oPage.Controls("cbx_p").enabled = true;

		oPage.Controls("tbx_height").enabled = true;
		oPage.Controls("txt_dbh").enabled = true;
		oPage.Controls("tb_width").enabled = true;
		oPage.Controls("cb_WorkC").enabled = true;
		oPage.Controls("cbx_works").enabled = true;
		oPage.Controls("tbx_works").enabled = true;
	}
}

function NewTree_onOK( oForm){
	var oLayerRS = null
	var oLayer = Map.Layers( g_sLayerName );
	if ( oLayer == null ){
		Application.MessageBox("Error finding " + g_sLayerName + " Layer.");
		return;
	}
	oLayerRS = oLayer.Records;
	if ( oLayerRS == null ){
		Application.MessageBox("Error finding " + g_sLayerName + " Layer.");
		return;
	}

	g_oCurrentPoint.CoordinateSystem = Map.CoordinateSystem;
	var oDS = oLayer.DataSource;
	var sSQL;

	//debugger
	//New Tree/Asset
	//Add point to Layer and set attributes
 
	if (!g_bChangeAudit ){
		//oLayer.editable = false;
/*		var tempDate = new Date();
		var tDate = formatDate( tempDate, "Full" )
		sSQL = "INSERT INTO ALBURYTREES (ASSET_ID, AXF_STATUS, SHAPE_X, SHAPE_Y) VALUES ( " + g_intAssetID + ", 1, " + g_oCurrentPoint.X + ", " + g_oCurrentPoint.Y + " );";

		Console.print (sSQL);
		oDS.Execute(sSQL);
		oDS.Close();
*/
		//Console.print ("Added " + iCount + " rows");

		//Console.print (oLayerRS.Bookmark);
		//oLayerRS.AddNew(g_oCurrentPoint);
		//oLayerRS.Fields("ASSET_ID").Value = g_intAssetID; // parseInt (oForm.Pages("PAGE1").Controls("tbx_id").Text);
		//oLayerRS.Fields("AXF_STATUS").Value = 1;
		//var tempDate = new Date();
		//oLayerRS.Fields("AXF_TIMESTAMP").Value = formatDate( tempDate, "Full" );
		
		//oLayerRS.Update();
/*		try {
			oLayerRS.Update();
		}
		catch (ex){
			AddToLog("Line 1056: There was an error updating the recordset after AddFeature and update of field values");
			Application.MessageBox("There was an error updating the layers recordset.");
			return;
		}
*/
	}

	var lAXFID = oLayerRS.Fields("AXF_OBJECTID").Value;
	//g_intAssetID = oLayerRS.Fields("AXF_OBJECTID").Value;

	try{
		var sBot, botSQL;
		if (oForm.Pages("PAGE1").Controls("chkVacant").Value){		
			botSQL = "SELECT code FROM [CVD_AUDITS_CVD_BOTANICA] where DESCRIPTION = 'VACANT';";
		}
		else{
			var sSearcTree = oForm.Pages("PAGE1").Controls("cb_Botanical").Value.replace(/'/g, "''");
			var str = oForm.Pages("PAGE1").Controls("cb_Botanical").Value;
			if (str.substr(str.length -3, str.length) == ' sp'){
				botSQL = "SELECT code FROM [CVD_AUDITS_CVD_BOTANICA] where DESCRIPTION = '" + sSearcTree + ".';";
				//sBot = new VBArray(oDS.Execute(botSQL).ToArray(false)).toArray();
			}
			else{
				botSQL = "SELECT code FROM [CVD_AUDITS_CVD_BOTANICA] where DESCRIPTION = '" + sSearcTree + "';";
			}
		}
		sBot = new VBArray(oDS.Execute(botSQL).ToArray(false)).toArray();
		//Console.print ("sBot length: " + sBot.length);
	}
	catch(ex){
		AddToLog( "Error finding Botanical code. ", sSQL );
		Application.MessageBox("Error Adding details to the Audit table: \n" + sSQL); // ex.Message);

	}

	/*try{
		sSQL = "INSERT INTO MAINTENANCE (ASSET_ID,OPERATOR,CANOPY_LIFT,WEIGHT_REDUCTION,FORMATIVE_PRUNE,STRUCTURAL_PRUNE,PROPERTY_ASSET_CLEARANCE,DEAD_WOOD,MISTLETOE_REMOVAL,REMOVE,LV_WIRE_CL,HV_WIRE_CL,BROKENBRANCH,EPICORMICREMOVAL,MULCHREQUIRED,EXCLUDETARGET,TREESTAKEREMOVAL,INSTALLSTAKES,IRRIGATION,ROOTBALLMAINTENANCE,VISIBILITY,ABC_CLEARANCE,SERVICE_WIRE_CLEARANCE,HABITAT_PRUNE,REPLANT_LIST,COMMENTS,AXF_TIMESTAMP,AXF_STATUS) ";
		var oPage1C = oForm.Pages( "PAGE1" ).Controls;
		var oPage2C = oForm.Pages( "PAGE2" ).Controls;
		var worksArray = oPage2C("tbx_works").text.split("\\");

		Application.UserProperties("MasterWorksRequestArray") = worksArray;

		var oCANOPY_LIFT = "'false'";
		var oWEIGHT_REDUCTION = "'false'";
		var oFORMATIVE_PRUNE = "'false'";
		var oSTRUCTURAL_PRUNE = "'false'";
		var oPROPERTY_ASSET_CLEARANCE = "'false'";
		var oDEAD_WOOD = "'false'";
		var oMISTLETOE_REMOVAL = "'false'";
		var oREMOVE = "'false'";
		var oLV_WIRE_CL = "'false'";
		var oHV_WIRE_CL = "'false'";
		var oBROKENBRANCH = "'false'";
		var oEPICORMICREMOVAL = "'false'";
		var oMULCHREQUIRED = "'false'";
		var oEXCLUDETARGET = "'false'";
		var oTREESTAKEREMOVAL = "'false'";
		var oINSTALLSTAKES = "'false'";
		var oIRRIGATION = "'false'";
		var oROOTBALLMAINTENANCE = "'false'";
		var oVISIBILITY = "'false'";
		var oABC_CLEARANCE = "'false'";
		var oSERVICE_WIRE_CLEARANCE = "'false'";
		var oHABITAT_PRUNE = "'false'";

		for (var w = 0; w < worksArray.length; w++){
			switch (worksArray[w]){
				//case "Annual inspection":
						//break;
				case "Broken branch":
					oBROKENBRANCH = "'true'";
					break;
				case "Canopy lift":
					oCANOPY_LIFT = "'true'";
					break;
				case "Deadwood removal":
					oDEAD_WOOD = "'true'";
					break;			
				case "Epicormic removal":
					oEPICORMICREMOVAL = "'true'";
					break;
				case "Exclude/move target":
					oEXCLUDETARGET = "'true'";
					break;
				case "Formative pruning":
					oFORMATIVE_PRUNE = "'true'";
					break;
				case "HV wire clearance":
					oHV_WIRE_CL = "'true'";
					break;
				case "Install stakes":
					oINSTALLSTAKES = "'true'";
					break;
				case "Irrigation":
					oIRRIGATION = "'true'";
					break;
				case "LV wire clearance":
					oLV_WIRE_CL = "'true'";
					break;
				case "Mulch required":
					oMULCHREQUIRED = "'true'";
					break;
				case "Property/Asset clearance":
					oPROPERTY_ASSET_CLEARANCE = "'true'";
					break;
				case "Removal":
					oREMOVE = "'true'";
					break;
				case "Rootball maintenance":
					oROOTBALLMAINTENANCE = "'true'";
					break;
				case "Service wire clearance":
					break;
				case "Structural pruning":
					oSTRUCTURAL_PRUNE = "'true'";
					break;
				case "Tree stake removal":
					oTREESTAKEREMOVAL = "'true'";
					break;
				case "Visibilty/clearance":
					oVISIBILITY = "'true'";
					break;
				case "Weight reduction":
					oWEIGHT_REDUCTION = "'true'";
					break;
			}
		}
 
		sSQL += "VALUES (";
		sSQL += g_intAssetID; //oPage1C("tbx_id").Value;
		sSQL += ",'" + oPage1C("cb_inspector").Value + "'";
		sSQL += "," + oCANOPY_LIFT ;
		sSQL += "," + oWEIGHT_REDUCTION ;
		sSQL += "," + oFORMATIVE_PRUNE ;
		sSQL += "," + oSTRUCTURAL_PRUNE ;
		sSQL += "," + oPROPERTY_ASSET_CLEARANCE ;
		sSQL += "," + oDEAD_WOOD ;
		sSQL += "," + oMISTLETOE_REMOVAL ;
		sSQL += "," + oREMOVE ;
		sSQL += "," + oLV_WIRE_CL ;
		sSQL += "," + oHV_WIRE_CL ;
		sSQL += "," + oBROKENBRANCH ;
		sSQL += "," + oEPICORMICREMOVAL ;
		sSQL += "," + oMULCHREQUIRED ;
		sSQL += "," + oEXCLUDETARGET ;
		sSQL += "," + oTREESTAKEREMOVAL ;
		sSQL += "," + oINSTALLSTAKES ;
		sSQL += "," + oIRRIGATION ;
		sSQL += "," + oROOTBALLMAINTENANCE ;
		sSQL += "," + oVISIBILITY ;
		sSQL += "," + oABC_CLEARANCE ;
		sSQL += "," + oSERVICE_WIRE_CLEARANCE ;
		sSQL += "," + oHABITAT_PRUNE ;

		if (oPage1C("cb_CURRENT_ST").value == "Proposed"){
			sSQL += ",'true'";
		}
		else{
			sSQL += ",'false'";
		}
		//sSQL += ",'" + oPage1C("chkPlantList").Value + "'";
		sSQL += ",'added from new tree form'";
		sSQL += ", getDate(), 1" ;
		sSQL += ");";


		
		oDS.Execute (sSQL);
	}
	catch(ex){
		Application.MessageBox (sSQL);
		//MessageBox("Error on ln 1100: " + ex.Message);
		AddToLog( "Error on ln 1100: ", sSQL );
	}*/

	Console.clear();
	Console.print (  oForm.Pages("PAGE1").Controls("cb_CURRENT_ST").Value );

	if ( oForm.Pages("PAGE1").Controls("cb_CURRENT_ST").Value == "Proposed"){
		try{
			sSQL = "INSERT INTO PROPOSED_TREE (ASSET_ID, REQUESTEDBY, CHECK_, AXF_TIMESTAMP, AXF_STATUS) ";

			var oPage1C = oForm.Pages( "PAGE1" ).Controls;

			sSQL += "VALUES (" + g_intAssetID + ", '" ;
			sSQL += oPage1C("cb_inspector").Text + "',";
			sSQL += "'true', ";
			sSQL += "getDate(), 1";
			sSQL += ");" ;

			Console.print (sSQL);
			oDS.Execute (sSQL);
	
		}
		catch(ex){
			MessageBox("Error Adding details to the Proposed Tree table: " + ex.Message);
			AddToLog( "NewTree_onOK_NewTree with sSQL = ", sSQL );	
		}
		sSQL = "";
	}
	//else { 
		//Change, still added an entry to Audits table.
		var oPage1C = oForm.Pages( "PAGE1" ).Controls;
		var oPage2C = oForm.Pages( "PAGE2" ).Controls;
		var oPage3C = oForm.Pages( "PAGE3" ).Controls;
		var oPage4C = oForm.Pages( "PAGE4" ).Controls;

//Application.MessageBox (sBot[0]);
		
		sSQL = "INSERT INTO AUDITS (STREET_P,STREET_P_1,ASSET_ID,GENUS,COMMON_N,ORIGIN,HEIGHT,DBH,AGE,HEALTH," +
		"STRUCTUR,ULE,WORKS_RE,PRIORITY,DEFECTS,COMMENTS,WORKS_CA,RISK_SCO,FAILURE,PROBABIL," +
		"TARGET_R,NATURE_S,HOUSE_NU,STREET_N,SUBURB,POSTCODE,PROPERTY,POWERLIN, CURRENT_ST, INSPECTO," +
		"INS_DATE,UTC_TIME,ZONE_LEDGER,WIDTH,HERITAGE_TREE,BOTANICA," +
		"AXF_TIMESTAMP,AXF_STATUS,INSPECTOR_TYPE,TREE_IMA)";

		var oPage1C = oForm.Pages( "PAGE1" ).Controls;
		var oPage2C = oForm.Pages( "PAGE2" ).Controls;
		var oPage3C = oForm.Pages( "PAGE3" ).Controls;

		//sSQL += "VALUES( 'Street Tree','" ;
		//sSQL += oPage1C("cb_street_p").Value + "','" ;

/*STREET_P*/	sSQL += "VALUES( '" + oPage1C("cb_street_p").Text + "','" ;
/*STREET_P_1*/
		if (!g_bChangeAudit ){
			sSQL += g_sStreetP_1 + "'," ;
		}
		else {
			sSQL += oPage1C("tb_streetPlanted").Text + "'," ;
		}

/*STREET_P_1*/	//sSQL += oPage1C("tb_streetPlanted").Text + "'," ;
/*ASSET_ID*/	sSQL += g_intAssetID + ",'"; //  oPage1C("tbx_id").Text + ",'" ;

		if (oPage1C("cb_CURRENT_ST").value == "Proposed"){
/*GENUS*/		sSQL += "" + "','" ;
/*COMMON_N*/	sSQL += "" + "','" ;
		}
		else{
/*GENUS*/		sSQL += oPage1C("cb_genus_spec").Text + "','" ;
/*COMMON_N*/	sSQL += oPage1C("tb_CommonN").Text + "','" ;
		}
/*ORIGIN*/		sSQL += oPage1C("tb_Origin").Text + "'," ;
/*HEIGHT*/		sSQL += oPage2C("tbx_height").Value + "," ;
/*DBH*/			sSQL += oPage2C("txt_dbh").Value + ",'" ;
/*AGE*/			sSQL += oPage2C("cbx_TA").Text + "','" ;
/*HEALTH*/		sSQL += oPage2C("cbx_TH").Text + "','" ;
/*STRUCTUR*/	sSQL += oPage2C("cbx_ts").Text + "','" ;
/*ULE*/			sSQL += oPage2C("cbx_ule").Text + "','" ;
/*WORKS_RE*/	sSQL += oPage2C("tbx_works").Value + "','" ;
/*PRIORITY*/	sSQL += oPage2C("cbx_p").Text + "','" ;
/*DEFECTS*/		sSQL += oPage3C("cb_Defects").Text + "','" ;
/*COMMENTS*/	sSQL += oPage3C("tb_Comments").Text + "','" ;
/*WORKS_CA*/	sSQL += oPage2C("cb_WorkC").Text + "',";
/*RISK_SCO*/	sSQL += oPage2C("tbx_rs").Value + ",'" ;
/*FAILURE*/		sSQL += oPage2C("cbx_fs").Value + "','" ;
/*PROBABIL*/	sSQL += oPage2C("cbx_fp").Value + "','" ;
/*TARGET_R*/	sSQL += oPage2C("cbx_to").Value + "','" ;
/*NATURE_S*/	sSQL += oPage2C("cb_NatureStrip").Value + "','" ;
/*HOUSE_NU*/	sSQL += oPage1C("tb_HouseNum").Text + "','" ;
/*STREET_N*/	sSQL += oPage1C("cb_StreetN").Text + "','" ;
/*SUBURB*/		sSQL += oPage1C("cb_Zone").Value + "'," ;
/*POSTCODE*/	sSQL += 2640 + ",'";
/*PROPERTY*/	sSQL += oPage1C("tb_HouseNum").Text + " " +oPage1C("tb_StreetN").Value + "','" ;
/*POWERLIN*/	sSQL += oPage1C("cb_PowerLine").Value + "','" ;
/*CURRENT_ST*/	sSQL += oPage1C("cb_CURRENT_ST").Value + "','" ;
/*INSPECTO*/	sSQL += oPage1C("cb_inspector").Value + "','" ;
/*INS_DATE*/	var dateTemp = new Date( oPage3C("dp_DateVisited").Value );
				var oDate = new Date();
				dateTemp.setHours( oDate.getHours(), oDate.getMinutes(), oDate.getSeconds());
				sSQL += formatDate( dateTemp, "Full" ) + "'" ;
/*UTC_TIME*/	sSQL += ", getDate(),'" ;
/*ZONE_LEDGER*/	sSQL += oPage1C("cb_Zone").Value + "'," ;
/*WIDTH*/		sSQL += oPage2C("tb_Width").Value + ",'" ;
/*HERITAGE_TREE*/	sSQL += oPage1C("cb_Heritage").Value + "'," ;
/*BOTANICA*/	sSQL += sBot[0] + "," ;
/*AXF_TIMESTAMP*/	sSQL += "getDate(), 1,'";
/*INSPECTOR_TYPE*/	sSQL += oPage1C("tb_InspectorType").Text + "', '";
/*TREE_IMA*/	sSQL += g_strTreeIma_Name + "'";
				sSQL += ");";

		//try{
			var sMessage = "";
			if ( oDS.IsOpen )
			{
				sMessage = "Error adding to Audits table 1211: ";
				oReturn = oDS.Execute ( sSQL );			
				oReturn = null;
				
				sMessage = "1304: Error updating Albury Trees: ";
				var tID = 
				sSQL = "Update ALBURYTREES Set AXF_STATUS = " + getAXFStatusFromAssetID( g_intAssetID, oDS ) + ", AXF_TIMESTAMP = getDate() where Asset_ID = " + g_intAssetID + ";";
				//Console.print ( sSQL );
				oReturn = oDS.Execute ( sSQL );				
			}
		//}catch( ex) {
		//	MessageBox(sMessage + ex.Message);
			AddToLog( sMessage + "NewTree_onOK_Update with sSQL = ", sSQL );
		//}
	//}
	Console.print ("gmt: " + g_MadeTreeCurrent);
	//if (g_MadeTreeCurrent == 'false'){
		//Console.print ("equals string");
		//oDS.execute ( "update [proposed_tree] set check_ = 'false' where asset_id = " + g_intAssetID + ";" );
	//}

	if (g_MadeTreeCurrent == true){
		Console.print ("equals boolean");
		oDS.execute ( "update [proposed_tree] set check_ = 'false', AXF_STATUS = 2 where asset_id = " + g_intAssetID + ";" );
	}
	
	oReturn = null;
	oDS.Close();
	Map.Refresh( true );
	g_bChangeAudit = false;
	g_Vacant = false;
	g_Redundant = false;
    g_LoadedStatusValue = "";
    g_MadeTreeCurrent = false;

}

function ELBInspectUpdate_onClick( oButton ) {
	//Console.print ("starting inspect from applet..");

	g_bELBTools = true
	if (  !validateELBIns( oButton.Parent ) ){
		return;
	}
	var p_SQL;
	var iAssetid = g_lAssetID //oButton.Parent.Parent.Pages( "PAGE1" ).Controls( "tbx_id" ).Value;
	//var sZL = oButton.Parent.Parent.PAges( "PAGE" ).Controls( "tb_ZL" ).Value;
	var SQL_Values;
	var dateTemp = new Date( oButton.Parent.Controls( "dp_InspectDate" ).text );
	
	SQL_Values =  "VALUES( " + iAssetid + ",'";
	SQL_Values += oButton.Parent.Controls( "tb_Comments" ).text + "','" ;
	SQL_Values += oButton.Parent.Controls( "cb_inspector" ).text + "','";
	SQL_Values += oButton.Parent.Controls( "cb_InfesLevel" ).text + "','";
	SQL_Values += oButton.Parent.Controls( "cb_RecTreatment" ).text  + "',getDate(),'";
	//SQL_Values += sZL + "','";
	SQL_Values += formatDate( dateTemp, "Full" ) + "',";
	SQL_Values += "getDate(),1,'";
	SQL_Values += oButton.Parent.Controls( "tb_InspectorType" ).text + "');";
	//p_SQL = "INSERT INTO ELB( ASSET_ID,INSPECTION_COMMENTS,INSPECTOR_OPERATOR, INFESTATION_LEVEL,RECOMMEndED_TREATMENT,INSPECTDATE,Zone_Ledger, " +
	//		"DATE_INSPECTED,AXF_TIMESTAMP,AXF_STATUS,INSPECTOR_TYPE ) " + SQL_Values;
	
	p_SQL = "INSERT INTO ELB( ASSET_ID,INSPECTION_COMMENTS,INSPECTOR_OPERATOR, INFESTATION_LEVEL,RECOMMEndED_TREATMENT,INSPECTDATE, " +
			"DATE_INSPECTED,AXF_TIMESTAMP,AXF_STATUS,INSPECTOR_TYPE ) " + SQL_Values;
	//Application.messagebox (p_SQL);
	InsertinTable( p_SQL );
	setAXFStatus( iAssetid );
	if ( InsertinTable  == 0 ){
		Application.Messagebox( "Error: ELB Inspection Details not added successfully." );
	}else {
		Application.Messagebox( "ELB Inspection Details added successfully." );
		g_iChangessToSave += 1;
		resetPageLayer( oButton.Parent );
	}
}


function ELBWorksUpdate_onClick( oButton ) {
	//Console.print ("starting works from applet..");

	g_bELBTools = true
	if ( !validateELBWorks( oButton.Parent ) ){
		return;
	}
	var p_SQL;
	var iAssetid;
	iAssetid = g_lAssetID //oButton.Parent.Parent.Pages( "PAGE1" ).Controls( "tbx_id" ).Value;
	//var sZL = oButton.Parent.Parent.PAges( "PAGE" ).Controls( "tb_ZL" ).Value;
	var SQL_Values;

	var dateTemp = new Date(oButton.Parent.Controls( "dp_DateCompleated" ).text);
	

	SQL_Values =  "VALUES( " + iAssetid + ",'"+ oButton.Parent.Controls( "tb_Comments" ).text + "','" + 
				oButton.Parent.Controls( "cb_inspector" ).Value + "','" + oButton.Parent.Controls( "sb_Volume" ).Value + "','" +
				oButton.Parent.Controls( "cb_TretMethod" ).text + "',getDate(),'" + oButton.Parent.Controls( "cb_ChemType" ).Text + "','" +
				oButton.Parent.Controls( "cb_Additives" ).text + "','" +
				//sZL + "','" +
				formatDate( dateTemp, "Full" ) + "',getDate(),1,'";

	SQL_Values += oButton.Parent.Controls( "tb_InspectorType" ).text + "');";

	p_SQL = "INSERT INTO ELB( ASSET_ID,WORKS_COMMENTS, WORKS_OPERATOR,VOLUME,TREATMENT_METHOD,COMPLETEDATE,CHEMICALTYPE, ADDITIVES," +
			"DATE_COMPLETED,AXF_TIMESTAMP,AXF_STATUS,INSPECTOR_TYPE) " + SQL_Values;
	//p_SQL = "INSERT INTO ELB( ASSET_ID,WORKS_COMMENTS, WORKS_OPERATOR,VOLUME,TREATMENT_METHOD,COMPLETEDATE,CHEMICALTYPE, ADDITIVES," +
	//		"Zone_Ledger,DATE_COMPLETED,AXF_TIMESTAMP,AXF_STATUS,INSPECTOR_TYPE) " + SQL_Values;
//Application.Messagebox (p_SQL);
	InsertinTable( p_SQL );
	setAXFStatus( iAssetid );

	if ( InsertinTable  == 0 ){
		Application.Messagebox( "Error: ELB Work Details not added successfully." );
	}else {
		Application.Messagebox( "ELB Work Details added successfully." );
		g_iChangessToSave += 1;
		resetPageLayer( oButton.Parent );
	}
}

function RootInspection_onClick( oButton ){
	g_bELBTools = true 
	if ( !validateRootIns( oButton.Parent ) ){
		return;
	}
	var p_SQL;
	var iAssetid;
	iAssetid = g_lAssetID //oButton.Parent.Parent.Pages( "PAGE1" ).Controls( "tbx_id" ).Value;
	//var sZL = oButton.Parent.Parent.PAges( "PAGE" ).Controls( "tb_ZL" ).Value;
	var SQL_Values ;
	var thePageControls = oButton.Parent.Controls;
	var dateTemp = new Date( thePageControls( "dp_Inspected" ).text );
	SQL_Values= "VALUES( " + iAssetid + ",'"+
				thePageControls( "cb_inspector" ).Value + "','"+
				thePageControls( "cb_Method" ).Value + "','" +
				thePageControls( "tb_Comments" ).text + "',getDate(),'" +	
				//sZL + "','"+	
				formatDate( dateTemp, "Full" ) + "',getDate(), 1,'";
	SQL_Values += oButton.Parent.Controls( "tb_InspectorType" ).text + "');";

	//p_SQL = "INSERT INTO ROOTPRUNE( ASSET_ID, INSPECT_OPERATOR, METHODEXCAVATION, INSPECTIONCOMMENTS,INSPECTDATE," +
	//"Zone_Ledger,INSPECT_DATE, AXF_TIMESTAMP, AXF_STATUS,INSPECTOR_TYPE) " + SQL_Values;
	p_SQL = "INSERT INTO ROOTPRUNE( ASSET_ID, INSPECT_OPERATOR, METHODEXCAVATION, INSPECTIONCOMMENTS,INSPECTDATE," +
	"INSPECT_DATE, AXF_TIMESTAMP, AXF_STATUS,INSPECTOR_TYPE) " + SQL_Values;

	InsertinTable( p_SQL );
	setAXFStatus( iAssetid );
	if ( InsertinTable  == 0 ){
		Application.Messagebox( "Error: Inspection Details not added successfully." );
	}else {
		Application.Messagebox( "Inspection Details added successfully." );
		g_iChangessToSave += 1;
		resetPageLayer( oButton.Parent );
	}
}
function RootWorks_onClick( oButton ) {
	g_bELBTools = true
	if ( !validateRootWorks( oButton.Parent ) ){
		return;
	}

	var p_SQL;
	var iAssetid = g_lAssetID.toString(); //oButton.Parent.Parent.Pages( "PAGE1" ).Controls( "tbx_id" ).Value;
	//var sZL = oButton.Parent.Parent.PAges( "PAGE" ).Controls( "tb_ZL" ).Value;

	var SQL_Values ;
	var dateTemp = new Date( oButton.Parent.Controls( "date_Pruned" ).text );
	SQL_Values = "VALUES( " + iAssetid + ",'";
	SQL_Values += oButton.Parent.Controls( "tb_Photo" ).text + "',";
	SQL_Values += BoolToInt(oButton.Parent.Controls( "chk_RootBarrier" ).Value) + ",'";
	SQL_Values += oButton.Parent.Controls( "tb_Comments" ).text + "',";
	SQL_Values += BoolToInt(oButton.Parent.Controls( "chk_SuckerControl" ).Value) + ",'";
	SQL_Values += oButton.Parent.Controls( "chk_Plan" ).Value + "','";
	SQL_Values += oButton.Parent.Controls( "chk_SurfaceRootRemove" ).Value + "',getDate(),'";
	//SQL_Values += sZL + "','";
	SQL_Values += oButton.Parent.Controls( "cb_inspector" ).Value + "','";
	SQL_Values += formatDate( dateTemp, "Full" ) + "',getDate(), 1,'";
	SQL_Values += oButton.Parent.Controls( "tb_InspectorType" ).text + "');";

	p_SQL = "INSERT INTO ROOTPRUNE( ASSET_ID, PHOTO_1, ROOTBARRIER, WorksComments, SUCKER_CONTROL," +
			"PLAN_Root,SURFACEROOTREMOVAL,PRUNEDATE,WORKS_OPERATOR ,DATE_PRUNED, AXF_TIMESTAMP, AXF_STATUS,INSPECTOR_TYPE) " + SQL_Values;

	InsertinTable( p_SQL );
	setAXFStatus( iAssetid );
	if ( InsertinTable  == 0 ){
		Application.Messagebox( "Error: Root Work Details not added successfully." );
	}else {
		Application.Messagebox( "Root Work Details added successfully." );
		g_iChangessToSave += 1;
		resetPageLayer( oButton.Parent );
	}
}

function validateELBIns( oPage ){
	if ( !validateInspectCB( oPage( "cb_inspector" ) ) ){
		return false;
		
	}
	if ( oPage( "cb_InfesLevel" ).ListIndex < 0 ){
		MessageBox( "Please Select Infestation Level" );
		return false;
	}

	if ( oPage( "cb_RecTreatment" ).ListIndex < 0 ){
		MessageBox( "Please Select Recommended Treatment" );
		return false;
	}
	return true;
}

function validateELBWorks( oPage ){
	if ( !validateInspectCB( oPage( "cb_inspector" ) ) ){
		return false;
		
	}
	if ( oPage( "cb_TretMethod" ).ListIndex < 0 ){
		MessageBox( "Please Select Treatment method" );
		return false;
	}

	if ( oPage( "cb_ChemType" ).ListIndex < 0 ){
		MessageBox( "Please Select Chemical Type" );
		return false;
	}
	if ( oPage( "cb_Additives" ).ListIndex < 0 ){
		MessageBox( "Please Select Additives" );
		return false;
	}
	if ( oPage( "sb_Volume" ).Value < 0 ){
		MessageBox( "Please enter volume" );
		return false;
	}

	return true;

}

function validateRootWorks( oPage ){
	if ( !validateInspectCB( oPage( "cb_inspector" ) ) ){
		return false;
		
	}

	return true;
}

function validateInspectCB( oCombobox ){
	if ( oCombobox.ListIndex < 0 ){
		MessageBox("Please select Inspector" );
		return false;
	}
	return true;
}
function validateRootIns( oPage ){
	if ( !validateInspectCB( oPage( "cb_inspector" ) ) ){
		return false;
		
	}
	if( oPage( "cb_Method" ).ListIndex < 0 ){
		MessageBox( "Please Select Recomended Treatment Method" );
		return false;
	}
	return true;
}

function resetPageLayer( oPage ){

	for( j =1; j <=  oPage.Controls.Count; j++ ){

		switch ( oPage.Controls(j).Type ){
			case "COMBOBOX":
				if ( oPage.Controls(j).Name != "cb_inspector" ){
					oPage.Controls(j).ListIndex = -1;						
					oPage.Controls(j).Value = "";
				}
				break;
			case "EDIT":
				if ( oPage.Controls(j).Name == "tbx_id" ){
					break;
				}
				if( oPage.Controls(j).Name == "tb_InspectorType" ){
					break;
				}
				oPage.Controls(j).Text = "";
				oPage.Controls(j).Value = "";
				break;
			case "LABEL":
				break;	
			case "CHECKBOX":
					oPage.Controls(j).Value = false;
				break;
			case "BUTTON":
				break;
			case "DATETIME":
				break;	
			case "SLIDER":
				oPage.Controls(j).Value = oPage.Controls(j).DefaultValue;
				break;		
		}
				
	}

}

function page_inspector_onValidate( oEvent ){
	//Console.print ("f: " & oEvent.Result);
	if(!g_bELBTools){
		ThisEvent.Result = false;             
        ThisEvent.MessageText = "You must press the UPDATE Button";
        ThisEvent.MessageType = 48;
		g_bELBTools = false;
	}
}

function AXFStatusID( sAXFObjectID){
	var oDS = OpenAXF( g_sAXFFileName );
	var sSQL = "Select AXF_STATUS from Audits where AXF_OBJECTID = " + sAXFObjectID;
	var sBot = new VBArray( oDS.Execute( sSQL ).ToArray( false ) ).toArray();
	var axfid = sBot[0];
	oDS.Close();
	return TransformAXFID( axfid );
}

function getAXFStatusFromAssetID( lAssetIT, oDS){
	//Console.print ("starting... " + oDS.isOpen + " " + lAssetIT);
	//var oDS = OpenAXF( g_sAXFFileName );
	var sSQL = "Select AXF_STATUS from ALBURYTREES where Asset_ID = " + lAssetIT;
	var sBot = new VBArray( oDS.Execute( sSQL ).ToArray( false ) ).toArray();
	var axfid = sBot[0];

	//Console.print ("AXFID " + axfid); 
	//oDS.Close();
	return TransformAXFID( axfid );

	//Console.print ("AXFID 2 " + axfid); 
}
function TransformAXFID( axfid ){
	if ( axfid == null ){
		return 2;
	}
	if ( axfid == 1 ){
		return 1;
	}
	return axfid;

}

function LoadCombobox( oRecordSet, oComboBox ){

	if ( oComboBox.listcount != 0 ) {
		return;
	}
	if ( oRecordSet.RecordCount > 10 ) {
		//Large RS were taking to long to load, this solved the issue.
		oComboBox.AddItemsFromTable( CreateTempDBF( oRecordSet ), "Value", "Text" );
	} else {

		oRecordSet.MoveFirst();
		while( !oRecordSet.EOF ){
			oComboBox.AddItem( oRecordSet.Fields("Value"), oRecordSet.Fields("Text"));
			oRecordSet.MoveNext();

		}/*
		for( var cI=0; cI< oRecordSet.RecordCount; cI++ ) {
		oComboBox.AddItem( oRecordSet.Fields("Value"), oRecordSet.Fields("Text"));
		oRecordSet.MoveNext();
		}*/
	}
}

function CreateTempDBF( oRSSource ){
	/*This was required with a large data set. When LoadCombobox was being used it was very slow to process large data sets
	This way only one ArcPad event is fired not one for each item.*/
	var myRS = Application.CreateAppObject("RecordSet");
	myRS.Create( g_sAppletPath + "\\Temp.dbf", 0 );
	myRS.Fields.Append( "Value", 129, 255 );
	myRS.Fields.Append( "Text", 129, 255 );
	myRS.Update();

	oRSSource.MoveFirst();
	while( !oRSSource.EOF ){
		myRS.AddNew();
		myRS.Fields("Value").Value = oRSSource.Fields("Value").Value
		myRS.Fields("Text").Value = oRSSource.Fields("Text").Value
		myRS.Update();
		oRSSource.MoveNext();
	}
	myRS.Close();
	return g_sAppletPath + "\\Temp.dbf";
}

function OpenAXF( p_strAXFPath ){
	var pDS;
	pDS = Application.CreateAppObject( "DataSource" );

	pDS.Open( p_strAXFPath )
	if ( pDS.IsOpen ) {
		return pDS;
	}else{
		return null;
	}
}

function setFileNameAXF(){
	var strAXFPath
	strAXFPath = Map.Layers( g_sLayerName ).FilePath
	strAXFPath = strAXFPath.substring(0,strAXFPath.indexOf("|"))
	g_sAXFFileName = strAXFPath;
}

function cbxTA_onSelChange( oComboBox ){
   if (oComboBox.Text == "Young Tree" ){
        setCBO( oComboBox.Parent.Controls( "cbx_TH" ), "Good" );
        setCBO( oComboBox.Parent.Controls( "cbx_ts" ), "Good" );
        setCBO( oComboBox.Parent.Controls( "cbx_ule" ), "20+ years" )
		setCBO( oComboBox.Parent.Controls( "cb_WorkC" ), "Young Tree" )
		setTB( oComboBox.Parent.Controls( "tbx_works" ), "Young Tree" );
        //oComboBox.Parent.Controls( "tbx_works" ).Text = "No works";
		return;
	}
}


function cbWorkC_onSelChange( oComboBox ){

   if (oComboBox.Text == "No works" ){
		setTB( oComboBox.Parent.Controls( "tbx_works" ), oComboBox.Text );
		return;
	}

    if (oComboBox.Text == "Removal" ){
		setCBO( oComboBox.Parent.Parent.Pages( "PAGE1" ).Controls("cb_CURRENT_ST"), "Current" );
		return;
	}
}

function cbWorks_onSelChange( oComboBox ){
	if ( oComboBox.Parent.Controls( "cbx_works" ).Text == "No works" ){
		setTB( oComboBox.Parent.Controls( "tbx_works" ), oComboBox.Text );
		return;
	}
	if ( oComboBox.Parent.Controls( "tbx_works" ).Text != "" ){
		oComboBox.Parent.Controls( "tbx_works" ).Text += "\\";
        //oComboBox.Parent.Controls( "tbx_works" ).Text = "";
	}
	oComboBox.Parent.Controls( "tbx_works" ).Text += oComboBox.Text;
}

function RiskScoreChange( oComboBox ){
	//If nothing has been selected from the three comboboxes then return.
	if ( oComboBox.Parent.Controls("cbx_fs").Text == "" ){
		return;
	}
	if ( oComboBox.Parent.Controls("cbx_fp").Text == "" ){
		return;
	}
	if ( oComboBox.Parent.Controls("cbx_to").Text == "" ){
		return;
	}
	var oCBFS = oComboBox.Parent.Controls("cbx_fs").Value;
	var oCBFP = oComboBox.Parent.Controls("cbx_fp").Value;
	var oCBTO = oComboBox.Parent.Controls("cbx_to").Value;

	var dBookmark, lfs, lfp, lto;
	var sFind;
	sFind = "[FAILURE] =\"" + oCBFS + "\"";


	//dBookmark = g_oRiskFailRS.Find( sFind );
	dBookmark = FindRecord(g_oRiskFailRS, "FAILURE",  oCBFS);

	if (dBookmark != 0 ){
		g_oRiskFailRS.bookmark = dBookmark;
		lfs = g_oRiskFailRS( "FAILURE_VA" );
	}
	dBookmark = FindRecord(g_oRiskProbRS, "PROBABIL", oCBFP );
	//dBookmark = g_oRiskProbRS.Find( "[PROBABIL] =\"" + oCBFP + "\"" );
	if ( dBookmark != 0 ){
		g_oRiskProbRS.bookmark = dBookmark;
		lfp = g_oRiskProbRS( "PROBABIL_V" );
	}
	//dBookmark = g_oRiskTargRS.Find( "[TARGETR] =\"" + oCBTO + "\"" );
	dBookmark = FindRecord(g_oRiskTargRS, "TARGETR", oCBTO );
	if ( dBookmark != 0 ){
		g_oRiskTargRS.bookmark = dBookmark;
		lto = g_oRiskTargRS( "TARGETR_VA" );
	}
	oComboBox.Parent.Controls( "tbx_rs" ).Text = lfp * lfs * lto;
	oComboBox.Parent.Controls( "tbx_rs" ).Value = lfp * lfs * lto;
	if ( oComboBox.Parent.Parent.Name != "NewTree" ){
		oComboBox.Parent.Parent.Pages( "Page1" ).Controls( "tbx_rs" ).Text = lfp * lfs * lto;
		oComboBox.Parent.Parent.Pages( "Page1" ).Controls( "tbx_rs" ).Value = lfp * lfs * lto;
		//UpdateRiskScoreDisplay( oComboBox.Parent.Parent );
	}
}

function ReportFileStatus( sfilePath ){

	var fso = Application.CreateAppObject("file");
	if ( !fso.Exists( sfilePath ) ){
		return false;
	}
	return true;
}

function OpenRiskRS(){

	if ( !g_bRiskFailisOpen ){
		if ( ReportFileStatus( g_sAppletPath + g_sRisk_Fail_RS_Location ) ){
			g_oRiskFailRS.open( g_sAppletPath + g_sRisk_Fail_RS_Location );
			bRiskFailisOpen = true;
		} else {
			Application.MessageBox ( "Please Check that Data path is configured correctly" );
		}
	}
	if ( !g_bRiskProbisOpen ){
		if ( ReportFileStatus( g_sAppletPath + g_sRisk_Prob_RS_Location ) ){
			g_oRiskProbRS.open( g_sAppletPath + g_sRisk_Prob_RS_Location );
			bRiskProbisOpen = true;
		} else {
			Application.MessageBox ( "Please Check that Data path is configured correctly" );
		}
	}
	if ( !g_bRiskTargisOpen ){
		if (ReportFileStatus( g_sAppletPath + g_sRisk_Targ_RS_Location ) ){
			g_oRiskTargRS.open( g_sAppletPath + g_sRisk_Targ_RS_Location );
			bRiskTargisOpen = true;
		} else {
			Application.MessageBox ( "Please Check that Data path is configured correctly" );
		}
	}

}

function CloseRiskRS(){

	if ( g_bRiskFailisOpen ){
		g_oRiskFailRS.Close;
		bRiskFailisOpen = false;
	}
	if ( g_bRiskProbisOpen ){
		g_oRiskProbRS.Close;
		bRiskProbisOpen = false;
	}
	if ( g_bRiskTargisOpen ){
		g_oRiskTargRS.Close;
		bRiskTargisOpen = false;
	}
}

/**************************************
Format a given date
**************************************/
function formatDate( oDate, sFormat ) {

	switch( sFormat ) {
		case "ANSI":
		return oDate.getYear() + ""
		+ pad20( ( oDate.getMonth() + 1 ) ) + ""
		+ pad20( oDate.getDate() ) + ""
		+ pad20( oDate.getHours() ) + ""
		+ pad20( oDate.getMinutes() ) + ""
		+ pad20( oDate.getSeconds() );
		break;

		case "Full":
		return oDate.getYear() + "-"
		+ pad20( ( oDate.getMonth() + 1 ) ) + "-"
		+ pad20( oDate.getDate() ) + " "
		+ pad20( oDate.getHours() ) + ":"
		+ pad20( oDate.getMinutes() ) + ":"
		+ pad20( oDate.getSeconds() );
		break;

		case "Short":
		return pad20( oDate.getDate() ) + "-"
		+ pad20( ( oDate.getMonth() + 1 ) ) + "-"
		+ oDate.getYear();
		break;
	}

}

/**************************************
Ensure the return value always has two digits.
**************************************/
function pad20( sValue ) {

	var tempVal = "00" + sValue;
	return tempVal.substr( tempVal.length -2 );
}

function getUTCFullDate( oDate ){
	var UTCDate;
	UTCDate = pad20( oDate.getUTCDate() ) + "/";
	UTCDate += pad20( oDate.getUTCMonth() ) + "/";
	UTCDate += pad20( oDate.getUTCYear() ) + "/";
	UTCDate += oDate.getUTCFullYear() + " ";
	UTCDate += pad20( oDate.getUTCHours().toString() ) + ":";
	UTCDate += pad20( oDate.getUTCMinutes().toString() ) + ":";
	UTCDate += pad20( oDate.getUTCSeconds().toString() );
	return UTCDate;
}

function Map_onSelectionChanged(){
	if  (! g_bAuditSession && ! g_bChangeAudit && ! g_bNewPlantingSession){
		return;
	}

	Application.ExecuteCommand ( "modeselect" );
	g_bAuditSession  = false;

	var oRS = Map.Layers( g_sLayerName ).Records;
	oRS.Bookmark = Map.SelectionBookmark;
	g_lAssetID = oRS.Fields("ASSET_ID").Value;
	//processZL( oRS.Fields.Shape.X, oRS.Fields.Shape.Y );

	if( g_lAssetID != undefined ){
		if ( g_bChangeAudit ){
			Applets("AlburyTrees").Forms("NewTree").Show();
		}else {
			Map.Layers( g_sLayerName ).Edit( Map.SelectionBookmark );
		}
	}
	//Console.print ("on map sel change: " & g_lAssetID);
}


/**************************************
Function to find a record in a record set
- returns the bookmark of the first record
where sFieldValue is found in sFieldName
- if not found, a value of 0 is returned
**************************************/
function FindRecord( pRS, sFieldName, sFieldValue ){
	try
	{

		// Check that record(s) exist in the RS
		if( pRS.RecordCount == 0 ) return 0;

		// Check that the field name exists in the RS
		try
		{
			var pTemp = pRS.Fields( sFieldName ).Name;
		}
		catch( ex )
		{
			logToFile( "Failed in FindRecord(): " + sFieldName + " is not a field in the current RS" );
			return 0;
		}

		// get the initial bookmark of the RS (if it exists)
		var iInitialBM
		try{
			iInitialBM = pRS.Bookmark;
		}
		catch( ex )
		{
			//Bookmark mustn't exist
			iInitialBM = -1;
		}

		// Visit each record in the recordset looking for a match to the expression
		pRS.MoveFirst();
		var cI = 0;
		while( !pRS.EOF )
		{
			// If the record is found, return its bookmark and set the RS back to its initial BM
			if( pRS.Fields( sFieldName ).Value == sFieldValue )
			{
				iFoundBM = pRS.Bookmark;
				if( iInitialBM != -1 ) pRS.Bookmark = iInitialBM;
				return iFoundBM;
			}
			pRS.MoveNext();
		}
		// if makes it here, not found => return 0
		return 0;

	}
	catch( ex )
	{
		logToFile("Failed in FindRecord()" )
		return 0;
	}
}
function LoadData( sAXF_ID, oForm){
	WaitCursor ( 1 );
	if ( g_sAXFFileName == "" ) {
		return;
	}
	//++ open the selected AXF file
	//var pDS = OpenAXF(g_sAXFFileName);
	var pDS = Map.Layers( g_sLayerName ).DataSource;
	if ( pDS == null ) {
		MessageBox( "Open DataSource failed" );
		return 0;
	}
	//++ execute the input SQL statement
	var sSQL = "SELECT A.* " +
			",B.DESCRIPTION As BotanicaName " +
 		"FROM Audits A " +
		"LEFT JOIN CVD_AUDITS_CVD_BOTANICA B " +
		"ON A.BOTANICA=B.CODE ";
	sSQL += " WHERE A.AXF_OBJECTID=" + sAXF_ID;
	g_oLatestAuditRS = pDS.Execute( sSQL );
	g_oLatestAuditRS.MoveFirst();
	if ( g_oLatestAuditRS == null ){
		MessageBox( "Error reading table" );
		pDS.Close();
		return null;
	}
	var oThePage = oForm.Pages( "PAGE1" );

	var oPage1C = oForm.Pages( "PAGE1" ).Controls;
	var oPage2C = oForm.Pages( "PAGE2" ).Controls;
	var oPage3C = oForm.Pages( "PAGE3" ).Controls;
	var oPage4C = oForm.Pages( "PAGE4" ).Controls;
	//PAGE1
	oForm.Pages( "PAGE1" ).Activate();

	setCBO( oPage1C( "cb_street_p" ), getFieldValue( g_oLatestAuditRS.Fields( "STREET_P" ) ) );
	setTB( oPage1C( "tb_streetPlanted" ), getFieldValue( g_oLatestAuditRS.Fields( "STREET_P_1" ) )  );
	setTB( oPage1C( "tb_HouseNum" ), getFieldValue( g_oLatestAuditRS.Fields( "HOUSE_NU" ) ) );
	setTB( oPage1C( "tb_StreetN" ), getFieldValue( g_oLatestAuditRS.Fields( "STREET_N" ) ) );
	setTB( oPage1C( "cb_Zone" ), getFieldValue( g_oLatestAuditRS.Fields( "ZONE_LEDGER" ) ) );

	if (oPage1C("cb_Zone").Value == "" ){
		setTB(oPage1C("cb_Zone"), g_sZL);
	}

	setCboOnText ( oPage1C( "cb_genus_spec" ), getFieldValue( g_oLatestAuditRS.Fields( "GENUS" ) ) );
	cb_genus_spec_onSelchange( oPage1C( "cb_genus_spec" ) );

	g_iBotIndex = getFieldValue( g_oLatestAuditRS.Fields( "Botanica" ) );
	//Application.MessageBox (g_iBotIndex);
	if (g_iBotIndex === 768){
		oPage1C("cb_Botanical").Enabled = false;
		oPage1C( "tb_CommonN" ).Enabled = false;
	}
	oPage1C("cb_Botanical").Enabled = true;
	oPage1C( "tb_CommonN" ).Enabled = true;
	setCBO( oPage1C( "cb_Botanical" ), getFieldValue( g_oLatestAuditRS.Fields( "BotanicaName" ) ) );


	setTB( oPage1C( "tb_CommonN" ),getFieldValue( g_oLatestAuditRS.Fields( "COMMON_N" ) ) );
	setTB( oPage1C( "tb_Origin" ), getFieldValue( g_oLatestAuditRS.Fields( "ORIGIN" ) ) );
	setCBO( oPage1C( "cb_PowerLine" ), getFieldValue( g_oLatestAuditRS.Fields( "POWERLIN" ) ) );
	setCBO( oPage1C( "cb_CURRENT_ST" ), getFieldValue( g_oLatestAuditRS.Fields( "CURRENT_ST" ) ) );
	setTB( oPage1C( "tbx_id" ), getFieldValue( g_oLatestAuditRS.Fields( "ASSET_ID" ) ) );
	setCBO( oPage1C( "cb_inspector" ), getFieldValue( g_oLatestAuditRS.Fields( "INSPECTO" ) ) );
	setTB( oPage1C( "tb_InspectorType" ), getFieldValue( g_oLatestAuditRS.Fields( "INSPECTOR_TYPE" ) ) );
	
	g_intAssetID = parseInt ( oPage1C( "tbx_id" ).Value );

	if ( oPage1C( "tb_InspectorType" ).Text == "" ){
		setTB( oPage1C( "tb_InspectorType" ), oPage1C( "cb_inspector" ).Text);
	}
	WaitCursor ( -1 );
	var bValue = getFieldValue( g_oLatestAuditRS.Fields( "HERITAGE_TREE" ) );
	if (  bValue == "" ){
		oPage1C("cb_Heritage").Value = false;
	}else{
		oPage1C("cb_Heritage").Value = bValue;
	}
	oForm.Pages( "PAGE2" ).Activate();
	WaitCursor ( 1 );
	//PAGE2
	setTB( oPage2C( "tbx_height" ), getFieldValue( g_oLatestAuditRS.Fields( "HEIGHT" ) ) );
	setTB( oPage2C( "txt_dbh" ), getFieldValue( g_oLatestAuditRS.Fields( "DBH" ) ) );
	setCBO( oPage2C( "cbx_TA" ), getFieldValue( g_oLatestAuditRS.Fields( "AGE" ) ) );
	setCBO( oPage2C( "cbx_TH" ), getFieldValue( g_oLatestAuditRS.Fields( "HEALTH" ) ) );
	setCBO( oPage2C( "cbx_ule" ), getFieldValue( g_oLatestAuditRS.Fields( "ULE" ) ) );
	setCBO( oPage2C( "cbx_ts" ), getFieldValue( g_oLatestAuditRS.Fields( "STRUCTUR" ) ) );
	setCBO( oPage2C( "cb_WorkC" ), getFieldValue( g_oLatestAuditRS.Fields( "WORKS_CA" ) ) );
	setTB( oPage2C( "tb_Width" ), getFieldValue( g_oLatestAuditRS.Fields( "WIDTH" ) ) );
	setCBO( oPage2C( "cbx_works" ),getFieldValue( g_oLatestAuditRS.Fields( "WORKS_RE" ) ) );
	setTB( oPage2C( "tbx_works" ), getFieldValue( g_oLatestAuditRS.Fields( "WORKS_RE" ) ) );
	setCBO( oPage2C( "cbx_fp" ), getFieldValue( g_oLatestAuditRS.Fields( "PROBABIL" ) ) );
	setCBO( oPage2C( "cbx_fs" ), getFieldValue( g_oLatestAuditRS.Fields( "FAILURE" ) ) );
	setTB( oPage2C( "tbx_rs" ), getFieldValue( g_oLatestAuditRS.Fields( "RISK_SCO" ) ) );
	setCBO( oPage2C( "cbx_to" ), getFieldValue( g_oLatestAuditRS.Fields( "TARGET_R" ) ) );
	setCBO( oPage2C( "cb_NatureStrip" ), getFieldValue( g_oLatestAuditRS.Fields( "NATURE_S" ) ) );
	setCBO( oPage2C( "cbx_p" ), getFieldValue( g_oLatestAuditRS.Fields( "PRIORITY" ) ) );

	WaitCursor ( -1 );
	oForm.Pages( "PAGE3" ).Activate();
	WaitCursor ( 1 );
	//PAGE3
	oPage3C( "tb_Comments" ).Text = getFieldValue( g_oLatestAuditRS.Fields( "COMMENTS" ) );
	var sValue = getFieldValue( g_oLatestAuditRS.Fields( "DEFECTS" ) );
	if ( sValue == "" ){
		sValue = "None";
	}
	setTB ( oPage3C( "tb_defects" ), sValue );
	
	
	//oPage3C( "dp_DateVisited" ).Value = getFieldValue( g_oLatestAuditRS.Fields( "INS_DATE" ) );
    oPage3C( "dp_DateVisited" ).Value = new Date().getVarDate();
    
    
	WaitCursor ( -1 );
	oForm.Pages( "PAGE1" ).Activate();
	//++ close the DataSource
	//pDS.Close();
	pDS = null;
	WaitCursor ( -1 );
	//return pRS;
}

function setTB( oControl, sValue ){

	if ( sValue == null ){
		oControl.Text = "";
	}else{
		oControl.Text = sValue;
		oControl.Value = sValue;
	}
}
function  setCBO( oControl, sValue ) {
	for( cI = 0; cI< oControl.ListCount; cI++ ) {

		oControl.ListIndex = cI;
		if( oControl.Value.toUpperCase() == sValue.toUpperCase() ) {
			oControl.Value = sValue;
			break;
		}
	}
}

function setCboOnText( oControl, sValue ) {

	for( cI = 0; cI< oControl.ListCount; cI++ ) {

		oControl.ListIndex = cI;
		if( oControl.Text == sValue ) {
			break;
		}

	}

	oControl.Value = sValue;

}

function  setCBOReg( oControl, sRegExp ) {
	
	for( cI = 0; cI< oControl.ListCount; cI++ ) {

		oControl.ListIndex = cI++;
		if( oControl.Value.search( sRegExp )> -1 ) {
			//oControl.Value = sValue;
			break;
		}
	}
}

function removeStsuff( sString ){
	try{
		var sReturn = "";
		var aString = sString.split(" ");
		for ( i =0; i < aString.length -1; i++ ){
			sReturn += aString[i] + " ";
		}
		return 	sReturn;
	}
	catch ( ex ){
		return "";
	}
}
function cb_inspector_onSelChange( oComboBox ){
	setInspector( oComboBox.Value, oComboBox.Text, oComboBox.Parent.Parent );
	/*
	oComboBox.Parent.Controls("tb_InspectorType").Text = oComboBox.Text;
	oComboBox.Parent.Controls("tb_InspectorType").Value = oComboBox.Text;
	*/
}

function setInspector( sName, sType, oForm){
	for( ci =1; ci <= oForm.Pages.Count; ci++ ){
		if ( oForm.Pages(ci).Controls("tb_InspectorType") != undefined ){
			setTB( oForm.Pages(ci).Controls("tb_InspectorType"), sType );
		}
		if ( oForm.Pages(ci).Controls("cb_inspector") != undefined ){
			setCBO( oForm.Pages(ci).Controls("cb_inspector"), sName );
		}
	}
}

function genus_OnLoad ( oPage ){
	WaitCursor ( 1 );
	var oComboBox = oPage.Controls("cb_genus_spec");
	var myRS = Application.CreateAppObject("RecordSet");
	myRS.Open( g_sAppletPath + g_sLUT_GS, 1 );
	oComboBox.Clear()
	myRS.MoveFirst();
	myRS.Move(4);
	oComboBox.AddItem(myRS.Fields("GENUS").Value, myRS.Fields("GENUS").Value);
	
	/*while ( !myRS.EOF ){

		if ( myRS.Fields("GENUS").Value == oComboBox.Text ){
			oComboBox.Parent.Controls("cb_Botanical").AddItem(myRS.Fields("BOTANICAL_").Value, myRS.Fields("BOTANICAL_").Value);
		}
		myRS.MoveNext();
	}*/

}

function cb_genus_spec_onSelchange( oComboBox ){
	if (Application.UserProperties("NewPlanting")){ return; }
	WaitCursor ( 1 );
	var myRS = Application.CreateAppObject("RecordSet");
	myRS.Open( g_sAppletPath + g_sLUT_GS, 1 );
	oComboBox.Parent.Controls("cb_Botanical").Clear()
	myRS.MoveFirst();
	while ( !myRS.EOF ){
		if ( myRS.Fields("GENUS").Value == oComboBox.Text ){
			oComboBox.Parent.Controls("cb_Botanical").AddItem(myRS.Fields("BOTANICAL_").Value, myRS.Fields("BOTANICAL_").Value);
		}
		myRS.MoveNext();
	}
	oComboBox.Parent.Controls("cb_Botanical").Enabled = true;
	myRS.Close();
	WaitCursor ( -1 );
}

function cb_Botanical_onSelChange( oComboBox ){

	var myRS = Application.CreateAppObject("RecordSet");
	myRS.Open( g_sAppletPath + g_sLUT_GS, 1 );
	myRS.MoveFirst();
	while ( !myRS.EOF ){
		if ( myRS.Fields("BOTANICAL_").Value == oComboBox.Value ){
			oComboBox.Parent.Controls("tb_CommonN").Text = myRS.Fields("COMMON_NAM").Value;
			oComboBox.Parent.Controls("tb_Origin").Text =  myRS.Fields("STATUS").Value;
			oComboBox.Parent.Controls("tb_CommonN").Value = myRS.Fields("COMMON_NAM").Value;
			oComboBox.Parent.Controls("tb_Origin").Value =  myRS.Fields("STATUS").Value;
			myRS.MoveLast();
		}
		myRS.MoveNext();
	}
	myRS.Close();
}

function btn_Finish_onClick( oButton ){


	oButton.Parent.Parent.Close( true );
}
function goToPage( oEvent, sPage ){

	var oButton = oEvent.Object;

	if (g_Redundant){
		oButton.Parent.Parent.Pages( "PAGE3" ).Activate();
		return;
	}

	if ( ThisEvent.Object.Parent.Validate() ){
		oButton.Parent.Parent.Pages( sPage ).Activate();
	}

}

function cb_Defects_onselChange( oComboBox ){

	//if ( oComboBox.Parent.Controls( "tb_defects" ).Text == "None" ){
    if ( oComboBox.Text == "None" ){
		//oComboBox.Parent.Controls( "tb_defects" ).Text = "";
        setTB( oComboBox.Parent.Controls( "tb_defects" ), "" );
        setTB( oComboBox.Parent.Controls( "tb_Comments" ), "" );
       return;
	}
	if ( oComboBox.Parent.Controls( "tb_defects" ).Text != "" ){
		oComboBox.Parent.Controls( "tb_defects" ).Text += "\\";
	}
	oComboBox.Parent.Controls( "tb_defects" ).Text += oComboBox.Text;
}

function page1_onValidate( oEvent ){

	if ( !g_bLoading ){
		var oPage1C = oEvent.Object;

	/*	if ( oPage1C( "cb_CURRENT_ST" ).Text == "Current" && oPage1C( "chkVacant" ).Checked){
				oEvent.MessageText =  "You CANNOT have vacant and Current status";
				oEvent.Result = false;
				return false;
		}
		if ( oPage1C( "cb_CURRENT_ST" ).Text == "Proposed"){
			if ( !oPage1C( "chkVacant" ).Checked ){
				oEvent.MessageText =  "You MUST have vacant and Proposed status";
				oEvent.Result = false;
				return false;
			}
		}
	*/
		if ( oPage1C( "cb_streetPlanted" ).Visible  ){		
			if ( oPage1C( "cb_streetPlanted" ).ListIndex < 0 ){
				oEvent.MessageText =  "Please Select Street Planted";
				oEvent.Result = false;
				return false;
			}
		}
		if ( oPage1C( "tb_HouseNum" ).Text == "" ){
			oEvent.MessageText =  "Please enter a house number";
			oEvent.Result = false;
			return false;
		}
		if ( oPage1C( "cb_StreetN" ).Visible ){
			if (oPage1C( "cb_StreetN" ).ListIndex < 0 ){
				oEvent.MessageText =  "Please Select Street Name";
				oEvent.Result = false;
				return false;
			}
		}
		/*if ( oPage1C( "cb_Zone" ).ListIndex < 0  ){
		oEvent.MessageText =  "Please Select Zone";
		oEvent.Result = false;
		return false;
		}*/
		if( oPage1C( "cb_street_p" ).ListIndex < 0  ){
			oEvent.MessageText =  "Please Select Current Location of Tree";
			oEvent.Result = false;
			return false;
		}
		if( !oPage1C( "chkVacant" ).Value ){
			if( oPage1C( "cb_genus_spec" ).ListIndex < 0 ){
				oEvent.MessageText =  "Please Select Genus";
				oEvent.Result = false;
				return false;
			}
			if ( oPage1C("cb_Botanical").ListIndex < 0  ){
				oEvent.MessageText =  "Please Select Botanical Name";
				oEvent.Result = false;
				return false;
			}
		}
		if( oPage1C( "cb_PowerLine" ).ListIndex < 0  ){
			oEvent.MessageText =  "Please Select Power Line";
			oEvent.Result = false;
			return false;
		}
			if ( oPage1C( "cb_inspector" ).ListIndex < 0  ){
			oEvent.MessageText =  "Please Select Inspector";
			oEvent.Result = false;
			return false;
		}
	}
	oEvent.Result = true;
	return true;
}

function page2_onValidate( oEvent ){

	if ( !g_bLoading ){
		var oPage2C = oEvent.Object;

		try{
			if ( isNaN( oPage2C( "tbx_height" ).Text ) || oPage2C( "tbx_height" ).Text == "" ){
				oEvent.MessageText = "Enter a number for Height";
				oEvent.Result = false;
				return false;
			}
		}catch (ex){
			oEvent.MessageText = "Enter a number for Height";
			oEvent.Result = false;
			return false;
		}
		try{
			if ( isNaN( oPage2C( "txt_dbh" ).Text ) || oPage2C( "txt_dbh" ).Text == "" ){
				oEvent.MessageText = "Enter a number for DBH";
				oEvent.Result = false;
				return false;
			}
		}
		catch (ex){
			oEvent.MessageText = "Enter a number for DBH";
			oEvent.Result = false;
			return false;
		}
		if (!g_Vacant){
			if( oPage2C( "cbx_TA" ).ListIndex < 0  ){
				oEvent.MessageText = "Please Select Tree Age";
				oEvent.Result = false;
				return false;
			}
			if( oPage2C( "cbx_TH" ).ListIndex < 0  ){
				oEvent.MessageText = "Please Select Tree Health";
				oEvent.Result = false;
				return false;
			}
			if( oPage2C( "cbx_ule" ).ListIndex < 0  ){
				oEvent.MessageText = "Please Select ULE";
				oEvent.Result = false;
				return false;
			}
			if( oPage2C( "cbx_ts" ).ListIndex < 0 ){
				oEvent.MessageText = "Please Select Tree Structure";
				oEvent.Result = false;
				return false;
			}
		}
		if( oPage2C( "cb_WorkC" ).ListIndex < 0  ){
			oEvent.MessageText = "Please Select Works Catagory";
			oEvent.Result = false;
			return false;
		}

		try{

			if ( isNaN( oPage2C( "tb_Width" ).Text ) || oPage2C( "tb_Width" ).Text == "" ){
				oEvent.MessageText = "Enter a number for Canopy Width";
				oEvent.Result = false;
				return false;
			}
		}
		catch (ex){
			oEvent.MessageText = "Enter a number for Canopy Width";
			oEvent.Result = false;
			return false;
		}
		//cbx_works
		if( oPage2C( "tbx_works" ).Text == "" ){
			oEvent.MessageText = "Please select work required";
			oEvent.Result = false;
			return false;
		}
		if( oPage2C( "cbx_fp" ).ListIndex < 0  ){
			oEvent.MessageText = "Please Select Failure Prob";
			oEvent.Result = false;
			return false;
		}
		if( oPage2C( "cbx_fs" ).ListIndex < 0  ){
			oEvent.MessageText = "Please Select FS";
			oEvent.Result = false;
			return false;
		}
		if( oPage2C( "cbx_to" ).ListIndex < 0  ){
			oEvent.MessageText = "Please Select Target Range";
			oEvent.Result = false;
			return false;
		}
		if( oPage2C( "cb_NatureStrip" ).ListIndex < 0  ){
			oEvent.MessageText = "Please Select Nature Strip";
			oEvent.Result = false;
			return false;
		}
		if( oPage2C( "cbx_p" ).ListIndex < 0  ){
			oEvent.MessageText = "Please Select Prob";
			oEvent.Result = false;
			return false;
		}
	}
	oEvent.Result = true;
	return true;
}

function page3_onValidate( oEvent ){

	if ( !g_bLoading ){
		var oPage3C = oEvent.Object;
		if( oPage3C( "tb_defects" ).Text == "" ){
			oEvent.MessageText = "Please Select defects";
			oEvent.Result = false;
			return false;
		}
	}
	oEvent.Result = true;
	return true
}

function resetForm( oForm ){
	oPage1 = oForm.Pages("PAGE1");
	oPage2 = oForm.Pages("PAGE2");
	oPage3 = oForm.Pages("PAGE3");
	oPage4 = oForm.Pages("PAGE4");
	for( i =1; i <= oForm.Pages.Count; i++ ){
		for( j =1; j <=  oForm.Pages(i).Controls.Count; j++ ){

			switch ( oForm.Pages(i).Controls(j).Type ){
				case "COMBOBOX":
				oForm.Pages(i).Controls(j).ListIndex = -1;
				oForm.Pages(i).Controls(j).Value = "";
				break;
				case "EDIT":
				oForm.Pages(i).Controls(j).Text = "";
				oForm.Pages(i).Controls(j).Value = "";
				break;
				case "LABEL":
				break;
				case "CHECKBOX":
				break;
				case "BUTTON":
				break;
				case "DATETIME":
				break;
			}

		}

	}
}

function getFieldValue( oField ){
	if ( oField == null ){
		return "";
	}
	if ( oField.Value == undefined ){
		return "";
	}
	return oField.Value;
}

function AddToLog( sMessageLine1, sMessageLine2 ){
	var sLogFile = g_sAppletPath + g_sLogFileName;
	var oOutPutFile = Application.CreateAppObject("file");
	if ( !oOutPutFile.Exists( sLogFile ) ){
		oOutPutFile.Open( sLogFile, FileMode.Write )

	}
	if ( !oOutPutFile.Open( sLogFile, FileMode.Append ) ){
		//MessageBox( "File open failed" );
		return 0;
	}

	var sDateStamp = new Date()

	oOutPutFile.WriteLine( sDateStamp + ": " + sMessageLine1 );
	oOutPutFile.WriteLine( "   " + sMessageLine2 );
	oOutPutFile.Close();

}


function btn_GPSChange( oToolButton ){

}

if (typeof String.prototype.FileParts !== 'function') {
	String.prototype.FileParts  = function() {
		
		var qry = /(?:(.*)\\)?(.+)\.(.+)$/;
		var matches = this.match( qry );
		
		var aParts;
		
		if( matches[1] === null ) {
			matches[1] = "";
		}
		
		var oFileParts = {};
		oFileParts.Path = matches[1];
		oFileParts.Filename = matches[2];
		oFileParts.Extension = matches[3];
		
		return oFileParts;
		 
	};
}

/*********************************************
Automatically open the file browse dialog for the user to browse to a jpg
file and display in the picture object and insert filename into audit table 'TREE_IMA' column
**********************************************/
function OnPagePicture_SetActive()
{
	var pageControl = ThisEvent.Object;
	var picControl = pageControl.Controls("Image1");
	//Set the arguments of the Open dialog box
	var szDefExt = "jpg";
	var szFileFilter = "Picture Files|*.jpg";
	var szTitle = "Select Picture File";

	//Show the Open dialog box and get the result
	var szResult = CommonDialog.ShowOpen(szDefExt, szFileFilter, szTitle);//, lngFlags);
	//var szResult = CommonDialog.ShowPicture();//, lngFlags);

	if(szResult == null)
	{
		//cancel must have been clicked.
		g_strTreeIma_Name = "";
		return;
	}
	picControl.Value = szResult;

	//szResult is a string that looks like \My Documents\Pictures\waterfall.jpg
	//need to only get filename
	var filenamePos = szResult.lastIndexOf("\\");
	var filename = szResult.substr(filenamePos + 1);

    g_strTreeIma_Name = filename;

	//bail now that we have set the global tree image which is now used in the new audit
    //forget about inserting image stuff below as it is handled for us later
	return;
	
}

function InsertinTable( p_SQL ) {
	try {
		/*if ( g_sAXFFileName == "" ) { 
			return;
		}
		//++ open the selected AXF file
		var pDS = OpenAXF(g_sAXFFileName);
		if ( pDS == null ) {
			//Console.Print ( "Open DataSource failed" );
			return 0;
		}*/
		//++ execute the input SQL statement
		var pDS = Map.SelectionLayer.DataSource;

		var pRS = pDS.Execute( p_SQL );
		var pRS = 1
		
		if ( pRS != 1 ){	
			//pDS.Close();
			//Console.Print ( "Error updating table" );
		}
	}catch ( ex ) {
		//pDS.Close();
		MessageBox( "Insert Failed" )
		
		Applets( APPLET_NAME ).Execute ( "AddToLog( \"InsertinTable Failed with sSQL =" + p_SQL + " \", \"" + ex.description + "\")");
	}
	//++ close the DataSource
	//pDS.Close();
	pDS = null;
	return pRS;
}

function BoolToInt( oBool ) {	
	if ( oBool ){
		return 1;
	}
	return 0;
}

function setAXFStatus( lAssetID ){
	sSQL = "Update ALBURYTREES Set AXF_STATUS = " + getAXFStatusFromAssetID( lAssetID, Map.SelectionLayer.DataSource ) +
		", AXF_TIMESTAMP = getDate() where Asset_ID = " + lAssetID;
	InsertinTable( sSQL );
}


function TestScript ( objButton){
	var cboGenus = objButton.Parent.Controls("cb_genus_spec");

	cboGenus.DefaultValue = "Vacant";
	cboGenus.ListIndex = 652;
	//print (cboGenus.ListIndex);
//"Vacant";
}
