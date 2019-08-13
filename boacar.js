var oFSO = new ActiveXObject("Scripting.FileSystemObject");

var sDatabaseName = WScript.Arguments.Named("name");
var sDataFolder = WScript.Arguments.Named("data");
var sTempFolder = WScript.Arguments.Named("temp");
var sBackupFolder = WScript.Arguments.Named("backup");

var bWorking = true;

if ( undefined === sDatabaseName ) {
	WScript.Echo( "Specify database name with /name:<name>" );
	bWorking = false;
}

if ( undefined === sDataFolder ) {
	WScript.Echo( "Specify data path with /data:<name>" );
	bWorking = false;
}

if ( undefined === sTempFolder ) {
	WScript.Echo( "Specify temp folder with /temp:<name>" );
	bWorking = false;
}

if ( undefined === sBackupFolder ) {
	WScript.Echo( "Specify backup folder with /backup:<name>" );
	bWorking = false;
}

if ( ! bWorking ) {
	WScript.Quit();
}
var sDatabaseFile = oFSO.BuildPath(sDataFolder, sDatabaseName);
var sBackupFile = oFSO.BuildPath(sBackupFolder, sDatabaseName);
var sTempFile = oFSO.BuildPath(sTempFolder, sDatabaseName);

try {
	oFSO.DeleteFile( sTempFile );
} catch ( e ) {
	//~ WScript.Echo( e.message + ': ' + sTempFile );
}

WScript.Echo("CompactRepair",sDatabaseFile,"to",sTempFile);
var oACC = new ActiveXObject('Access.Application');
oACC.CompactRepair( sDatabaseFile, sTempFile, true );
oACC.Quit();

try {
	oFSO.DeleteFile( sBackupFile );
} catch( e ) {
	//~ WScript.Echo( e.message + ': ' + sBackupFile );
}
WScript.Echo("Moving",sDatabaseFile,"to",sBackupFile);
oFSO.MoveFile( sDatabaseFile, sBackupFile );

// copy temp to source, overwriting
try {
	oFSO.DeleteFile( sDatabaseFile );
} catch( e ) {
	//~ WScript.Echo( e.message + ': ' + sDatabaseFile );
}
WScript.Echo("Moving",sTempFile,"to",sDatabaseFile);
oFSO.MoveFile( sTempFile, sDatabaseFile );

WScript.Quit();
