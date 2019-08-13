import System.Environment;
import System.IO;
import Microsoft.Office.Interop;
import Microsoft.Office.Interop.Access;

var sDatabaseName : String = argNamed("name");
var sDataFolder : String = argNamed("data");
var sTempFolder : String = argNamed("temp");
var sBackupFolder : String = argNamed("backup");

if (  sDatabaseName === "" ) {
	usage();
}

if (  sDataFolder === "" ) {
	usage();
}

if (  sTempFolder  === "" ) {
	usage();
}

if (  sBackupFolder === "" ) {
	usage();
}

var sDatabaseFile : String = System.IO.Path.Combine(sDataFolder, sDatabaseName);
var sBackupFile : String = System.IO.Path.Combine(sBackupFolder, sDatabaseName);
var sTempFile : String = System.IO.Path.Combine(sTempFolder, sDatabaseName);

try {
	File.Delete( sTempFile );
} catch ( e ) {
	//~ print( e.message + ': ' + sTempFile );
}

print("CompactRepair ",sDatabaseFile," to ",sTempFile);
/*var oACC = new ActiveXObject("Access.Application");
oACC.CompactRepair( sDatabaseFile, sTempFile, true );
var acQuitSaveNone = 2;
oACC.Quit(acQuitSaveNone);
*/
var oEngine = Microsoft.Office.Interop.Access.Dao.DBEngine;
oEngine.CompactDatabase(sDatabaseFile, sTempFile);
// copy source to backup, overwriting
try {
	File.Delete( sBackupFile );
} catch( e ) {
	//~ print( e.message + ': ' + sBackupFile );
}
print("Moving ",sDatabaseFile," to ",sBackupFile);
File.Move( sDatabaseFile, sBackupFile );

// copy temp to source, overwriting
try {
	File.Delete( sDatabaseFile );
} catch( e ) {
	//~ print( e.message + ': ' + sDatabaseFile );
}
print("Moving ",sTempFile," to ",sDatabaseFile);
File.Move( sTempFile, sDatabaseFile );
System.Environment.Exit(4);

function argNamed( sname : String ) {
	var result : String = "";
	var aCmdline = System.Environment.GetCommandLineArgs();
	var i : short = 1;
	for ( ; i < aCmdline.length; i++ ) {
		if (aCmdline[i].toLowerCase().slice(0, sname.length + 2) == ( "/" + sname.toLowerCase() + ":" )) {
			var inner : String = aCmdline[i].slice( sname.length + 2 ) 
			result = (inner.charAt(0) == '"' ? inner.slice(1,inner.length-1) : inner);
		}
	}
	return result;
}

function usage() {
	var aArgs = System.Environment.GetCommandLineArgs(); 
	print( aArgs[0] );
	print( " /name:<mdbname> /data:<datafolder> /temp:<tempfolder> /backup:<backupfolder>");
	System.Environment.Exit(1);
}
