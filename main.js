function replacing_path(path_string){
	return path_string.replace(/\\/g,"/");	
}
const vbs_path =  'C:\\Users\\atul.narayan.phirke\\Desktop\\chatbot_training\\development\\backups\\\Protector\\qcConnector_vbs.vbs';
var vbs_path_final = replacing_path(vbs_path);

const vba_path = 'C:\\Users\\atul.narayan.phirke\\Desktop\\chatbot_training\\development\\backups\\\Protector\\logsUpdate_vba.xlsm'
var vba_path_final = replacing_path(vba_path);

const log_path = 'C:\\Users\\atul.narayan.phirke\\Desktop\\chatbot_training\\development\\backups\\\Protector\\log.txt'
var log_path_final = replacing_path(log_path);

const qc_log_path = 'C:\\Users\\atul.narayan.phirke\\Desktop\\chatbot_training\\development\\backups\\Protector\\QC_update_logs';
var qc_log_path_final = replacing_path(qc_log_path);

var testSetID = "123456";

var status = "Passed";

var testCaseName = "TF00_abc"
const
    spawn = require( 'child_process' ).spawnSync,
    vbs = spawn( 'cscript.exe', [ vbs_path_final,vba_path_final,testSetID,log_path_final,status,testCaseName,qc_log_path_final] );

