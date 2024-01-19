<?php
/*
 * gocas_request.php
 * kalyptus - 2019.12.09 04:36pm
 * use this API in requesting a copy of GOCAS..
 * Note:
 */

require_once 'config.php';
require_once 'Nautilus.php';
require_once 'CommonUtil.php';
require_once 'WebClient.php';
require_once 'WSHeaderValidatorFactory.php';
require_once 'MySQLAES.php';

$myheader = apache_request_headers();

if(!isset($myheader['g-api-id'])){
    echo "anggapoy nagawaan ti labat awa!";
    return;
}

if(stripos(APPSYSX, $myheader["g-api-id"]) === false){
    echo "anto la ya... sika lamet!";
    return;
}

//default-request charset
$chr_rqst = "UTF-8";
if(isset($myheader['g-char-request'])){
    $chr_rqst = $myheader['g-char-request'];
}
header("Content-Type: text/html; charset=$chr_rqst");

$validator = (new WSHeaderValidatorFactory())->make($myheader['g-api-id']);
//var_dump($myheader);
$json = array();
if(!$validator->isHeaderOk($myheader)){
    $json["result"] = "error";
    $json["error"]["code"] = $validator->getErrorCode();
    $json["error"]["message"] = $validator->getMessage();
    echo json_encode($json);
    return;
}

//GET HEADERS HERE
//Product ID
$prodctid = $myheader['g-api-id'];
//Computer Name / IEMI No
$pcname 	= $myheader['g-api-imei'];
//SysClient ID
$clientid = $myheader['g-api-client'];
//Log No
$logno 		= $myheader['g-api-log'];
//User ID
$userid		= $myheader['g-api-user'];

if(isset($myheader['g-api-mobile'])){
    $mobile = $myheader['g-api-mobile'];
}
else{
    $mobile = "";
}

$app = new Nautilus(APPPATH);
if(!$app->LoadEnv($prodctid)){
    $json["result"] = "error";
    $json["error"]["code"] = $app->getErrorCode();
    $json["error"]["message"] = $app->getErrorMessage();
    echo json_encode($json);
    return;
}

//Google FCM token
$token = $myheader['g-api-token'];
if(!$app->loaduser($prodctid, $userid)){
    $json["result"] = "error";
    $json["error"]["code"] = $app->getErrorCode();
    $json["error"]["message"] = $app->getErrorMessage();
    echo json_encode($json);
    return;
}

$param = file_get_contents('php://input');

//parse into json the PARAMETERS
$parjson = json_decode($param, true);
$par4sql = json_decode($param, true);


//detect the encoding used in the parameter...
//we perform the detection here so that we can properly handle characters
//such as (�). These characters are received as two part ASCII characters
//but can be detected once decoded(?) and encoded(?) again...
$enc_param = json_encode($parjson, JSON_UNESCAPED_UNICODE);
$encoding = mb_detect_encoding($enc_param);

//set the encoding to UTF-8/ISO-8859-1 if not ASCII
if($encoding !== "ASCII"){
    //primarily used by JAVA/PHP
    if($encoding !== "UTF-8"){
        $parjson = mb_convert_encoding($parjson, "UTF-8", $encoding);
    }
    
    //Possibly VB6/We used as default encoding for MySQL
    if($encoding !== "ISO-8859-1"){
        $par4sql = mb_convert_encoding($par4sql, "ISO-8859-1", $encoding);
    }
}

if(isset($par4sql['refernox'])){
    $value = $par4sql["refernox"];
    $field = "sTransNox";
}
else{
    $value = $par4sql["clientnm"];
    $field = "sClientNm";
}
//check if GOCAS was already saved previously...
$sql = "SELECT sTransNox" .
            ", sBranchCd" .
            ", dTransact" .
            ", sClientNm" . 
            ", sQMatchNo" .
            ", IFNULL(sGOCASNoF, sGOCASNox) sGOCASNox" . 
            ", cUnitAppl" . 
            ", sSourceCD" .
            ", sDetlInfo" .
            ", IFNULL(sCatInfox, sDetlInfo) sCatInfox" . 
            ", nDownPaym" . 
            ", sRemarksx" .
            ", sCreatedx" . 
            ", dCreatedx" .
            ", sVerified" . 
            ", dVerified" .
            ", cWithCIxx" .
            ", IFNULL(cTranStat, '') cTranStat" .
            ", sBranchCD" .
            ", cDivision" .
    " FROM Credit_Online_Application" .
    " WHERE $field LIKE '$value'" . 
    " ORDER BY sTransNox DESC" . 
    " LIMIT 1";

if(null === $rows = $app->fetch($sql)){
    $json["result"] = "error";
    $json["error"]["code"] = $app->getErrorCode();
    $json["error"]["message"] = $app->getMessage();
    echo json_encode($json);
    return;
}
elseif(empty($rows)){
    $json["result"] = "error";
    $json["error"]["code"] = AppErrorCode::RECORD_NOT_FOUND;
    $json["error"]["message"] = "Record not found";
    echo json_encode($json);
    return false;
}

//Check the status of GOCAS
//whatever the status may... allow download
if($rows[0]["cTranStat"] < "1"){
    $json["result"] = "error";
    $json["error"]["code"] = AppErrorCode::INVALID_APPLICATION;
    $json["error"]["message"] = "Application was not yet verified.";
    echo json_encode($json);
    return false;
}

// //Check if GOCAS is for branch
// if($rows[0]["sBranchCD"] != substr($clientid, 5, 4)){
//     $json["result"] = "error";
//     $json["error"]["code"] = AppErrorCode::INVALID_APPLICATION;
//     $json["error"]["message"] = "Application was not for this verified.";
//     echo json_encode($json);
//     return false;
// }

$json["result"] = "success";
$json["sTransNox"] = $rows[0]["sTransNox"];
$json["sBranchCd"] = $rows[0]["sBranchCd"];
$json["dTransact"] = $rows[0]["dTransact"];
$json["sClientNm"] = mb_convert_encoding($rows[0]["sClientNm"], $chr_rqst, "ISO-8859-1");
$json["sQMatchNo"] = $rows[0]["sQMatchNo"];
$json["sGOCASNox"] = $rows[0]["sGOCASNox"];
$json["cUnitAppl"] = $rows[0]["cUnitAppl"];
$json["sSourceCD"] = $rows[0]["sSourceCD"];

if (strtolower(substr($rows[0]["sQMatchNo"], 0, 2)) == "ap"){
    $json["sCatInfox"] = mb_convert_encoding($rows[0]["sDetlInfo"], $chr_rqst, "ISO-8859-1");
} else {
    $json["sCatInfox"] = mb_convert_encoding($rows[0]["sCatInfox"], $chr_rqst, "ISO-8859-1");
}

$json["nDownPaym"] = $rows[0]["nDownPaym"];
$json["sRemarksx"] = mb_convert_encoding($rows[0]["sRemarksx"], $chr_rqst, "ISO-8859-1");
$json["sCreatedx"] = $rows[0]["sCreatedx"];
$json["dCreatedx"] = $rows[0]["dCreatedx"];
$json["sVerified"] = $rows[0]["sVerified"];
$json["dVerified"] = $rows[0]["dVerified"];
$json["cWithCIxx"] = $rows[0]["cWithCIxx"];
$json["cTranStat"] = $rows[0]["cTranStat"];
$json["cDivision"] = $rows[0]["cDivision"];
echo json_encode($json, JSON_UNESCAPED_UNICODE );
?>