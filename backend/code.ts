import { FileInfo } from "../Models/FileInfo";
import { GSResponse } from "../Models/GSResponse";
import { P } from "../Models/Parameters";
import { GSLog } from "./GSLog";
import { Service } from "./service";
import { SysLog } from "./SysLog";
import { Utils } from "./Utils";



function testMailMerge() {
    let sv = new Service();
    let p = sv.loadParameters();
    P.testMode = true;
    let result = sv.mailMerge(p);
    Logger.log(result);
}


function testGetFileInfo() {
    let result = Utils.getFileInfo("Automatizacion");
   
    result = Utils.getFileInfo("https://docs.google.com/spreadsheets/d/1g7-7Tem8TDkuoGbYx_33sLvM8rDYCetPxx3ix7j9R7c/edit#gid=0");

    result = Utils.getFileInfo("https://docs.google.com/document/d/1moWklfdcVr8euhpn3gREVnYfqkNi1EWZx5Mr9KZnNGo/edit?usp=sharing");


    result = Utils.getFileInfo("1g7-7Tem8TDkuoGbYx_33sLvM8rDYCetPxx3ix7j9R7c/edit#gid=0");


}


function testDirectoryAPI() {
    //todo: needs to enable directory api
    var user = AdminDirectory.Users.get("cesar.alfredo@pwc.com");
    Logger.log('User data:\n%s', JSON.stringify(user, null, 2));
}


function loadSettings() {

    let sv = new Service();
    let p = sv.loadParameters();

    let pars = p.getParsObject();
    return JSON.stringify(pars);
}


function getUrl(fileName) {

    let fi = Utils.getFileInfo(fileName);
    if ( fi != null )
    {
        if ( fileName.toLowerCase().indexOf("http")==0)
            fi.url = fileName;
            return JSON.stringify(fi);
    }
    else return `{url:""}`;
}

function getFileInfo(id,text,fileName)
{
    let fi =Utils.getFileInfo(fileName);
    fi.controlId = id;
    fi.controlText = text;
    return JSON.stringify(fi);

}

function getFileInfos(dataFile,namesFile,templateFile)
{
    let sv = new Service();
    let result = sv.getFileInfos(dataFile,namesFile,templateFile);
    return JSON.stringify(result);
}

function getCreateDataFile(controlId, labelId, controlText, fileName) {

    let result = null;
    try {
        result = Utils.getFileInfo(fileName);
        if (result != null) {
            result.controlId = controlId;
            result.controlText = controlText;
            result.labelId = labelId;
        }
    }
    catch (ex) {
        return `{error: "${ex.message}" }`
    }
    return JSON.stringify(result);

    //let sv = new Service();
    //return sv.getUrlDataFile(fileName);

}



/* @Include JavaScript and CSS Files */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent(); 303
}



function testMerge(formObject) {
    let sv = new Service();
    let result: GSResponse;
    let html = "";
    let p = new P();
    try {
        p.setParsObject(formObject);
        P.testMode = true;
        Logger.log("parameters", p);
        result = sv.mailMerge(p);
    }
    catch (ex) {
        Logger.log("Exception. parameters received:");
        Logger.log(p);
        Logger.log(ex);
        result.addError("error", `Exception at testMerge: ${ex.message}`);
    }
    return JSON.stringify(result);
}


/* @Process Form */
function processForm(formObject) {
    Logger.log("processForm()", formObject);
    let sv = new Service();
    let html = "";
    let result = new GSResponse();
    try {
        let p = new P();
        p.setParsObject(formObject);
        Logger.log("parameters", p);
        result = sv.mailMerge(p);
    }
    catch (ex) {
        Logger.log("Exception. parameters received:");
        Logger.log(ex);
        result.addError("error", ex.message);
    }
    return JSON.stringify(result);
}


function doGet(e) {
    //return HtmlService.createTemplateFromFile('nKit2').evaluate();
    return HtmlService.createTemplateFromFile('frontend/Parameters_AppKit').evaluate();
}


function getTemplateText(controlId, controlText, fileName)
{
    let result = new FileInfo(null);
    try {
        result = Utils.getFileInfo(fileName);
        if (result != null) {
            result.controlId = controlId;
            result.controlText = controlText;
;
            let sv = new Service();
            result.content = sv.getTemplateText(fileName);
        }
    }
    catch (ex) {
        Logger.log("getTemplateText() exception", ex.message);
    }
    return JSON.stringify(result);

}






