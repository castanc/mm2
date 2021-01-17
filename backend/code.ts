import { GSResponse } from "../Models/GSResponse";
import { P } from "../Models/Parameters";
import { GSLog } from "./GSLog";
import { Service } from "./service";
import { Utils } from "./Utils";


function testMailMerge()
{
    let sv = new Service();
    let p = sv.loadParameters();
    P.testMode = true;
    let result = sv.mailMerge(p);
    Logger.log(result);
}

function testhardCoded() {
    let sv = new Service();
    let p = new P();
    p.addParameter(P.DATA_FILE_NAME, "userfile");
    p.addParameter(P.TEMPLATE_FILE_NAME, "MailTemplate");
    p.addParameter(P.OUTPUT_FOLDER_NAME, "MailMerge2_Output");
    p.addParameter("SENDER_NAME", "Cesar Alfredo Castano");
    p.addParameter("SENDER_TITLE", "Developer");
    p.addParameter(P.SUBJECT, "Test Mail");
    P.testMode = true;
    let result = sv.mailMerge(p);
    Logger.log(result);
    return result;
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

function getColumnNames() {
    let sv = new Service();
}

function getUrl(fileName) {
    return Utils.getUrl(fileName);
}

function getCreateDataFile(fileName) {
    let sv = new Service();
    return sv.getUrlDataFile(fileName);

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
            result.addError("error",`Exception at testMerge: ${ex.message}`);
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







