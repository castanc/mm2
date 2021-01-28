import { G } from './G';
import { Utils } from './Utils';
import { P } from '../Models/Parameters';
import { MailValidation } from '../Models/MailValidation';
import { R } from '../Models/R';
import { FileInfo } from '../Models/FileInfo';
import { GSResponse } from '../Models/GSResponse';
import { SysLog } from './SysLog';

export class Service {

    mailListText = "";
    templateText = "";
    grid: Array<Array<string>> = new Array<Array<string>>();


    saveParameters(sheet, p: P) {
        let grid = p.getGrid();
        for (var i = 0; i < p.parameters.length; i++) {
            sheet.appendRow(grid[i]);
        }
    }


    saveTemplate(url,text)
    {
        let result = -1;
        let doc = DocumentApp.openByUrl(url);
        if ( doc != undefined )
        {
            var body = doc.getBody();
            body.clear();
            body.appendParagraph(text);
            result = 0;
        }
        return result;
    }


    getColumnNamesSelects() {
        let parameters = new P();
        let folder = Utils.getCreateFolder(parameters.getParameter(P.FOLDER_NAME));
        let spreadSheet = Utils.getCreateSpreadSheet(folder, parameters.getParameter(P.DATA_FILE_NAME));

        let sheet = spreadSheet.getActiveSheet();
        var rangeData = sheet.getDataRange();
        var lastColumn = rangeData.getLastColumn();
        var lastRow = rangeData.getLastRow();
        //TODO: Read only row 1

        let grid = rangeData.getValues();
        return Utils.getHtmlFromArray("SELECT_COLUMNS", "None", grid[0], true);
    }

    createDataFile(fileName) {
        let parameters = new P();
        let folderName = parameters.getParameter(P.FOLDER_NAME);
        let folder = Utils.getCreateFolder(folderName);
        let spreadSheet = Utils.getCreateSpreadSheet(folder, fileName);
        let sheet = spreadSheet.getActiveSheet();
        var rangeData = sheet.getDataRange();
        var lastColumn = rangeData.getLastColumn();
        var lastRow = rangeData.getLastRow();
        //todo: read only first row
        if (lastRow == 1) {
            let arr = [];
            //arr.push(P.COL_SELECT);
            arr.push(P.COL_STATUS);
            arr.push(P.COL_RESULT);
            arr.push(P.COL_TIMESTAMP);
            arr.push(P.COL_INVALID_NAMES);
            arr.push(P.COL_INVALID_MAILS);
            arr.push(P.COL_RESOLVED_MAIL);
            arr.push(P.COL_NAMES);
            sheet.appendRow(arr);
        }
        return spreadSheet.getUrl();
    }

    getUrlDataFile(fileName: string): string {
        let ss = Utils.openSpreadSheet(fileName);
        let url = "";
        if (ss != null)
            url = ss.getUrl();;

        //if (url.length == 0)
        //    url = this.createDataFile(fileName);

        return url;
    }

    loadParameters(): P {

        let parameters = new P();

        let folderName = parameters.getParameter(P.FOLDER_NAME);
        let folder = Utils.getCreateFolder(folderName);
        //P.FOLDER_ID = folder.getId();
        //P.FOLDER_URL = `https://drive.google.com/drive/folders/${P.FOLDER_ID}`;

        let spreadSheet = Utils.getCreateSpreadSheet(folder, "Parameters", "Parameters");

        let sheet = spreadSheet.getSheetByName("Parameters");
        var rangeData = sheet.getDataRange();
        var lastColumn = rangeData.getLastColumn();
        var lastRow = rangeData.getLastRow();
        P.PARAMETERS_URL = spreadSheet.getUrl();

        if (lastRow == 1) {
            let resultTemplateName = "ResultTemplate";
            let resultTemplateText = HtmlService.createTemplateFromFile('frontend/resulttemplate').evaluate().getContent();
            Utils.writeTextFile(resultTemplateName, resultTemplateText, folder);
            parameters.addParameter(P.RESULT_TEMPLATE_FILE_NAME, resultTemplateName);

            parameters.addParameter(P.DATA_FILE_NAME, "");
            parameters.addParameter(P.NAMES_FILE, "");
            parameters.addParameter(P.TEMPLATE_FILE_NAME, "");
            parameters.addParameter(P.OUTPUT_FOLDER_NAME, `${folderName}_Output`);
            parameters.addParameter("SUBJECT", "");
            parameters.addParameter("SENDER_NAME", "");
            parameters.addParameter("SENDER_TITLE", "");
            //parameters.addParameter("STAKEHOLDERS_NAMES", "");

            this.saveParameters(sheet, parameters);
        }
        else {
            let grid = rangeData.getValues();
            for (var i = 0; i < grid.length; i++) {
                parameters.addParameter(grid[i][0], grid[i][1]);
            }
        }

        return parameters;
    }


    findMail(name) {
        name = name.toLowerCase().trim();
        var index = this.mailListText.indexOf(name);
        var email = "";
        var index2 = -1;
        if (index >= 0) {
            index += name.length;
            index2 = this.mailListText.indexOf(",", index);
            if (index2 < index)
                index2 = this.mailListText.indexOf("\t", index);
            if (index2 >= index) {
                index2++;
                var index3 = this.mailListText.indexOf("\n", index2);
                if (index3 > index2) {
                    var len = index3 - index2;
                    email = this.mailListText.substr(index2, len);
                }
            }
        }
        // else if ( this.grid.length > 0 )
        // {
        //     let row: Array<string[]> = this.grid.filter(x=>x[0]==name);
        //     if ( row.length > 0 )
        //     {
        //         email = row[0][1];
        //     }
        // }
        return email.trim();
    }


    validateNames(names: string): MailValidation {
        let mv = new MailValidation();
        mv.result = false;   //(list.length == mv.mails.length);

        let list = names.split('\n');
        let mail = "";

        for (var i = 0; i < list.length; i++) {
            list[i] = list[i].trim();
            if (list[i].length > 0) {
                if (!Utils.ValidateEmail(list[i])) {
                    mail = this.findMail(list[i]);
                    if (mail.length == 0) {
                        mv.invalidNames = `${mv.invalidNames}\n${list[i]}`;
                        mv.totalInvalidNames++;
                        mv.error = "Name not found";
                        mv.result = false;
                    }
                    else if (Utils.ValidateEmail(mail)) {
                        mv.mails.push(mail);
                        mv.validMails = `${mv.validMails}\n${mail}`;
                        mv.names = `${mv.names}\n${list[i]}`;
                        mv.result = true;
                    }
                    else {
                        mv.error = "Invalid Mail";
                        mv.invalidMails = `${mv.invalidMails}\n${mail}`;
                        mv.totalInvalidMails++;
                        mv.result = false;
                    }
                }
                else {
                    mv.mails.push(list[i]);
                    mv.validMails = `${mv.validMails}\n${list[i]}`;
                    mv.result = true;
                }

            }
        }
        return mv;

    }

    resultReport(r: R): string {
        let html = `
          <div id='body'>
            <p class='colored'><b>Subject: Mail Merge Execution Report</b></p>
          </br>
          <table>
              <tr>
                  <td>Execution Time:</td>
                  <td>${r.EXECUTION_TIME}</td>
              </tr>
          </tr>
          <tr>
          <td>Data File:</td>
          <td><a href="${r.FILE_LINK}">Data File</a></td>
      </tr>

          <tr>
                <td>Total Lines:</td>
                <td>${r.TOTAL_LINES}</td>
            </tr>
    
            <tr>
                <td>Successful Lines:</td>
                <td>${r.OK_LINES}</td>
            </tr>
            <tr>
                <td>Error Lines:</td>
                <td>${r.ERROR_LINES}</td>
            </tr>
            <tr>
            <td>Mail Effectively Sent:</td>
            <td>${r.MAILS_SENT}</td>
        </tr>       
        <tr>
        <td>Mail FAILED:</td>
        <td>${r.MAILS_FAILED}</td>
    </tr>         
            <tr>
                <td>Messages:</td>
                <td></td>
            </tr>
        <!--MESSAGES!-->
    </table>
      </div>`;
        return html;
    }


    getTemplateText(fileName) {
        let text = Utils.getDocTextByName(fileName);
        return text;
    }

    mailMerge(p: P): GSResponse {
        let folderName = p.getParameter(P.FOLDER_NAME);
        let processedFolderName = `${folderName}_Output`;
        let dataFileUrl = "";
        let folder = Utils.getCreateFolder(folderName);
        let ssName = p.getParameter(P.DATA_FILE_NAME);
        let columnNames = "";
        let cols = new Array<string>();
        let mailText = "";
        let subject = "";
        let sheet;
        let rangeData;
        let lastColumn = 0;
        let lastRow = 0;
        let mailListFileName = p.getParameter(P.NAMES_FILE);
        let templateFileName = p.getParameter(P.TEMPLATE_FILE_NAME);
        let mv;
        let sentMailsCount = 0;
        let notSentMailsCount = 0;
        let linsProcessed = 0;
        let i = 0;
        let j = 0;
        let colName: string = "";
        let r = new R();
        let fiData;
        let fiMailList;
        let fiTemplate;
        let response = new GSResponse();

        let fixedColumns = "status,result,timestamp,invalidnames,invalidmails,resolvedmail";
        let userColumns = "";
        let maxCols: number = 0;
        let newLastCol: number = 0;
        let namesColumn: string = p.getParameter("COL_NAMES");
        let go = true;
        let html = "";

        r.EXECUTION_TIME = Utils.getTimeStamp();
        SysLog.log(`XXXXXXXXXXXXXXXXXXXX   namesCol: ${p.getParameter(P.COL_NAMES)}`);
        SysLog.log(`XXXXXXXXXXXXXXXXXXXX   namesCol: ${p.getParameter(P.COL_NAMES)}`);
        let ss = Utils.openSpreadSheet(ssName);
        if (ss != null) {
            dataFileUrl = ss.getUrl();
            r.FILE_LINK = dataFileUrl;
            sheet = ss.getActiveSheet();
            rangeData = sheet.getDataRange();
            lastColumn = rangeData.getLastColumn();
            maxCols = sheet.getMaxColumns();
            lastRow = rangeData.getLastRow();
            r.TOTAL_LINES = lastRow - 1;

            for (j = 1; j <= lastColumn; j++) {
                let userColumn: string = rangeData.getCell(1, j).getValue();
                if (userColumn.trim().length > 0)
                    userColumns = `${userColumns},${userColumn}`
            }
            if (userColumns.length > 0) {
                P.StartCol = 1;
                P.EndCol = lastColumn;

            }
            if (userColumns.length == 0) {
                go = false;
                response.messages.push("Data file is empty.");
                response.domainResult = -3;

            }
            SysLog.log(`User:columns: ${userColumns}`);
            let arrUserColumns = userColumns.split(',');
            userColumns = `,${userColumns},`;
            let arrMissingColumns = new Array<string>();
            let arrFixedColumns = fixedColumns.split(',');
            for (j = 0; j < arrFixedColumns.length; j++) {
                if (userColumns.indexOf(`,${arrFixedColumns[j]},`) < 0)
                    arrMissingColumns.push(arrFixedColumns[j]);
            }

            let newCols: number = 6;
            if (userColumns.indexOf(`,${namesColumn}`) < 0)
                newCols = 7;

            //todo: find namesColumn
            if (arrMissingColumns.length == arrFixedColumns.length) {
                SysLog.log("Sheet has no standard MM columns");
                rangeData = sheet.getRange(1, 1, lastRow, lastColumn + newCols);
                newLastCol = rangeData.getLastColumn();
                maxCols = sheet.getMaxColumns();
                lastRow = rangeData.getLastRow();
                P.StartCol = 1;
                P.EndCol = lastColumn;

                i = 0;
                for (j = lastColumn + 1; j <= lastColumn + 6; j++) {
                    rangeData.getCell(1, j).setValue(arrFixedColumns[i]);

                    if (arrFixedColumns[i] == P.COL_INVALID_MAILS)
                        P.COLN_INVALID_MAILS = j;
                    else if (arrFixedColumns[i] == P.COL_INVALID_NAMES)
                        P.COLN_INVALID_NAMES = j;
                    else if (arrFixedColumns[i] == namesColumn)
                        P.COLN_NAMES = j;
                    else if (arrFixedColumns[i] == P.COL_RESOLVED_MAIL)
                        P.COLN_RESOLVED_MAIL = j;
                    else if (arrFixedColumns[i] == P.COL_RESULT)
                        P.COLN_RESULT = j;
                    else if (arrFixedColumns[i] == P.COL_STATUS)
                        P.COLN_STATUS = j;
                    else if (arrFixedColumns[i] == P.COL_TIMESTAMP)
                        P.COLN_TIMESTAMP = j;
                    i++;
                }

                SysLog.log(`
                NAMES COLUMN:${namesColumn}
                COL_INVALID_MAILS:${P.COLN_INVALID_MAILS}
                COL_INVALID_NAMES:${P.COLN_INVALID_NAMES}
                NAMES_COL:${P.COLN_NAMES}
                COL_RESOLVED_MAIL:${P.COLN_RESOLVED_MAIL}
                COL_RESULT:${P.COLN_RESULT}
                COL_STATUS:${P.COLN_STATUS}
                COL_TIMESTAMP:${P.COLN_TIMESTAMP}
                START_COL: ${P.StartCol}
                END_COL:${P.EndCol}`);


                if (newLastCol > lastColumn) {
                    lastColumn = newLastCol;
                    let found = false;
                    for (j = 1; j <= lastColumn; j++) {
                        if (namesColumn == rangeData.getCell(1, j).getValue()) {
                            P.COLN_NAMES = j;
                            found = true;
                            break;
                        }
                    }
                    if (!found) {
                        rangeData.getCell(1, lastColumn).setValue(namesColumn)
                        P.COLN_NAMES = lastColumn;
                        P.EndCol = lastColumn;
                        P.StartCol = lastColumn;
                        SysLog.log(`Names column ${namesColumn} NOT FOUND suer columns. COLN_NAMES: ${P.COLN_NAMES}`);
                    }
                }
            }
            SysLog.log(`validating columns line 402. namesCol:${namesColumn}`);
            for (j = 1; j <= lastColumn; j++) {
                colName = rangeData.getCell(1, j).getValue().trim();
                SysLog.log(`${j} colName: ${colName}`);

                if (colName == P.COL_INVALID_MAILS)
                    P.COLN_INVALID_MAILS = j;
                else if (colName == P.COL_INVALID_NAMES)
                    P.COLN_INVALID_NAMES = j;
                else if (colName == namesColumn)
                    P.COLN_NAMES = j;
                else if (colName == P.COL_RESOLVED_MAIL)
                    P.COLN_RESOLVED_MAIL = j;
                else if (colName == P.COL_RESULT)
                    P.COLN_RESULT = j;
                else if (colName == P.COL_STATUS)
                    P.COLN_STATUS = j;
                else if (colName == P.COL_TIMESTAMP)
                    P.COLN_TIMESTAMP = j;
            }
            SysLog.log(`
            NAMES COLUMN:${namesColumn}
            COL_INVALID_MAILS:${P.COLN_INVALID_MAILS}
            COL_INVALID_NAMES:${P.COLN_INVALID_NAMES}
            NAMES_COL:${P.COLN_NAMES}
            COL_RESOLVED_MAIL:${P.COLN_RESOLVED_MAIL}
            COL_RESULT:${P.COLN_RESULT}
            COL_STATUS:${P.COLN_STATUS}
            COL_TIMESTAMP:${P.COLN_TIMESTAMP}
            START_COL: ${P.StartCol}
            END_COL:${P.EndCol}`);


            r.TEMPLATE_FILE_NAME = templateFileName;

            this.mailListText = Utils.getDocTextByName(mailListFileName).toLowerCase();
            this.templateText = Utils.getDocTextByName(templateFileName);
            if (this.mailListText.length == 0) {
                let ssMailsList = Utils.openSpreadSheet(mailListFileName);
                if (ssMailsList != null) {
                    let rangeNL = sheet.getRange(1, 1, lastRow, lastColumn + newCols);
                    let lcNL = rangeData.getLastColumn();
                    let lrNL = rangeData.getLastRow();
                    this.grid = rangeNL.getValues();
                    go = go && true;
                }
            }
            else
                go = go && this.templateText.length > 0;



            //verify columns

            //validate mails
            SysLog.log("Validating mails ***********************************");
            for (i = 2; i <= lastRow; i++) {
                if (rangeData.getCell(i, P.COLN_STATUS).getValue() != "OK") {

                    rangeData.getCell(i, P.COLN_RESOLVED_MAIL).setValue("");
                    rangeData.getCell(i, P.COLN_STATUS).setValue("");
                    rangeData.getCell(i, P.COLN_RESULT).setValue("");
                    rangeData.getCell(i, P.COLN_INVALID_NAMES).setValue("");
                    rangeData.getCell(i, P.COLN_INVALID_MAILS).setValue("");
                    rangeData.getCell(i, P.COLN_TIMESTAMP).setValue(Utils.getTimeStamp());

                    let names = rangeData.getCell(i, P.COLN_NAMES).getValue();
                    mv = this.validateNames(names);
                    r.VALID_MAILS += mv.mails.length;
                    r.INVALID_EMAILS += mv.totalInvalidMails;
                    r.NOT_FOUND_NAMES += mv.totalInvalidNames;
                    if (mv.result) {
                        rangeData.getCell(i, P.COLN_RESOLVED_MAIL).setValue(mv.validMails);
                        rangeData.getCell(i, P.COLN_STATUS).setValue("");
                        //rangeData.getCell(i, P.COLN_NAMES).setValue(mv.names);
                        r.OK_LINES++;
                    }
                    else {
                        rangeData.getCell(i, P.COLN_STATUS).setValue("ERROR");
                        rangeData.getCell(i, P.COLN_RESULT).setValue(mv.error);
                        rangeData.getCell(i, P.COLN_INVALID_NAMES).setValue(mv.invalidNames);
                        rangeData.getCell(i, P.COLN_INVALID_MAILS).setValue(mv.invalidMails);
                        r.ERROR_LINES++;
                        go = false;
                        SysLog.log(`${i} ${mv.result} names: ${names} error:[${mv.error}]`);
                    }
                }
            }
        }
        else {
            response.messages.push(`Data file [${ssName}] not found.`);
            response.domainResult = -1;
        }

        //send mails
        SysLog.log(`finished validating mail. GO:${go} testMOde: ${P.testMode}`);
        SysLog.log(`OK Lines: ${r.OK_LINES} Error Lines: ${r.ERROR_LINES}`);
        go = (r.ERROR_LINES == 0);
        SysLog.log(`Startiong sending mails. GO:${go} testMOde: ${P.testMode}`);

        if (go) {
            subject = p.getParameter(P.SUBJECT);
            let lb = '{';
            let eb = '}';
            for (i = 2; i <= lastRow; i++) {
                rangeData.getCell(i, P.COLN_TIMESTAMP).setValue(Utils.getTimeStamp());
                if (rangeData.getCell(i, P.COLN_STATUS).getValue() == "") {
                    mailText = this.templateText;
                    for (j = P.StartCol; j <= P.EndCol; j++) {

                        colName = rangeData.getCell(1, j).getValue().trim().toUpperCase();
                        mailText = Utils.replace(mailText, `$${lb}${colName}${eb}`, rangeData.getCell(i, j).getValue());

                    }
                    let index = mailText.indexOf("${");
                    while (index > 0) {
                        let word = Utils.extract(mailText, "${", "}");
                        if (word.length > 0) {
                            //word = Utils.replace(word, "$", "");
                            mailText = Utils.replace(mailText, `$${word}$`, p.getParameter(word));
                        }
                        index = mailText.indexOf("${", index + word.length);
                    }
                    //todo: look for not resolved variables pointing to parameters
                    let result = 0;
                    let sendTo = rangeData.getCell(i, P.COLN_RESOLVED_MAIL).getValue();
                    if (!P.testMode) {
                        result = Utils.sendMail(sendTo, subject, mailText);
                        if (result == 0) {
                            rangeData.getCell(i, P.COLN_STATUS).setValue("OK");
                            r.MAILS_SENT++;
                        }
                        else {
                            rangeData.getCell(i, P.COLN_STATUS).setValue("ERROR");
                            rangeData.getCell(i, P.COLN_RESULT).setValue(`Error sending mail\n${Utils.ex.message}`);
                            r.MAILS_FAILED++;
                            SysLog.log(`${i} Error sending mail to: ${sendTo}\n${Utils.ex.message}`)

                        }
                    }
                    else {
                        let rdata: string = rangeData.getCell(i, P.COLN_RESULT).getValue();
                        if (rdata.trim().length == 0)
                            rangeData.getCell(i, P.COLN_RESULT).setValue("OK TEST MODE");
                        else
                            rangeData.getCell(i, P.COLN_RESULT).setValue(`${rdata}\nTEST MODE`);
                    }
                }
            }

            SysLog.log("Finishing process");

            r.PROCESSED_FILE_NAME = ssName;
            if (r.TOTAL_LINES == r.OK_LINES && r.TOTAL_LINES > 0 && r.OK_LINES > 0) {
                response.domainResult = 0;
                if (r.MAILS_FAILED > 0) {
                    response.messages.push("Some mails were not sent.");
                }
                else
                    r.OVERALL_RESULT = "SUCCESS: All lines and mails processed successfully";

                //let processedFolder = Utils.getCreateFolder(processedFolderName);
                //SysLog.log(`Processed Folder:${processedFolderName}`);

                // r.PROCESSED_FILE_NAME = `${ssName}.Processed.${r.EXECUTION_TIME}`;
                // //Move file
                // SysLog.log(`new name for data file:${r.PROCESSED_FILE_NAME}`);
                // let newFile = DriveApp.getFileById(ss.getId()).makeCopy(r.PROCESSED_FILE_NAME, processedFolder);
                // dataFileUrl = newFile.getUrl();
                // try {
                //     SysLog.log("moving file to processed");
                //     Drive.Files.remove(ss.getId());
                //     SysLog.log("file was moved");
                // }
                // catch (ex) {
                //     SysLog.log(`error deleting data file: ${ex.message}`);
                //     response.messages.push(`Data file can not be deleted. ${ex.message}`);
                // }
            }
            else {
                response.domainResult = -1;
                r.OVERALL_RESULT = "ERROR. Some lines not processed.";
            }
            response.messages.push(r.OVERALL_RESULT);

            // let logSheet = Utils.getCreateSpreadSheet(folder, `${folderName}_Log`, "TimeStamp,Total Lines,OK Lines,Error Lines,Data File url");
            // let sheetLog = logSheet.getActiveSheet();
            // let row = r.getRow();
            // r.LOG_LINK = logSheet.getUrl();
            // sheetLog.appendRow(row);

            r.TEST_MODE = P.testMode
        }
        if (!go)
            response.messages.push("Mails NOT SENT due to some lines in error. Check Data File");
        if (this.templateText.length == 0) {
            response.messages.push("Template empty or file could not be read.");
            response.domainResult = -2;
        }
        html = this.resultReport(r);
        if (response.messages.length > 0) {
            let msgs = "";
            let color = "txt-info";
            if (response.domainResult != 0)
                color = "txt-danger";
            for (i = 0; i < response.messages.length; i++) {
                msgs = `${msgs}<tr><td></td><td><p class="${color}">${response.messages[i]}</p></td></tr>`
            }
            html = html.replace("<!--MESSAGES!-->", msgs);
        }
        else html = html.replace("<!--MESSAGES!-->", "");

        response.addHtml("result", html);
        SysLog.log("mailMerge() results:", JSON.stringify(response));
        return response;
    }

    getFileInfos(dataFile, namesFile, templateFile) {
        let arr = new Array<FileInfo>();
        let fi = Utils.getFileInfo(dataFile);
        arr.push(fi);

        fi = Utils.getFileInfo(namesFile);
        arr.push(fi);

        fi.content = this.getTemplateText(templateFile);
        arr.push(fi);


        return arr;

    }



}
