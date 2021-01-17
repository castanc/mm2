import { FileInfo } from "../Models/FileInfo";
import { GSResponse } from "../Models/GSResponse";
import { GSLog } from './GSLog'

export class Utils {

    static ex;
    static getCreateFolder(folderName) {
        var folders = DriveApp.getFoldersByName(folderName);
        var folder = null;
        if (folders.hasNext())
            folder = folders.next();
        else
            folder = DriveApp.createFolder(folderName);
        return folder;
    }


    static ValidateEmail(mail) {
        if (/^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/.test(mail))
            return true;
        return false;
    }

    static getTimeStamp() {
        return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss');
    }

    static getDocTextByName(fileName): string {
        var text = "";
        let doc = this.openDoc(fileName);
        if ( doc != null )
            text = doc.getBody().getText();
        return text;
    }


    static openSpreadSheet(ssName)
    {
        let ss = null;
        try {
            if (ssName.toLowerCase().indexOf("http") >= 0) {
                ss = SpreadsheetApp.openByUrl(ssName);
            }
            else {
                let fileInfos = Utils.getFilesByName(ssName);
                if (fileInfos.length > 0)
                    ss = SpreadsheetApp.openById(fileInfos[0].id);
                else
                    ss = SpreadsheetApp.openById(ssName);
            }
        }
        catch (ex) {
            ss = null;
        } 
        return ss;       
    }


    static openDoc(docName)
    {
        let ss = null;
        try {
            if (docName.toLowerCase().indexOf("http") >= 0) {
                ss = DocumentApp.openByUrl(docName);
            }
            else {
                let fileInfos = Utils.getFilesByName(docName);
                if (fileInfos.length > 0)
                    ss = DocumentApp.openById(fileInfos[0].id);
                else
                    ss = DocumentApp.openById(docName);
            }
        }
        catch (ex) {
            ss = null;
        } 
        return ss;       
    }

    static getFilesByName(name:string)
    {
        let fileInfos = new Array<FileInfo>();
        let files = DriveApp.getFilesByName(name);
        while(files.hasNext())
        {
            let file = files.next();
            //if ( !file.isTrashed)
            {
                let fi = new FileInfo(file);
                fileInfos.push(fi);
            }
        }
        return fileInfos;

    }

    static getSpreadSheet(folder, fileName) {
        let spreadSheet = null;

        let file = Utils.getFileFromFolder(fileName, folder);
        if (file != null) {
            spreadSheet = SpreadsheetApp.openById(file.getId());
        }
        return spreadSheet;
    }

    static getCreateSpreadSheet(folder, fileName, tabNames: string = "") {
        let file = Utils.getFileFromFolder(fileName, folder);
        let tabs = tabNames.split(',');
        let spreadSheet = null;

        if (file == null) {
            spreadSheet = SpreadsheetApp.create(fileName);
            if (tabs.length > 0) {
                if (tabs[0].length > 0) {
                    var sh = spreadSheet.getActiveSheet();
                    sh.setName(tabs[0]);
                }

                for (var i = 1; i < tabs.length; i++) {
                    if (tabs[i].length > 0) {
                        let itemsSheet = spreadSheet.insertSheet();
                        itemsSheet.setName(tabs[i]);
                    }
                }

            }

            var copyFile = DriveApp.getFileById(spreadSheet.getId());
            folder.addFile(copyFile);
            DriveApp.getRootFolder().removeFile(copyFile);
            file = Utils.getFileFromFolder(fileName, folder);
        }
        spreadSheet = SpreadsheetApp.openById(file.getId());
        return spreadSheet;
    }

    static getFileByName(fileName) {
        var files = DriveApp.getFilesByName(fileName);
        while (files.hasNext()) {
            var file = files.next();
            return file;
            break;
        }
        return null;
    }

    static getFileFromFolder(name: string, folder) {
        let files: FileIterator;


        files = folder.getFilesByName(name);
        if (files.hasNext()) {
            return files.next();
        }

        return null;
    }


    static getData(ss, sheetName: string): [][] {
        let sheet = ss.getSheetByName(sheetName);
        var rangeData = sheet.getDataRange();
        var lastColumn = rangeData.getLastColumn();
        var lastRow = rangeData.getLastRow();
        let grid = rangeData.getValues();
        return grid;
    }

    static getUrl(fileName, folder = null) {
        let files;
        if (folder == null)
            files = DriveApp.getFilesByName(fileName);
        else
            files = folder.getFilesByName(fileName);

        if (files.hasNext()) {
            let file = files.next();
            return file.getUrl();
        }
        return "";
    }

    static getHtmlFromArray(name: string, caption: string = "", list: Array<string>, required: boolean = false): string {
        let onChange = "";
        let requiredText = "";

        if (required)
            requiredText = "required";

        name = name.trim();

        if (caption.length == 0)
            caption = "None";

        var options = `<option value="-1" selected>${caption}</option>`;
        for (var i = 0; i < list.length; i++) {
            options = options + `<option value="${i}">${list[i]}</option>`;
        }
        onChange = `onChange="onChange_${name}('${name}',this.options[this.selectedIndex].value)"`;
        return `<select id="SELECTID" name="SELECTID" ${onChange} ${required}>${options}</select>`;

    }

    static replace(text: string, value: string, newValue: string): string {
        try {
            while (text.indexOf(value) >= 0)
                text = text.replace(value, newValue);
        }
        catch (ex) {
            //return GSLog.handleException(ex, "Utils.replace()");
        }
        return text;
    }

    static extract(text: string, start: string, end: string): string {
        let word = "";
        try {
            let index = text.indexOf(start);
            let index2 = 0;
            while (index >= 0) {
                index += start.length;
                index2 = text.indexOf(end, index);
                if (index2 > index) {
                    word = text.substr(index, index2 - index);
                }
                index = text.indexOf(start, index2 + end.length);
            }
        }
        catch (ex) {
            return GSLog.handleException(ex, "Utils.replace()");
            //return text;
        }
        return word;
    }


    static moveFiles(sourceFileId, targetFolderId) {
        try {
            let file = DriveApp.getFileById(sourceFileId);
            let folder = DriveApp.getFolderById(targetFolderId);
            file.moveTo(folder);
        }
        catch (ex) {
            Logger.log("Exception moving file.");
        }
    }

    static sendMail(to, subject, body) {
        let result = 0;
        let mails = to.split('\n');
        let mailsList = "";

        if ( mails.length == 0 )
            mailsList = to;
        else
            for(var i =0; i<mails.length; i++)
            {
                mails[i] = mails[i].trim();
                if ( mails[i].length > 0 )
                {
                    if ( mailsList.length == 0 )
                        mailsList = mails[i];
                    else 
                        mailsList = `${mailsList},${mails[i]}`;
                }
            }
        Logger.log("SendMail to " + mailsList);

        try {
            body = Utils.replace(body, "\\n", "</br>");
            MailApp.sendEmail({
                to: mailsList,
                subject: subject,
                htmlBody: body
            });
            result = 0;
        }
        catch (ex) {
            Utils.ex = ex;
            Logger.log(`Exception sending mail to [${mailsList}]\n${ex.message}\n${ex.stacktrace}`);
            result = -1;
        }
        return result;

    }

    static deleteFiles(fileName: string, folder = null) {
        let files;
        if (folder == null)
            files = DriveApp.getFilesByName(fileName);
        else
            files = folder.getFilesByName(fileName);

        while (files.hasNext()) {
            let file = files.next();
            file.setTrashed(true);
        }
    }

    static async removeFileByName(fileName) {
        var files = DriveApp.getFilesByName(fileName);
        if (files.hasNext()) {
            var file = files.next();
            file.setTrashed(true);
        }
    }

    static getTextFile(fileName: string, folder = null) {
        let file;
        if (folder == null) {
            let files = DriveApp.getFilesByName(fileName);
            if (files.hasNext())
                file = files.next()
        }
        else {
            file = Utils.getFileFromFolder(folder, fileName)

            if (file != null)
                return file.getBlob().getDataAsString();
            return "";
        }
    }

    static getTextFileFromFolder(folder, fileName: string) {
        let file = Utils.getFileFromFolder(folder, fileName)
        if (file != null)
            return file.getBlob().getDataAsString();
        return "";
    }

    static writeTextFile(fileName: string, text: string, folder = null) {
        var existing;
        if (folder == null)
            existing = DriveApp.getFilesByName(fileName);
        else
            existing = folder.getFilesByName(fileName);

        // Does file exist? if (existing.hasNext()) {

        var file = null;
        if (existing.hasNext()) {
            file = existing.next();
            file.setTrashed(true);
        }
        folder.createFile(fileName, text, MimeType.PLAIN_TEXT);
    }


}
