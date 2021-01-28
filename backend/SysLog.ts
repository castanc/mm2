
/*
    SysLog logs messages to a Googgle Sheet file in your local Google Drive

    This is a static class, use it directly:
    SysLog.log(message,method,additional);
    All the messages are logged in Logger.log as well

    Syslog.logException(exception,method,data);
    
    cesar.alfredo@pwc.com
*/
export class SysLog {
    static folder = DriveApp.getRootFolder();
    static ssLog = null;

    constructor()
    {
       SysLog.initialize();
    }

    static initialize(){
        SysLog.ssLog = SysLog.getCreateSpreadSheet("SysLog", SysLog.folder);
        let sheet = SysLog.ssLog.getActiveSheet();
        var range = sheet.getDataRange();
        range.clearContent();

    }

    static log(msg, method = "", additional = ""){
        if ( SysLog.ssLog == null )
            SysLog.initialize();

        let ts = SysLog.getTimeStamp();
        let row = [ts,"INFO", method.trim(),msg.trim(), additional.trim()];
        SysLog.ssLog.appendRow(row);
        Logger.log(`${method} ${msg}`);
        if ( additional.length > 0)
            Logger.log(`\t${additional}`);
    }

    static logException(ex,method,data="")
    {
        if ( SysLog.ssLog == null )
            SysLog.initialize();

        let ts = SysLog.getTimeStamp();
        let row = [ts,"EXCEPTION", method.trim(),ex.message, data, ex.stackTrace];
        SysLog.ssLog.appendRow(row);
        Logger.log(`${method} ${ex.message}`);
        Logger.log(ex.stackTrace);
        if ( data.length > 0)
            Logger.log(`\t${data}`);
    }


    static getFileFromFolder(name: string, folder) {
        let files: FileIterator;


        files = folder.getFilesByName(name);
        if (files.hasNext()) {
            return files.next();
        }

        return null;
    }

    static getTimeStamp(dt: Date = null) {
        if (dt == null)
            dt = new Date();
        return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss');
    }



    /*
    Gets or createas a spreadsheet file.
    Parameters:
        folder: the folder where the file resides, null if root folder.
        tabNames: for new spreadsheets, a list of tabSheet names separated by comma.
    */
   
    static getCreateSpreadSheet(fileName, folder = null, tabNames: string = "") {
        if ( folder == null )
        {
            folder = DriveApp.getRootFolder();
        }

        let file = SysLog.getFileFromFolder(fileName, folder);
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
            file = SysLog.getFileFromFolder(fileName, folder);
        }
        spreadSheet = SpreadsheetApp.openById(file.getId());
        return spreadSheet;
    }


}
