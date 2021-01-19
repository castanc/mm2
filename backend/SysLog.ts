import { Utils } from "./Utils";

export class SysLog {
    folder;
    ssLog;

    constructor(folderName)
    {
        this.folder = Utils.getCreateFolder(folderName);
        this.ssLog = Utils.getCreateSpreadSheet(this.folder,"SysLogs.txt");
    }

    logMessage(msg, method = "", additional = ""){
        let ts = Utils.getTimeStamp();
        let row = [ts,method,msg, additional];
        this.ssLog.appendRow(row);
    }
}
