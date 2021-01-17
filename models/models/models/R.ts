import { Utils } from "../backend/Utils";

export class R{
     EXECUTION_TIME:string = "";
     PROCESSED_FILE_NAME: string = "";
     FILE_LINK:string = "";
     LOG_LINK: string = "":
     LINES_PROCESSED:number = 0;
     TOTAL_MAILS:number = 0;
     MAILS_SENT:number = 0;
     MAILS_FAILED:number = 0;
     VALID_MAILS: number = 0;
     INVALID_EMAILS:number = 0;
     NOT_FOUND_NAMES:number = 0;
     TOTAL_LINES: number = 0;
     OK_LINES: number = 0;
     ERROR_LINES: number = 0;
     TEMPLATE_FILE_NAME: string = "";
     OVERALL_RESULT: string = "";
     TEST_MODE:boolean = false;
     

     getRow():Array<string>
     {
          let arr = new Array<string>();
          arr.push(Utils.getTimeStamp());
          arr.push(this.TOTAL_LINES.toString());
          arr.push(this.OK_LINES.toString());
          arr.push(this.ERROR_LINES.toString());
          arr.push(this.FILE_LINK);
          return arr;
     }
}
