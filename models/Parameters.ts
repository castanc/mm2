import {KeyValuePair} from './KeyValuePair';

export class P{
    static DATA_FILE_NAME:string = "DATA_FILE_NAME";
    static NAMES_FILE: string = "NAMES_FILE";
    static TEMPLATE_FILE_NAME =  "TEMPLATE_FILE_NAME";
    static RESULT_TEMPLATE_FILE_NAME: string = "RESULT_TEMPLATE_FILE_NAME";
    static FOLDER_NAME: string = "FOLDER_NAME";
    static OUTPUT_FOLDER_NAME: string = "OUTPUT_FOLDER_NAME";
    static COL_SELECT: string = "unselect";
    static COL_STATUS: string = "status";
    static COL_RESULT: string = "result";
    static COL_INVALID_MAILS: string = "invalidmails";
    static COL_INVALID_NAMES: string = "invalidnames";
    static COL_NAMES: string = "names";
    static COL_RESOLVED_MAIL: string = "resolvedmail";
    static COL_TIMESTAMP: string = "timestamp";
    static COL_TEMPLATE: string = "template";
    static FixedColStart = 0;
    static FixedColEnd = 0;
    static SUBJECT = "SUBJECT";

    static COLN_SELECT: number = -1;
    static COLN_STATUS: number = 1;
    static COLN_RESULT: number = 2;
    static COLN_TIMESTAMP: number = 3;
    static COLN_INVALID_MAILS: number = 4;
    static COLN_INVALID_NAMES: number = 5;
    static COLN_RESOLVED_MAIL: number = 6;
    static COLN_NAMES: number = 7;
    static StartCol = 8;
    static EndCol = 6;
    static COLN_TEMPLATE: number = -1;
    static testMode: boolean = false;
    static PARAMETERS_URL = "";
    static mainParameters: string  = "DATA_FILE_NAME,NAMES_FILE,TEMPLATE_FILE_NAME,RESULT_TEMPLATE_FILE_NAME,FOLDER_NAME,FOLDER_NAME,OUTPUT_FOLDER_NAME";

    parameters: Array<KeyValuePair<string,string>> = new Array<KeyValuePair<string,string>>();

    constructor()
    {
        this.addParameter(P.FOLDER_NAME,"MailMerge2");
        this.addParameter("COL_NAMES",P.COL_NAMES);
    }
    getParameter(key:string):string{
        //todo: not working
        //let result = this.parameters.filter(x=>x.id == key);
        //if ( result.length > 0 )
        //    return result[0].value;
        let value = "";
        for(var i=0; i< this.parameters.length; i++)
        {
            if ( this.parameters[i].id == key )
            {
                value = this.parameters[i].value;
                break;
            }
        }
        return value;
    }

    addParameter(key:string, value:string)
    {
        let kvp = new KeyValuePair<string,string>(key,value);
        let found: boolean = false;

        for(var i=0; i<this.parameters.length; i++)
        {
            if ( this.parameters[i].id == kvp.id)
            {
                this.parameters[i] = kvp;
                found = true;
                break;
            }
        }
        if ( !found )
            this.parameters.push(kvp);
    }

    getGrid():Array<Array<string>>{
        let arr = new Array<Array<string>>();
        for(var i=0;i<this.parameters.length; i++)
        {
            let row = [];
            row.push(this.parameters[i].id);
            row.push(this.parameters[i].value);
            arr.push(row);
        }
        return arr;
    }

    setParsObject(formObject)
    {
        this.addParameter("DATA_FILE_NAME",formObject.DATA_FILE_NAME)
        this.addParameter("NAMES_FILE",formObject.NAMES_FILE);
        this.addParameter("TEMPLATE_FILE_NAME", formObject.TEMPLATE_FILE_NAME);
        //this.addParameter("RESULT_TEMPLATE_FILE_NAME",formObject.RESULT_TEMPLATE_FILE_NAME);
        this.addParameter("OUTPUT_FOLDER_NAME", formObject.OUTPUT_FOLDER_NAME);
        this.addParameter("SUBJECT",formObject.SUBJECT);
        this.addParameter("SENDER_MAIL", formObject.SENDER_MAIL);
        this.addParameter("SENDER_TITLE",formObject.SENDER_TITLE);
        this.addParameter("COL_NAMES",formObject.COL_NAMES);
        //this.addParameter("STAKEHOLDERS_NAMES",formObject.STAKEHOLDERS_NAMES);
        if ( formObject.SAVE_PARAMETERS )
        {

        }
    }

    getParsObject(){
        let op = {}
        for(var i=0;i<this.parameters.length;i++)
        {
            op[this.parameters[i].id] = this.parameters[i].value;
        }
        //op[P.COL_NAMES] = P.COL_NAMES;
        op["TEST_MODE"] = P.testMode;
        op["PARAMETERS_URL"] = P.PARAMETERS_URL;
        op[P.COL_RESULT] = P.COL_RESULT;
        op[P.COL_SELECT] = P.COL_SELECT;
        op[P.COL_STATUS] = P.COL_STATUS;
        op[P.COL_INVALID_MAILS] = P.COL_INVALID_MAILS
        op[P.COL_INVALID_NAMES] = P.COL_INVALID_NAMES;
        return op;
    }
}
