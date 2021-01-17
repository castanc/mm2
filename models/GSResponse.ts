import { KeyValuePair } from "./KeyValuePair";

export class GSResponse {
    result: number;
    message: string;
    html: Array<KeyValuePair<string,string>> = new Array<KeyValuePair<string,string>>();
    script: Array<KeyValuePair<string,string>> = new Array<KeyValuePair<string,string>>();
    error: Array<KeyValuePair<string,string>> = new Array<KeyValuePair<string,string>>();
    localRenders: Array<KeyValuePair<string,string>> = new Array<KeyValuePair<string,string>>();
    messages = new Array<string>();
    domainResult: number = 0;
    
    constructor(){
        this.result = 200;
        this.message = "";
    }

    addHtml(key:string,value:string)
    {
        this.html.push(new KeyValuePair<string,string>(key, value))
    }
    
    addScript(key:string,value:string)
    {
        this.script.push(new KeyValuePair<string,string>(key, value))
    }

    addError(key:string,value:string)
    {
        this.error.push(new KeyValuePair<string,string>(key, value))
    }

    addLocalRenders(key:string,value:string)
    {
        this.localRenders.push(new KeyValuePair<string,string>(key, value))
    }


}
