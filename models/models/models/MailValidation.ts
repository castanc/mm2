export class MailValidation{
    error: string = "";
    invalidNames: string = "";
    invalidMails: string = "";
    totalInvalidNames: number = 0;
    totalInvalidMails: number = 0;
    mails = new Array<string>();
    updatedNames = new Array<string>();
    validMails: string = "";
    result: boolean = true;
}
