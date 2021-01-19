import { Utils } from "../backend/Utils";

export class FileInfo{
    name:string;
    nameUrl: string="";
    id: string;
    url: string;
    description:string="";
    dateCreated: Date;
    dateModified: Date;
    //lastModified: string;
    parentDirs = new Array<string>();
    size: number;
    type: string;
    controlId: string;
    controlText: string;
    labelId: string;
    content:string = "";

    constructor( file)
    {
        this.dateModified = file.getLastUpdated();
        this.size = file.getSize();
        this.name = file.getName();
        this.id = file.getId();
        this.url = file.getUrl();

        var folders = file.getParents();
        while ( folders.hasNext())
        {
            var folder = folders.next();
            this.parentDirs.push(folder.getName());
        }


    }
}
