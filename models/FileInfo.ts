import { Utils } from "../backend/Utils";

export class FileInfo {
    name: string;
    id: string;
    url: string = "";
    description: string = "";
    dateCreated: Date;
    dateModified: Date;
    parentDirs = new Array<string>();
    size: number;
    type: string;
    controlId: string;
    controlText: string;
    prefixId: string;
    content: string = "";

    constructor(file) {
        if (file != null) {
            this.dateModified = file.getLastUpdated();
            this.size = file.getSize();
            this.name = file.getName();
            this.id = file.getId();
            this.url = file.getUrl();

            var folders = file.getParents();
            while (folders.hasNext()) {
                var folder = folders.next();
                this.parentDirs.push(folder.getName());
            }
        }


    }
}
