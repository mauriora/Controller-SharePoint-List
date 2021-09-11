import { Expose } from "class-transformer";
import { DataBase } from "./DataBase";

export class Link extends DataBase
{
    constructor() {
        super();
    }
    
    @Expose({ name: 'Description'})
    public description: string;

    @Expose({ name: 'Url'})
    public url: string;

    public static is= (prospect: unknown): prospect is Link => {
        return ('string' === typeof (prospect as Link).url) &&
        (undefined === (prospect as Link).description || 'string' === typeof (prospect as Link).description);
    }

}