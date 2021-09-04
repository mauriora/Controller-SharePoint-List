import { Expose } from "class-transformer";
import { DataBase } from "./DataBase";

export class Link extends DataBase
{
    constructor() {
        super();
    }
    
    @Expose({ name: 'Description'})
    public description: string = undefined;

    @Expose({ name: 'Url'})
    public url: string = undefined;
}