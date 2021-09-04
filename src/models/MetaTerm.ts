import { Expose } from "class-transformer";
import { DataBase } from "./DataBase";

export class MetaTerm extends DataBase
{
    constructor() {
        super();
    }
    
    @Expose({ name: 'Label'})
    public label: string = undefined;

    @Expose({ name: 'TermGuid'})
    public termGuid: string = undefined;

    @Expose({ name: 'WssId'})
    public wssId: number = undefined;
}

export interface MetaTermSP
{
    Label: string;

    TermGuid: string;

    WssId: number;
}