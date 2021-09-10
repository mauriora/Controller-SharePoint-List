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

    static is = (prospect: any): prospect is MetaTerm => {
        return ('string' === typeof (prospect as MetaTerm).label) &&
        (undefined === (prospect as MetaTerm).termGuid || 'string' === typeof (prospect as MetaTerm).termGuid) &&
        (undefined === (prospect as MetaTerm).wssId || 'number' === typeof (prospect as MetaTerm).wssId);
    }
}

export interface MetaTermSP
{
    Label: string;

    TermGuid: string;

    WssId: number;
}

