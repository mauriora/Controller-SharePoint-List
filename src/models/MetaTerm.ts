import { Expose } from "class-transformer";
import { DataBase } from "./DataBase";

export interface MetaTermFields {
    label: string;
    termGuid: string;
    wssId: number;
}

export class MetaTerm extends DataBase
{
    constructor(values?: MetaTermFields) {
        super();

        if(values) {
            this.label = values.label;
            this.termGuid = values.termGuid;
            this.wssId = values.wssId;
        }
    }
    
    @Expose({ name: 'Label'})
    public label: string;

    @Expose({ name: 'TermGuid'})
    public termGuid: string;

    @Expose({ name: 'WssId'})
    public wssId: number;

    static is = (prospect: unknown): prospect is MetaTerm => {
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

