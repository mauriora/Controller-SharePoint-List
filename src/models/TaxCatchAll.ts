import { Exclude, Expose } from "class-transformer";
import { SharePointList } from "../controller/SharePoint/SharePointList";
import { DataBase } from "./DataBase";
import { ListItemBase } from "./ListItemBase";

export class TaxCatchAll extends DataBase
{
    constructor() {
        super();
    }
    
    @Expose( { name: 'ID'})
    public id: number;

    @Expose({ name: 'Term'})
    public term: string;

    @Exclude()
    public controller: SharePointList<ListItemBase>;

    public setController = (controller: SharePointList<ListItemBase>) => this.controller = controller;
}

export class TaxCatchAllFull extends TaxCatchAll
{
    constructor() {
        super();
    }
    
    @Expose({ name: 'Title'})
    public title: string;

    @Expose({ name: 'IdForTermStore' })
    public idForTermStore: string;

    @Expose({ name: 'IdForTermSet' })
    public idForTermSet: string;

    @Expose({ name: 'IdForTerm' })
    public idForTerm: string;
    
    @Expose({ name: 'Path' })
    public path: string;
}

