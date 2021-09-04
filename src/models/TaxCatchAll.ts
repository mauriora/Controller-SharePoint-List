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
    public id: number = undefined;

    @Expose({ name: 'Term'})
    public term: string = undefined;

    @Exclude()
    public controller: SharePointList<ListItemBase> = undefined;

    public setController = (controller: SharePointList<ListItemBase>) => this.controller = controller;
}

export class TaxCatchAllFull extends TaxCatchAll
{
    constructor() {
        super();
    }
    
    @Expose({ name: 'Title'})
    public title: string = undefined;

    @Expose({ name: 'IdForTermStore' })
    public idForTermStore: string = undefined;

    @Expose({ name: 'IdForTermSet' })
    public idForTermSet: string = undefined;

    @Expose({ name: 'IdForTerm' })
    public idForTerm: string = undefined;
    
    @Expose({ name: 'Path' })
    public path: string = undefined;
}

