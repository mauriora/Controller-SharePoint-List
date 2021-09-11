import 'reflect-metadata';
import {
    Exclude,
    Expose
} from 'class-transformer';
import { IItem } from '@pnp/sp/items';
import { Deleteable } from './Deleteable';
import { DataBase, InitOpions } from './DataBase';
import { SharePointList } from '../controller/SharePoint/SharePointList';


// odata.editLink: "Web/Lists(guid'15fb610b-0db6-43e9-a6a1-a4a1fe7fcced')/Items(1)"
// odata.etag: "\"5\""
// odata.id: "533f110d-57d6-478b-8194-7a3e57ee1503"
// odata.type: "SP.Data.Test1ListItem"
/**
 * Minimal SharePoint ListItem inteface, extended to contain author & editor
 */
export class ListItemBase extends Deleteable {

    public constructor() {
        super();
    }

    public init(options?: InitOpions): this {
        options = DataBase.initOptions(options);
        options.nonObservableProperties.push('pnpItem')
        return super.init(options);
    }

    @Expose({ name: 'ID' })
    public id: number;

    @Expose({ name: 'Title' })
    public title: string;

    @Exclude()
    public pnpItem: IItem | undefined;

    @Exclude()
    public deleted = false;

    @Exclude()
    public get canBeDeleted(): boolean { return (undefined !== this.pnpItem?.delete); }

    @Exclude()
    public delete = async (): Promise<void> => {
        if (!this.canBeDeleted) throw new Error(`ListItemBase[${this.id}] can't be deleted`);

        await this.pnpItem.delete();
        this.deleted = true;
    }

    @Exclude()
    public controller: SharePointList;

    public setController = (controller: SharePointList): void => {this.controller = controller; }
}

export interface ListItemBaseConstructor<ListItemType extends ListItemBase = ListItemBase> {
    new(): ListItemType;
}
