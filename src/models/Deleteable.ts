import { Exclude } from "class-transformer";
import { DataBase } from "./DataBase";

export interface IDeleteable {
    readonly canBeDeleted: boolean;
    deleted: boolean;
    delete: () => Promise<void>;
}


/**
 * Base class for all deletable data-entities.
 */
export abstract class Deleteable extends DataBase implements IDeleteable {

    public constructor() {
        super();
    }

    @Exclude()
    abstract readonly canBeDeleted: boolean;
    abstract delete: () => Promise<void>;
    abstract deleted: boolean;
}
