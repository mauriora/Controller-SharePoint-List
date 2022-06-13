import { Exclude } from "class-transformer";
import { DataBase } from "./DataBase";

export interface IDeletable {
    readonly canBeDeleted: boolean;
    deleted: boolean;
    delete: () => Promise<void>;
}


/**
 * Base class for all deletable data-entities.
 */
export abstract class Deletable extends DataBase implements IDeletable {

    public constructor() {
        super();
    }

    @Exclude()
    abstract readonly canBeDeleted: boolean;

    @Exclude()
    abstract delete: () => Promise<void>;

    @Exclude()
    abstract deleted: boolean;
}
