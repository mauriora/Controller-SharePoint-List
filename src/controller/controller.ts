import { ExtensionContext } from "@microsoft/sp-extension-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DataBase, DataBaseConstructor } from "../models/DataBase";

export interface Controller<ControllerType extends DataBase, DataType extends ControllerType> {
    addModel: (jsFactory: DataBaseConstructor<DataType>, filter: string) => Promise<Model<ControllerType, DataType>>;
    initialised: boolean;

    // newRecord: ControllerType;
    records: Array<DataType>;
    loadAllRecords: (filter: string) => Promise<void>;

    context: WebPartContext | ExtensionContext;

    submit: (newRecord: DataType) => Promise<void>;
    getByIdSync: (id: number) => DataType | undefined;
    getById: (id: number) => Promise<DataType>;
    getNew: () => Promise<ControllerType>;

    addGetPartial<T extends Partial<DataType> & ControllerType>(item: T): Promise<T>;
}

export interface Model<ControllerType extends DataBase, DataType extends ControllerType> {
    controller: Controller<ControllerType, DataType>;
    model: DataType;
    jsFactory: DataBaseConstructor<DataType>;

    filter: string;

    records: Array<DataType>;
    loadAllRecords: () => Promise<void>;

    newRecord: DataType;
    submit: (newRecord?: DataType) => Promise<void>;
}

