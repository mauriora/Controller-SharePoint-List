import { Exclude } from "class-transformer";
import { makeObservable, observable } from "mobx";
import { Controller } from "../controller/controller";

export interface InitOpions {
    nonObservableProperties?: Array<string>;
}

export interface IDataBase {
    /** Call after all constructors have been finished.
     * @example const myData = new MyData(); myData.init();
     * @example const myData = new MyData().init();
     */
    init: (options?: InitOpions) => this;

    source: unknown;
}


/**
 * Base class for all data-entities.
 */
export class DataBase implements IDataBase {

    /** Source object this has been created from */
    @Exclude()
    public source: unknown = undefined;

    public constructor() {
    }

    protected static initOptions(options?: InitOpions): InitOpions & Required<InitOpions> {
        options = options ?? {};
        const fullOptions = { 
            ...options,
            nonObservableProperties: options.nonObservableProperties ?? new Array<string>()
        }

        return fullOptions;
    }

    @Exclude()
    private initalised = false;
    /**
     * Makes this instance observable. Needs to be called after all constructors are finished.
     * Don't call init() from inside a constructor !
     * @returns this
     */
    public init(options?: InitOpions) {

        if (this.initalised) {
            console.error(`DataBase[${this.constructor.name}].init already initialised`);
        } else {
            this.initalised = true;
            options = DataBase.initOptions(options);
            options.nonObservableProperties.push('source');

            const observableProperties: any = {};
            for (const property in this) {
                if (typeof (this[property]) === 'function') {
                    observableProperties[property] = false;
                } else if (options.nonObservableProperties.indexOf(property) >= 0) {
                    observableProperties[property] = false;
                } else {
                    observableProperties[property] = observable;
                }
            }
            console.debug(`DataBase[${this.constructor.name}].init() makeObservable`, { observableProperties, options, meNow: { ...this }, me: this });
            makeObservable(this, observableProperties);
        }
        return this;
    }
}

export interface DataBaseConstructor<DataType extends DataBase = DataBase> {
    new(): DataType;
}
