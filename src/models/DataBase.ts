import { Exclude } from "class-transformer";
import {
    IObjectDidChange,
    IObservableFactory,
    makeObservable,
    observable,
    observe,
} from "mobx";

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

    /** If true the item has been modified and not submitted yet*/
    dirty: boolean;
}

/**
 * Base class for all data-entities.
 */
export class DataBase implements IDataBase {
    [key: string]: unknown;
    /** Don't allow number indexing as array */
    [key: number]: never;

    // /** Don't allow symbols indexing as array */
    //TODO Test with TypeScript 4.4 and higher
    // Causes: error TS1023: An index signature parameter type must be either 'string' or 'number'.
    //TODO Would this work with mobx?
    // [key: symbol]: never;

    /** Source object this has been created from */
    @Exclude()
    public source: unknown;

    /** If true the item has been modified and not submitted yet*/
    @Exclude()
    public dirty = false;

    protected static initOptions(
        options?: InitOpions
    ): InitOpions & Required<InitOpions> {
        options = options ?? {};
        const fullOptions = {
            ...options,
            nonObservableProperties:
                options.nonObservableProperties ?? new Array<string>(),
        };

        return fullOptions;
    }

    @Exclude()
    private onChange = (change: IObjectDidChange) => {
        if (change.name !== "dirty" && !this.dirty) {
            console.log(
                `DataBase[${this.constructor.name}].onChange(${String(
                    change.name
                )}) set dirty`
            );
            this.dirty = true;
        }
        if (change.name === "dirty" && !this.dirty) {
            console.log(
                `DataBase[${this.constructor.name}].onChange (not dirty)`
            );
        }
    };

    @Exclude()
    private initalised = false;
    /**
     * Makes this instance observable. Needs to be called after all constructors are finished.
     * Don't call init() from inside a constructor !
     * @returns this
     */
    public init(options?: InitOpions): this {
        if (this.initalised) {
            console.error(
                `DataBase[${this.constructor.name}].init already initialised`
            );
        } else {
            this.initalised = true;
            options = DataBase.initOptions(options);
            options.nonObservableProperties.push("source");

            // const observableProperties: AnnotationsMap<this, never> = {};
            const observableProperties: Record<
                keyof DataBase,
                false | IObservableFactory
            > = {};
            for (const property in this) {
                if (typeof this[property] === "function") {
                    observableProperties[property] = false;
                } else if (options.nonObservableProperties.includes(property)) {
                    observableProperties[property] = false;
                } else {
                    observableProperties[property] = observable;
                }
            }
            console.debug(
                `DataBase[${this.constructor.name}].init() makeObservable`,
                { observableProperties, options, meNow: { ...this }, me: this }
            );
            makeObservable(this, observableProperties);
            observe(this, this.onChange);
        }
        return this;
    }
}

export interface DataBaseConstructor<DataType extends DataBase = DataBase> {
    new (): DataType;
}
