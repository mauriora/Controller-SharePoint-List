// causes webpack to fail -> import { defaultMetadataStorage } from 'class-transformer/types/storage';
// declared in ./MetadataStorage.d.ts
import { defaultMetadataStorage } from "class-transformer/esm5/storage";

// import { FieldTypes, IFieldInfo } from "@pnp/sp/fields";
// import { IList, IListInfo } from "@pnp/sp/lists";
import { IFieldInfo, IList, IListInfo, FieldTypes } from "@pnp/sp/presets/all";
import {
    ListItemBase,
    ListItemBaseConstructor,
} from "../../models/ListItemBase";
import { Controller } from "../controller";
import { SharePointModel } from "./Model";
import { getDefaultSite, getSite, normaliseSiteUrl, SiteInfo } from "./Site";
import { makeAutoObservable, when } from "mobx";
import { UserLookup } from "../../models/User";
import {
    DeferredContainer,
    fixSingleTaxonomyFields,
    parseImages,
    removeNullValues,
    resultArrayToArray,
    setNullArrays,
    toSubmit,
} from "./Transformer";
import { ClassTransformOptions, plainToClass, plainToClassFromExist } from "class-transformer";
import { ODataError } from "../../models/OData/Error";
import { DataError } from "../../models/DataError";
import { allowsMultipleValues, getLookupList } from "./FieldInfo";
import { ListItem } from "../..";
import { BaseComponentContext } from "@microsoft/sp-component-base";

const TAX_CATCH_ALL_FIELD = "TaxCatchAll";

const LIST_PATH = "Lists";
/**
 * Returns the full URL for a list.
 * If siteUrl is undefined, then the default URL from the default Context is used.
 * If siteUrl is given, then the listUrl will be siteUrl + /Lists/ + listId
 * @param listId internal name of list as shown in the URL
 * @param siteUrl optional siteUrl if not default context site
 * @returns /sites/SiteName/Lists/ListId
 */
const getListUrl = (listId: string, siteUrl: string): string => {
    const url = normaliseSiteUrl(siteUrl ?? getDefaultSite().url);

    return `${url}/${LIST_PATH}/${listId}`;
};

export const getLists = async (siteUrl?: string): Promise<Array<IListInfo>> => {
    const normalisedSiteUrl =
        siteUrl === undefined ? undefined : normaliseSiteUrl(siteUrl);
    const site: SiteInfo = siteUrl
        ? await getSite(normalisedSiteUrl)
        : getDefaultSite();

    if (undefined === site.lists) {
        site.lists = await site.web.lists.orderBy("Title")();
    }

    return site.lists;
};

interface LookupFieldMapping {
    listId: string;
    loadLookupController: boolean;
    thisFieldName: string;
    controller: SharePointList;
}

export class SharePointList<DataType extends ListItemBase = ListItemBase>
    implements Controller<ListItemBase, DataType>
{
    /**
     * The most specialist model used to create Objects from
     */
    private baseModel: SharePointModel<DataType>;

    /**
     * Items created through lookup, not necessarily containing all model fields.
     */
    public partialItems = new Array<Partial<DataType> & ListItemBase>();
    public listId: string;
    public listTitle: string;
    public listInfo: IListInfo;
    public votingExperience: undefined | "Ratings" | "Likes";

    public rootFolderProperties: Record<string, string | number>;
    public context: BaseComponentContext;
    public initialised: boolean;

    /** If newRecord is submitted, then it will be replaced with a new instance. */
    public newRecord: DataType;

    public records: Array<DataType> = new Array<DataType>();

    public getByIdSync = (id: number): DataType | undefined =>
        this.records.find(item => item.id === id);


    private addToRecords = (item: DataType) => {
        if (this.records.includes(item)) {
            throw new Error(`SharePointList[${this.getName()}].addToRecords(id= ${item?.id}) already includes item`);
        } else if (this.records.some(prospect => item.id === prospect.id)) {
            throw new Error(`SharePointList[${this.getName()}].addToRecords(id= ${item?.id}) different instance with same id already exists`);
        }
        this.records.push(item);
        when(() => item.deleted).then(() => this.removeItem(item));
    }


    public getById = async (id: number): Promise<DataType> => {
        const local = this.getByIdSync(id);
        if (undefined !== local) return local;

        try {
            const plain = await this.getPlainById(id);
            const existing = this.getRecord(id);
            const instance = await this.getObject(plain, existing);

            if (!existing) {
                this.addToRecords(instance);
            }
            console.debug(
                `SharePointList[${this.getName()}].getById(${id})`,
                { plain, existing, instance, me: this }
            );
            return instance;
        } catch (getItemsError: unknown) {
            throw new Error(`SharePointList[${this.getName()}].getById(${id}) failed: ${(getItemsError as Error).message ?? getItemsError}`);
        }
    };

    /**
     * Submits a record
     * If jsRecord is not specified it submits this.newReocrd.
     * !! If this.newRecord is submitted, then it will be replaced with a new instance. !!
     * @param jsRecord is .id is undefined or <0, then a new item is created, filled with the returned list item values (id, default values, ...)
     */
    public submit = async (jsRecord?: DataType): Promise<void> => {
        jsRecord = jsRecord ?? this.newRecord;

        const submitRecord = await toSubmit(
            jsRecord,
            this.selectedFields,
            this.allFields
        );

        if (submitRecord.ID && submitRecord.ID > 0) {
            const updateResponse = await jsRecord.pnpItem.update(submitRecord);

            console.log(`SharePointList[${this.getName()}].submit() update response.data.odata.etag=${updateResponse.data["odata.etag"]}`,
                { jsRecord, submitRecord, updateResponse }
            );
        } else {
            const createResponse = await this.list.items.add(submitRecord);
            const newID = createResponse.data.ID;
            jsRecord = await this.getObject(createResponse.data, jsRecord);
            this.addToRecords(jsRecord);

            if (this.newRecord === jsRecord) {
                for (const model of this.models.values()) {
                    model.newRecord = this.newRecord;
                }
            }
            console.log(`SharePointList[${this.getName()}].submit() add response=${newID}`, { jsRecord, submitRecord, createResponse });
        }
        jsRecord.dirty = false;
    };

    protected constructor(public site: SiteInfo, listIdorTitle: string) {
        if (
            /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-5][0-9a-f]{3}-[089ab][0-9a-f]{3}-[0-9a-f]{12}\}?$/i.test(
                listIdorTitle
            )
        ) {
            this.listId = listIdorTitle;
            this.listTitle = undefined;
        } else {
            this.listId = undefined;
            this.listTitle = listIdorTitle;
        }

        this.context = site.isDefault
            ? site.context
            : ({
                pageContext: {
                    web: {
                        absoluteUrl: site.url,
                        language:
                            getDefaultSite().context.pageContext.web.language,
                    },
                },
                spHttpClient: getDefaultSite().context.spHttpClient,
            } as BaseComponentContext);

        makeAutoObservable(this, {});
    }

    private list: IList;
    /** all fieldInfos mapped by fieldName/internalName */
    private _allFields = new Map<string, IFieldInfo>();
    public get allFields(): Map<string, IFieldInfo> {
        if (0 === this._allFields.size) {
            throw new Error(
                `SharePointList[${this.getName()
                }].get allFields() not fields. Call getFieldInfos first`
            );
        }
        return this._allFields;
    }

    /** SharePoint internalname mapped to fieldInfos */
    public selectedFields = new Map<string, IFieldInfo>();

    /**
     * @deprecated Please remove the call
     */
    public init = async (): Promise<void> => {
        console.warn(`Call to controller/SharePointList.init() is deprecated. Please remove the call.`);
    }

    /**
     * Gets this.list, fieldInfos, selects all fields, expands all Person/User fields.
     * Call this for controllers interacting with sharepoint.
     * Not needed if this controller is created internally as a lookup controller from expanded fields.
     */
    private initialise = async (): Promise<this> => {
        if (this.initialised) {
            throw new Error(`SharePointList[${this.getName()}].init() already done voting=${this.votingExperience}`);
        }

        try {
            if (this.listId) {
                this.list = this.site.sp.web.lists.getById(this.listId);
            } else {
                this.list = this.site.sp.web.lists.getByTitle(this.listTitle);
            }
        } catch (getListError: unknown) {
            throw new Error(
                `SharePointList[${this.getName()
                }].init(): ${this.listId ? "getById" : "getByTitle"}: ${(getListError as Error).message ?? getListError
                }`
            );
        }

        await Promise.all([
            this.getListInfo(),
            this.getFieldInfos(),
            this.getRootFolderProperties(),
        ]);

        await this.addAllSelectAndExpands();

        this.initialised = true;
        console.log(`SharePointList[${this.getName()}].init() done voting=${this.votingExperience}`, {me: this});
        return this;
    };

    private async getFieldInfos() {
        try {
            const fieldsInfos = await this.list.fields();

            for (const info of fieldsInfos) {
                this._allFields.set(info.InternalName, info);
            }
        } catch (getFieldsError: unknown) {
            throw new Error(
                `SharePointList[${this.getName()
                }].getFieldInfos() caught: ${(getFieldsError as Error).message ?? getFieldsError
                }`
            );
        }
    }

    private getListInfo = async () => {
        try {
            this.listInfo = await this.list();
        } catch (getListInfoError: unknown) {
            throw new Error(
                `SharePointList[${this.getName()}].getListInfo(): getList ${(getListInfoError as Error).message ?? getListInfoError}`
            );
        }
    };

    /**
     * gets RootFolderProperties to set votingExperience to undefined, Ratings or Likes
     */
    private getRootFolderProperties = async () => {
        try {
            this.rootFolderProperties =
                await this.list.rootFolder.properties.get();
            this.votingExperience = this.rootFolderProperties["Ratings_x005f_VotingExperience"] as undefined | "Ratings" | "Likes";
        } catch (getError: unknown) {
            throw new Error(
                `SharePointList[${this.getName()}].getRootFolderProperties(): getList ${(getError as Error).message ?? getError}`
            );
        }
    };

    private addAllSelectAndExpands = async () => {
        for (const model of this.models.values()) {
            await this.addModelSelectAndExpand(model);
        }
    };

    /** fields for $select part of the query */
    private selectFields = new Set<string>();

    /** fields for $expand part of the query */
    private expandFields = new Set<string>();

    /** Propertyname mapped to Sharepoint fieldInfos */
    private propertyFields = new Map<keyof DataType, IFieldInfo>();

    private models = new Map<ListItemBaseConstructor, SharePointModel<DataType>>();

    /**
     * Add the model to selectFields, selectedFields, expandFields and possible set as baseModel
     * @param model to merge in this
     */
    private addModelSelectAndExpand = async (model: SharePointModel<DataType>) => {
        for (const selectField of model.selectFields) {
            this.selectFields.add(selectField);
        }
        for (const [fieldName, fieldInfo] of model.selectedFields) {
            this.selectedFields.set(fieldName, fieldInfo);
        }
        for (const [propertyName, fieldInfo] of model.propertyFields) {
            this.propertyFields.set(propertyName, fieldInfo);
        }
        for (const expandField of model.expandFields) {
            this.expandFields.add(expandField);
        }
        if (
            undefined === this.baseModel ||
            this.baseModel.selectFields.length < model.selectFields.length
        ) {
            this.baseModel = model;
            model.newRecord = this.newRecord = await this.getObject();
        }
    };

    private getName = () => this?.listInfo?.Title ?? this.listId ?? this.listTitle;

    async addModel<ModelType extends DataType>(
        jsFactory: ListItemBaseConstructor<ModelType>,
        filter: string
    ): Promise<SharePointModel<ModelType>> {
        const existing = this.models.get(jsFactory) as unknown as SharePointModel<ModelType>;

        if (existing) {
            return existing;
        }
        const newModel = new SharePointModel(
            jsFactory,
            this as unknown as SharePointList<ModelType>,
            filter
        );
        newModel.records = this.records as Array<ModelType>;
        this.models.set(
            jsFactory,
            newModel as unknown as SharePointModel<DataType>
        );

        if (this.initialised) {
            await this.addModelSelectAndExpand(
                newModel as unknown as SharePointModel<DataType>
            );
            await this.createRequiredChildController();
        }

        return newModel;
    }

    private static EXPANDABLE_FIELD_TYPES: FieldTypes[] = [FieldTypes.Number, FieldTypes.Text, FieldTypes.DateTime];

    /**
     * Maps propertyname to { Lookup list Id, loadLookupController, thisFieldName }
     */
    private lookupMappings = new Map<keyof DataType, LookupFieldMapping>();

    private static isNotExpandableField = (field: IFieldInfo) =>
        allowsMultipleValues(field) || (!SharePointList.EXPANDABLE_FIELD_TYPES.includes(field.FieldTypeKind));

    private static isNotExpandable = (lookupController: SharePointList): boolean =>
        Array.from(lookupController.selectedFields.values())
            .some(field => SharePointList.isNotExpandableField(field));

    /**
     * Change the $select to only include lookup information (ID + (Title or Term)) 
     * The caller needs to call lookupController.startLoadPartials
     * @param fieldName the internal/static fieldname (not property name)
    */
    private changeExpandToLookup = (fieldName: string) => {
        // Remove all expansions of the fieldname
        const filteredSelects = Array.from(this.selectFields.values()).filter(
            selectField => !selectField.startsWith(fieldName + "/")
        );
        const removedSomething = filteredSelects.length !== this.selectFields.size;
        // Add required lookup expansion
        filteredSelects.push(
            ...(TAX_CATCH_ALL_FIELD === fieldName
                ? [fieldName + "/Term", fieldName + "/ID"]
                : [fieldName + "/Title", fieldName + "/ID"])
        );

        if (removedSomething) {
            this.selectFields.clear();
            filteredSelects.forEach(filteredSelect =>
                this.selectFields.add(filteredSelect)
            );
        } else {
            throw new Error(
                `SharePointList[${this.getName()}].changeExpandToLookup FAILED to remove ${fieldName} from [${Array.from(
                    this.selectFields.values()
                ).join(", ")}]`
            );
        }
    }

    private createRequiredChildController = async () => {
        for (const [property, info] of this.propertyFields.entries()) {
            if ("string" === typeof property) {
                switch (info.FieldTypeKind) {
                    case FieldTypes.User:
                    case FieldTypes.Lookup:
                        {
                            const lookUpListId = getLookupList(info);

                            if (false === lookUpListId) {
                                throw new Error(
                                    `SharePointList[${this.getName()}].createRequiredChildController no LookupList (ID) for ${property} of type ${info.TypeAsString}[${info.FieldTypeKind}]`
                                );
                            } else if (!this.lookupMappings.has(property)) {
                                const lookupController = await getCreateById(lookUpListId, this.site.url);

                                // Add Lookup models of all registered models
                                for (const model of this.models.values()) {
                                    this.addLookupModeltoController(
                                        model,
                                        property,
                                        lookupController
                                    );
                                }

                                const isNotExpandable = lookupController.autoLoadPartials || SharePointList.isNotExpandable(lookupController);

                                if (isNotExpandable) {
                                    lookupController.startLoadPartials();
                                    this.changeExpandToLookup(info.InternalName);
                                }
                                this.lookupMappings.set(property, {
                                    listId: lookUpListId,
                                    loadLookupController: isNotExpandable,
                                    thisFieldName: info.InternalName,
                                    controller: lookupController,
                                });
                            }
                        }
                        break;
                }
            } else {
                throw new Error(`Can only have string members not ${typeof property}`);
            }
        }
    };

    private addLookupModeltoController(
        model: SharePointModel<DataType>,
        property: string,
        lookupController: SharePointList<ListItemBase>
    ) {
        const typeData = defaultMetadataStorage.findTypeMetadata(
            model.jsFactoryFactory(),
            property
        );

        const lookupFactory = typeData.typeFunction() as ListItemBaseConstructor;

        if (!lookupController.models.has(lookupFactory)) {
            lookupController.addModel(lookupFactory, "");
        }
    }

    /** Called with a partial instance from an expanded Lookup. If the item already exist, return the existing one,
     * otherwise use this one.
     * Used to "resolve" intiaially created ChildItems
     **/
    public async addGetPartial<T extends Partial<DataType> & ListItemBase>(
        item: T
    ): Promise<T> {
        const existing = this.getPartial(item.id);
        if (existing) {
            return existing as T;
        }
        if ("function" === typeof item.setController) {
            item.setController(this as unknown as SharePointList);
        } else {
            console.error(
                `SharePointList[${this.getName()}].addGetPartial item has no setController`,
                { item, controller: this }
            );
        }

        this.partialItems.push(item);
        if (this.autoLoadPartials) {
            await this.loadAndShiftPartials();
        }
        return item;
    }

    public getRecord = (id: number): undefined | DataType =>
        this.records.find(prospect => prospect.id === id);

    public getPartial = (id: number): undefined | (Partial<DataType> & ListItemBase) =>
        this.getRecord(id) ||
        this.partialItems.find(prospect => prospect.id === id);

    private autoLoadPartials = false;

    private startLoadPartials = () => {
        if (!this.autoLoadPartials) {
            this.autoLoadPartials = true;
            this.loadAndShiftPartials();
        }
    };

    private loadingAndShiftingPartials = false;

    private loadAndShiftPartials = async () => {
        if (!this.loadingAndShiftingPartials) {
            this.loadingAndShiftingPartials = true;

            try {
                while (this.partialItems.length) {
                    const partial = this.partialItems[0];
                    const partialAsFull = await this.loadFull(partial);
                    this.partialItems.splice(0, 1);
                    this.addToRecords(partialAsFull);
                }
            } finally {
                this.loadingAndShiftingPartials = false;
            }
        }
    };

    private loadFull = async (partial: Partial<DataType> & ListItemBase) =>
        this.getObject(await this.getPlainById(partial.id), partial);

    private static EXPANDED_FIELD_ERROR_500 = RegExp(
        /^Cannot get value for projected field ([A-z0-9_]+)_x005f_([A-z0-9_]+).$/
    );

    private static isExpanededFieldError500 = (error: {
        message?: string;
    }): string | undefined => {
        if (error.message && /error/.test(error.message)) {
            try {
                const errorObject = JSON.parse(
                    (error.message as string).substring(
                        (error.message as string).indexOf("{")
                    )
                );
                const odataError: ODataError = errorObject;
                const dataError: DataError = errorObject;
                const content = odataError?.["odata.error"] ?? dataError.error;

                const errorMatches =
                    SharePointList.EXPANDED_FIELD_ERROR_500.exec(
                        content?.message?.value
                    );

                if (
                    content?.code ===
                    "-2146232832, Microsoft.SharePoint.SPException" &&
                    errorMatches.length >= 2
                ) {
                    return errorMatches[1] + "/" + errorMatches[2];
                }
            } catch (parseErrorError: unknown) {
                throw new Error(
                    `SharePointList.isExpanededFieldError500 JSON parse error: ${(parseErrorError as Error).message ?? parseErrorError
                    }: parsing error: ${error?.message ?? error}`
                );
            }
        } else {
            console.debug("SharePointList.isExpanededFieldError500 no match", {
                error,
            });
        }
        return undefined;
    };

    private static EXPANDED_FIELD_ERROR_400 = RegExp(
        /^The query to field '([A-z0-9_]+)\/([A-z0-9_]+)' is not valid.$/
    );

    private static isExpanededFieldError400 = (error: {
        message?: string;
    }): string | undefined => {
        // if(error.message && /\[500\]\s+::>\{"odata.error"/.test(error.message))  {
        if (error.message && /odata.error/.test(error.message)) {
            try {
                const odataError: ODataError = JSON.parse(
                    (error.message as string).substring(
                        (error.message as string).indexOf("{")
                    )
                );
                const errorMatches =
                    SharePointList.EXPANDED_FIELD_ERROR_400.exec(
                        odataError?.["odata.error"]?.message?.value
                    );

                console.debug(
                    `SharePointList.isExpanededFieldError400 parsed`,
                    { error, odataError, errorMatches }
                );

                if (odataError?.["odata.error"]?.code === "-1, Microsoft.SharePoint.SPException" &&
                    errorMatches.length >= 2
                ) {
                    return errorMatches[1] + "/" + errorMatches[2];
                }
            } catch (parseErrorError: unknown) {
                throw new Error(
                    `SharePointList.isExpanededFieldError400 JSON parse error: ${(parseErrorError as Error).message ?? parseErrorError
                    }: parsing error: ${error?.message ?? error}`
                );
            }
        } else {
            console.debug("SharePointList.isExpanededFieldError400 no match", { error });
        }
        return undefined;
    };

    private handleFailedExpand = (expandedField: string) => {
        const fieldName = expandedField.substring(
            0,
            expandedField.indexOf("/")
        );
        // this.changeExpandToLookup(fieldName);
        const mappingsArray = Array.from(this.lookupMappings.values() ?? []);
        const mapping = mappingsArray.find(
            (prospect) => prospect.thisFieldName === fieldName
        );

        if (undefined === mapping) {
            throw new Error(`SharePointList[${this.getName()}].handleFailedExpand can't find mapping for expandedField ${fieldName}`);
        } else {
            if (!mapping.loadLookupController) {
                const controller = getById(mapping.listId) as SharePointList;

                controller.startLoadPartials();
            } else {
                console.warn(`SharePointList[${this.getName()}].handleFailedExpand(${fieldName}) already loaded, ignore this`);
            }

            const sameControllers = mappingsArray.filter((prospect) => prospect.listId === mapping.listId);
            console.log(`SharePointList[${this.getName()}].handleFailedExpand(${fieldName}) found ${sameControllers.length} same mappings`);

            for (const sameControllerMapping of sameControllers) {
                sameControllerMapping.loadLookupController = true;
                this.changeExpandToLookup(sameControllerMapping.thisFieldName);
            }
        }
    };

    public loadAllRecords = async (filter: string): Promise<void> => {
        try {
            const plainItems = await this.list.items
                .filter(filter)
                .select(...this.selectFields)
                .expand(...this.expandFields)
                .get();

            for (const plain of plainItems) {
                const existing = this.getRecord(plain.ID);
                const instance = await this.getObject(plain, existing);

                if (!existing) {
                    this.addToRecords(instance);
                }
            }
            console.debug(
                `SharePointList[${this.getName()}].getAll gotInstances records=${this.records.length}`,
                { plainItems, records: [...this.records], me: this }
            );
        } catch (getItemsError: unknown) {
            let failedExpandedField = undefined;
            const status = (getItemsError as Record<string, number>)["status"]
            switch (status) {
                case 404:
                    failedExpandedField = TAX_CATCH_ALL_FIELD + "/Title";
                    break;
                case 400:
                    failedExpandedField = SharePointList.isExpanededFieldError400(getItemsError);
                    break;
                case 500:
                    failedExpandedField = SharePointList.isExpanededFieldError500(getItemsError);
                    break;
            }
            if (undefined !== failedExpandedField) {
                switch (status) {
                    case 404:
                        failedExpandedField = TAX_CATCH_ALL_FIELD + "/Title";
                        break;
                    case 400:
                        console.error(
                            `SharePointList[${this.getName()}].loadAllRecords() failed because of expanding ${failedExpandedField}, this should not happen. Please report and then ignore this error`,
                            {
                                selectFields: [...this.selectFields],
                                lookupMappings: this.lookupMappings,
                                lookupMappingsCount: this.lookupMappings.size,
                                lookupMappingsArray: Array.from(
                                    this.lookupMappings.entries()
                                ),
                            }
                        );
                        break;
                    case 500:
                        console.warn(
                            `SharePointList[${this.getName()}].loadAllRecords() failed because of expanding ${failedExpandedField}, this is because the field doesn't contain a value. Getting these lookup items directly now, (instead of expand)`,
                            {
                                selectFields: [...this.selectFields],
                                lookupMappings: this.lookupMappings,
                                lookupMappingsCount: this.lookupMappings.size,
                                lookupMappingsArray: Array.from(
                                    this.lookupMappings.entries()
                                ),
                            }
                        );
                        failedExpandedField = SharePointList.isExpanededFieldError500(getItemsError);
                        break;
                }
                this.handleFailedExpand(failedExpandedField);
                // console.warn(`SharePointList[${this.getName()}].getAll removing ${failedExpandedField}`, {getItemsError, failedExpandedField, selectFields: [...this.selectFields]});
                console.warn(
                    `SharePointList[${this.getName()
                    }].getAll removed ${failedExpandedField} and try again`,
                    {
                        getItemsError,
                        failedExpandedField,
                        selectFields: [...this.selectFields],
                    }
                );
                await this.loadAllRecords(filter);
            } else {
                console.error(
                    `SharePointList[${this.getName()}].getAll failed (${status}) allItmes=${this.records.length}: ${(getItemsError as Error).message ?? getItemsError}`,
                    {
                        status,
                        getItemsError,
                        factory: this.baseModel.jsFactoryFactory(),
                        selects: this.selectFields,
                        expands: this.expandFields,
                    }
                );
                throw new Error(
                    `SharePointList[${this.getName()}].getAll failed: ${(getItemsError as Error).message ?? getItemsError}`
                );
            }
        }
    };

    private getPlainById = async (id: number) =>
        (
            await this.list.items
                .select(...this.selectFields)
                .expand(...this.expandFields)
                .filter(`ID eq ${id}`)
                .get()
        )[0];

    /**
     * Returns an initialised instance
     * @returns new object created by the factory, init called.
     * Checks that arrays are created, and call instance.init()
     */
    public getNew = async (): Promise<DataType> => this.getObject();

    /**
     * Returns an initialised instance. Could be filled from plain, into a new or existing partial ClassItem instance.
     * @param plain values to convert
     * @param existing instance to fill and initialise
     * @returns the existing instance or a new object created from plain or the factory.
     * Checks that arrays are created.
     * If filled with values from plain then removeNullValues, parseImages, connectsLookup and fixes single Taxonomy fields.
     * Connect pnpItem, calls instance.init() and removes it from records when deleted
     */
    private getObject = async (plain?: Record<string, unknown>, existing?: Partial<DataType> & ListItemBase): Promise<DataType> => {
        if (plain) {
            resultArrayToArray(plain, this.selectedFields);
        }
        let instance: DataType = undefined;
        let newInstance: DataType = undefined;

        if (plain) {
            const transformOptions: ClassTransformOptions = { excludeExtraneousValues: true, exposeDefaultValues: true };
            removeNullValues(plain);
            parseImages(plain, this.selectedFields);
            if (existing) {
                instance = plainToClassFromExist(existing, plain, transformOptions) as DataType;
            } else {
                newInstance = instance = plainToClass(this.baseModel.jsFactoryFactory(), plain, transformOptions);
            }
        } else if (existing as DataType) {
            instance = existing as DataType;
        } else {
            newInstance = instance = new (this.baseModel.jsFactoryFactory())();
        }

        instance.source = plain;

        setNullArrays(instance, this.propertyFields);
        if (plain) {
            await this.connectLookUp(instance);
            if (instance instanceof ListItem)
                fixSingleTaxonomyFields(instance, this.propertyFields);
        }
        if (instance.setController) {
            instance.setController(this as unknown as SharePointList);
        }
        if (instance.id && !instance.pnpItem) {
            instance.pnpItem = this.list.items.getById(instance.id);
        }
        if (newInstance) {
            newInstance.init();
        }

        return instance;
    };

    private connectLookUp = async (item: DataType) => {
        for (const [property, mappingInfo] of this.lookupMappings) {
            const tempLookup = item[property] as ListItemBase;
            if (undefined !== tempLookup) {
                const value = (item.source as DeferredContainer)[mappingInfo.thisFieldName];
                if (value && "__deferred" in value && value.__deferred &&
                    "id" in tempLookup && undefined === tempLookup.id
                ) {
                    console.warn(
                        `SharePointList[${this?.listInfo?.Title ??
                        this.listId ??
                        this.listTitle
                        }].connectLookUp .${property}=${tempLookup} don't know what to do with deferred, set ${property}=undefined !!`,
                        { item, property, tempLookup, mappingInfo }
                    );
                    const fieldInfo = this.propertyFields.get(property);

                    if (allowsMultipleValues(fieldInfo)) {
                        console.error(
                            `SharePointList[${this?.listInfo?.Title ??
                            this.listId ??
                            this.listTitle
                            }].connectLookUp .${property}=${tempLookup} don't know what to do with deferred, set ${property}=Array THIS SHOULD BE DONE BY setNullArrays !!`,
                            { item, property, tempLookup, mappingInfo }
                        );
                        (item[property] as unknown) = new Array<ListItem>();
                    } else {
                        item[property] = undefined;
                    }
                } else {
                    const isArray = Array.isArray(tempLookup);
                    const lookupController = getById(
                        mappingInfo.listId
                    ) as SharePointList;
                    const lookupItems = isArray
                        ? (tempLookup as unknown as Array<ListItem>)
                        : [tempLookup as ListItem];

                    for (let index = 0; index < lookupItems.length; index++) {
                        const lookup = lookupItems[index];
                        if (undefined === lookup.id) {
                            throw new Error(
                                `SharePointList[${this?.listInfo?.Title ??
                                this.listId ??
                                this.listTitle
                                }].connectLookUp(${item?.id}).${property} no Id`
                            );
                        } else {
                            const lookupItem =
                                await lookupController.addGetPartial(lookup);
                            if (undefined === lookupItem) {
                                throw new Error(
                                    `SharePointList[${this?.listInfo?.Title ??
                                    this.listId ??
                                    this.listTitle
                                    }].connectLookUp(${item?.id
                                    }) .${property}[${lookup.id
                                    }] problem addGetPartial`
                                );
                            } else {
                                if (isArray) {
                                    (item[property] as Array<ListItem>)[index] = lookupItem;
                                } else {
                                    (item[property] as unknown) = lookupItem;
                                }
                            }
                        }
                    }
                }
            }
        }
    };

    private removeItem = (item: DataType) => {
        const index = this.records.indexOf(item);
        if (-1 == index) {
            throw new Error(`SharePointList.removeItem( item.id = ${item?.id} ) not found`);
        } else {
            this.records.splice(index, 1);
        }
    };

    public static create = async (site: SiteInfo, listIdorTitle: string): Promise<SharePointList> =>
        new SharePointList(site, listIdorTitle).initialise();
}

/**
 * Maps a full listUrls to a list controller
 */
const controllers = new Map<string, SharePointList>();

/**
 * Creates a new instance of a list controller.
 * If siteUrl is undefined, then a controller is created for the default context.
 * If siteUrl is set, then an isolated context is created for the list.
 * If this function has been called with the same parameters, a warning is shown
 * and the previous created controller is returned.
 * @param listIdOrTitle internal name of the list as in the URL
 * @param siteUrl optional, specify if accessing a list on a different site than the webpart is running on.
 * @returns instance of the new List Controller
 */
export const create = async (
    listIdOrTitle: string,
    siteUrl?: string
): Promise<SharePointList> => {
    const url = getListUrl(listIdOrTitle, siteUrl);

    const existingController = controllers.get(url) as SharePointList;

    if (undefined !== existingController) {
        console.warn(`SharePointList:create(${url}): already exist`);
        return existingController;
    }

    try {
        const site =
            siteUrl === undefined
                ? getDefaultSite()
                : await getSite(normaliseSiteUrl(siteUrl));
        console.log(`SharePointList:create ( ${url} )`, {
            site,
            listNameOrId: listIdOrTitle,
            siteUrl,
        });

        const newController = await SharePointList.create(site, listIdOrTitle);
        controllers.set(url, newController);

        return newController;
    } catch (initError: unknown) {
        throw new Error(
            `SharePointList:create( ${url} ) failed: ${(initError as Error).message ?? initError
            }`
        );
    }
};

/**
 * Returns a previously created list controller.
 * @throws if no controller has been created prior for the listId and siteUrl
 * @param listIdOrName internal name of the list as in the URL
 * @param siteUrl optional, specify if accessing a list on a different site than the site the webpart is running on.
 * @returns
 */
export const getByUrl = (
    listIdOrName: string,
    siteUrl: string
): SharePointList => {
    const url = getListUrl(listIdOrName, siteUrl);

    const controller = controllers.get(url);

    if (undefined === controller) {
        throw new Error(
            `SharePointList:getByUrl: Can't find controller for ${listIdOrName} at ${siteUrl} = ${url}. Call create first. Got ${controllers.size
            } controllers: ${new Array(controllers.keys()).join(" ; ")}`
        );
    }

    return controller;
};

/**
 * Returns a previously created list controller by list id (Guid).
 * This is used to get the controller of LookUps
 * @throws If throwExceptionIfDoesntExist and no controller has been created prior for the list Guid
 * @param listGuid
 * @returns
 */
export const getById = (
    listGuid: string,
    throwExceptionIfDoesntExist = true
): SharePointList => {
    const controller = Array.from(controllers.values()).find(
        prospect => prospect.listId === listGuid
    );

    if (throwExceptionIfDoesntExist && undefined === controller) {
        throw new Error(
            `SharePointList:get: Can't find controller for ${listGuid}.Got ${controllers.size
            } controllers: ${Array.from(controllers.values())
                .map(item => item.listId ?? item.listTitle)
                .join("; ")}`
        );
    }

    return controller;
};

export const getCreateById = async (
    listId: string,
    siteUrl?: string
): Promise<SharePointList> => {
    let controller = getById(listId, false);

    if (undefined === controller) {
        controller = await create(listId, siteUrl);
    }
    return controller;
};

export const getCreateByIdOrTitle = async (listId: string, siteUrl?: string): Promise<SharePointList> => {
    const url = getListUrl(listId, siteUrl);
    const existingController = controllers.get(url) as SharePointList;

    if (undefined !== existingController) {
        return existingController;
    }
    return create(listId, siteUrl);
}

export const getUserLookupSync = (
    userId: number,
    info: IFieldInfo
): UserLookup => {
    const lookUpListId = getLookupList(info);

    if (false === lookUpListId)
        throw new Error(
            `SharePointList getUserLookupSync no LookupListID in (field)info`
        );

    const controller = getById(
        lookUpListId
    ) as unknown as SharePointList<UserLookup>;
    const user: UserLookup = controller.getPartial(userId) as UserLookup;

    return user;
};
