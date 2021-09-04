import { defaultMetadataStorage } from 'class-transformer/esm5/storage';
//import { defaultMetadataStorage } from 'class-transformer/types/storage';
import { FieldTypes, IFieldInfo } from "@pnp/sp/fields";
import { ListItemBase, ListItemBaseConstructor } from "../../models/ListItemBase";
import { Model } from "../controller";
import { SharePointList } from './SharePointList';

export class SharePointModel<DataType extends ListItemBase = ListItemBase> implements Model<ListItemBase, DataType> {
    public model: DataType;
    public newRecord: DataType;

    public records: Array<DataType>;
    public loadAllRecords = async () => this.controller.loadAllRecords(this.filter);

    public jsFactoryFactory: () => ListItemBaseConstructor<DataType>;

    public constructor(
        public jsFactory: ListItemBaseConstructor<DataType>,
        public controller: SharePointList<DataType>,
        public filter: string
    ) {
        this.jsFactoryFactory = () => jsFactory;
        this.newRecord = controller.newRecord;
        this.records = controller.records;
    }

    public submit = async (newRecord?: DataType): Promise<string | void> => { 
        const result = await this.controller.submit(newRecord);
        this.newRecord = this.controller.newRecord;
        return result;
    };

    /** fields for $select part of the query */
    public get selectFields(): Array<string> {
        if( undefined === this._selectFields) {
            this.initSelectAndExpands();
        }
        return this._selectFields;
    };

    /** fields for $expand part of the query */
    public get expandFields(): Array<string> {
        if( undefined === this._expandFields) {
            this.initSelectAndExpands();
        }
        return this._expandFields;
    };

    /** Property name mapped to SharePoint Fieldinfo  */
    public get propertyFields(): Map<string, IFieldInfo> {
        if( undefined === this._propertyFields) {
            this.initSelectAndExpands();
        }
        return this._propertyFields;
    };

    /** Property name mapped to SharePoint Fieldinfo  */
    public get selectedFields(): Map<string, IFieldInfo> {
        if( undefined === this._propertyFields) {
            this.initSelectAndExpands();
        }
        return this._selectedFields;
    };

    private _selectedFields: Map<string, IFieldInfo>;
    private _propertyFields: Map<string, IFieldInfo>;

    /** fields for $select part of the query */
    private _selectFields: Array<string>;

    /** fields for $expand part of the query */
    private _expandFields: Array<string>;

    private initSelectAndExpands = () => {
        const { expands, selects, selectedFields, propertyFields } = this.getSelectAndExpand(
            this.jsFactoryFactory(),
            this.controller.allFields
        );
        this._expandFields = expands;
        this._selectFields = selects;
        this._selectedFields = selectedFields;
        this._propertyFields = propertyFields;
    }

    private static IGNORED_PROPERTIES = [];
    private static IGNORED_EXPANDS = ['Author', 'Editor', 'Attachments', 'AverageRating', 'RatingCount', 'Ratings', 'LikesCount', 'TaxKeyword' ];
    private static OPTIONAL_FIELDS = ['TaxKeyword', 'TaxCatchAll', 'AverageRating', 'RatingCount', 'RatedBy', 'Ratings', 'LikesCount', 'LikedBy'];
    
    /**
     * Adds the fieldName to expandedFields and adds each expansion as fieldName/expansion to selectFields
     * @param fieldName e.g. author
     * @param expansions e.g. ['ID', 'Title']
     */
    private static addExpandField = (selectFields: Array<string>, expandFields: Array<string>, fieldName: string, expansions: string[]) => {
        selectFields.push(...
            expansions.map(
                expansion => (fieldName + '/' + expansion)
            )
        );
        expandFields.push(fieldName);
    }

    private getSelectAndExpand = (jsFactory: ListItemBaseConstructor<DataType>, fields?: Map<string, IFieldInfo>): { selects: Array<string>, expands: Array<string>, selectedFields?: Map<string, IFieldInfo>, propertyFields?: Map<string, IFieldInfo> } => {
        const blankJs = new jsFactory();
        const selects = new Array<string>();
        const expands = new Array<string>();
        const selectedFields = undefined === fields ? undefined : new Map<string, IFieldInfo>();
        const propertyFields = undefined === fields ? undefined : new Map<string, IFieldInfo>();

        for (const propertyName in blankJs) {
            if (typeof (blankJs[propertyName]) === 'function') {
                // console.debug(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} NOT add ${typeof (blankJs[propertyName])} to select`);
            } else if (propertyName in SharePointModel.IGNORED_PROPERTIES) {
                // console.debug(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} NOT add to select`);
            } else {
                const excludeData = defaultMetadataStorage.findExcludeMetadata(jsFactory, propertyName);

                if (undefined !== excludeData) {
                    // console.debug(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}) exclude ${propertyName}`, { blankJs, excludeData });
                } else {
                    const typeData = defaultMetadataStorage.findTypeMetadata(jsFactory, propertyName);
                    const exposeData = defaultMetadataStorage.findExposeMetadata(jsFactory, propertyName);
                    const fieldName = exposeData?.options?.name ?? propertyName;

                    if (undefined !== fields) {
                        const fieldInfo = fields.get(fieldName);
                        const fieldType = fieldInfo?.FieldTypeKind;

                        if (undefined === fieldType) {
                            if (SharePointModel.OPTIONAL_FIELDS.findIndex(prospect => prospect === fieldName) >= 0) {
                                console.log(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName} ignore`);
                            } else {
                                throw new Error(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName} not found in fields`);
                            }
                        } else if (0 <= [FieldTypes.Lookup, FieldTypes.User].indexOf(fieldType)) { // } else if ( fieldType in [FieldTypes.Lookup, FieldTypes.User]) {                        
                            if (undefined !== typeData) {
                                const lookUpFields = this.getSelectAndExpand(typeData.typeFunction() as ListItemBaseConstructor<DataType>);

                                SharePointModel.addExpandField(selects, expands, fieldName, lookUpFields.selects);
                                selectedFields.set(fieldName, fieldInfo);
                                propertyFields.set(propertyName, fieldInfo);
                            } else {
                                throw new Error(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName}[${fieldInfo.TypeAsString}] no type info`);
                                // selects.push(fieldName);
                            }
                        } else if (fieldType === FieldTypes.Attachments && (true !== this.controller.listInfo.EnableAttachments)) {
                            console.warn(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName} ignore, attachments are disabled`, { blankJs, exposeData, typeData, excludeData, selects, expands, listInfo: this?.controller?.listInfo, fieldInfo });
                        } else {
                            selects.push(fieldName);
                            selectedFields.set(fieldName, fieldInfo);
                            propertyFields.set(propertyName, fieldInfo);
                        }
                    } else { // (undefined === fields)
                        if ( SharePointModel.IGNORED_EXPANDS.findIndex(prospect => prospect === fieldName) >= 0  ) {
                            console.warn(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName} Ignore Expansion`, { blankJs, exposeData, typeData, excludeData, selects, expands });
                        } else if (undefined !== typeData) {
                            // const lookUpFields = this.getSelectAndExpand(typeData.typeFunction());
                            // SharePointModel.addExpandField( selects, expands, fieldName, lookUpFields.selects );
                            console.warn(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName} Ignore Sub Expansion`, { blankJs, exposeData, typeData, excludeData, selects, expands });
                        } else if (true || ['ID', 'Title'].find(prospect => prospect === fieldName)) {
                            // console.warn(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName}[] add to select`, { exposeData, typeData, excludeData, });
                            selects.push(fieldName);
                        } else {
                            console.warn(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName}[] NOT add to select`, { exposeData, typeData, excludeData, });
                        }
                    }
                }
            }
        }
        return { selects, expands, selectedFields, propertyFields };
    }

}