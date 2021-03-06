// causes webpack to fail -> import { defaultMetadataStorage } from 'class-transformer/types/storage';
// declared in ./MetadataStorage.d.ts
import { defaultMetadataStorage } from 'class-transformer/esm5/storage';

import { FieldTypes, IFieldInfo } from "@pnp/sp/fields";
import { ListItemBase, ListItemBaseConstructor } from "../../models/ListItemBase";
import { Model } from "../controller";
import { SharePointList } from './SharePointList';
import { WritablePart } from '../../models/WriteableParts';

export class SharePointModel<DataType extends ListItemBase = ListItemBase> implements Model<ListItemBase, DataType> {
    public model: DataType;
    public newRecord: DataType;

    public records: Array<DataType>;
    public loadAllRecords = async (): Promise<void> => this.controller.loadAllRecords(this.filter);

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

    public submit = async (newRecord?: DataType): Promise<void> => this.controller.submit(newRecord);

    /** fields for $select part of the query */
    public get selectFields(): Array<string> {
        if (undefined === this._selectFields) {
            this.initSelectAndExpands();
        }
        return this._selectFields;
    }

    /** fields for $expand part of the query */
    public get expandFields(): Array<string> {
        if (undefined === this._expandFields) {
            this.initSelectAndExpands();
        }
        return this._expandFields;
    }

    /** Property name mapped to SharePoint Fieldinfo  */
    public get propertyFields(): Map<keyof WritablePart<DataType>, IFieldInfo> {
        if (undefined === this._propertyFields) {
            this.initSelectAndExpands();
        }
        return this._propertyFields;
    }

    /** Property name mapped to SharePoint Fieldinfo  */
    public get selectedFields(): Map<string, IFieldInfo> {
        if (undefined === this._propertyFields) {
            this.initSelectAndExpands();
        }
        return this._selectedFields;
    }

    private _selectedFields: Map<string, IFieldInfo>;
    private _propertyFields: Map<keyof WritablePart<DataType>, IFieldInfo>;

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

    private static IGNORED_SUB_EXPANDS = ['Author', 'Editor', 'Attachments', 'AverageRating', 'RatingCount', 'Ratings', 'LikesCount', 'TaxKeyword', 'TaxCatchAll', 'RatedBy', 'LikedBy'];
    private static OPTIONAL_FIELDS = ['Attachments', 'TaxKeyword', 'TaxCatchAll', 'AverageRating', 'RatingCount', 'RatedBy', 'Ratings', 'LikesCount', 'LikedBy'];

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

    private getSelectAndExpand = (jsFactory: ListItemBaseConstructor<DataType>, fields?: Map<string, IFieldInfo>): { selects: Array<string>, expands: Array<string>, selectedFields?: Map<string, IFieldInfo>, propertyFields?: Map<keyof WritablePart<DataType>, IFieldInfo> } => {
        const blankJs = new jsFactory();
        const selects = new Array<string>();
        const expands = new Array<string>();
        const selectedFields = undefined === fields ? undefined : new Map<string, IFieldInfo>();
        const propertyFields = undefined === fields ? undefined : new Map<keyof WritablePart<DataType>, IFieldInfo>();

        const exposedMetadatas = defaultMetadataStorage.getExposedMetadatas(jsFactory)

        // for (const propertyName in blankJs) {
        for (const exposeData of exposedMetadatas) {
            const propertyName = exposeData.propertyName as keyof WritablePart<DataType>;
            const excludeData = defaultMetadataStorage.findExcludeMetadata(jsFactory, propertyName);

            if (undefined === excludeData) {
                const typeData = defaultMetadataStorage.findTypeMetadata(jsFactory, propertyName);
                const fieldName = exposeData?.options?.name ?? propertyName;

                if (undefined !== fields) {
                    const fieldInfo = fields.get(fieldName);
                    const fieldType = fieldInfo?.FieldTypeKind;

                    if (undefined === fieldType) {
                        if (SharePointModel.OPTIONAL_FIELDS.includes(fieldName)) {
                            // console.debug(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName} ignore optional (not found in fields)`);
                        } else {
                            throw new Error(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName} not found in fields`);
                        }
                    } else if ([FieldTypes.Lookup, FieldTypes.User].includes(fieldType)) {
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
                        // console.debug(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName} ignore, attachments are disabled`, { blankJs, exposeData, typeData, excludeData, selects, expands, listInfo: this?.controller?.listInfo, fieldInfo });
                    } else {
                        selects.push(fieldName);
                        selectedFields.set(fieldName, fieldInfo);
                        propertyFields.set(propertyName, fieldInfo);
                    }
                } else { // (undefined === fields) => child list
                    if (SharePointModel.IGNORED_SUB_EXPANDS.includes(fieldName)) {
                        // console.warn(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName} Ignore predefined sub Expansion`, { blankJs, exposeData, typeData, excludeData, selects, expands });
                    } else if (undefined !== typeData) {
                        console.warn(`SharePointModel[${this?.controller?.listInfo?.Title ?? this?.controller?.listId ?? this?.controller?.listTitle}].getSelectAndExpand(${jsFactory.name}).${propertyName} => ${fieldName} Ignore unkown Sub Expansion. If you want to access this property, then you need to create a controller for it.`, { blankJs, exposeData, typeData, excludeData, selects, expands });
                    } else {
                        selects.push(fieldName);
                    }
                }
            }
        }
        return { selects, expands, selectedFields, propertyFields };
    }

}