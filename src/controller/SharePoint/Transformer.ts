import { FieldTypes, IFieldInfo } from "@pnp/sp/fields";
import { EmptyGuid } from '@pnp/spfx-controls-react';
import { classToPlain } from "class-transformer";
import { ListItem } from "../../models/ListItem";
import { ListItemBase } from "../../models/ListItemBase";
import { MetaTerm, MetaTermSP } from "../../models/MetaTerm";
import { addTerm } from "../Taxonomy";
import { allowsMultipleValues, getTermSetId, isKeyword } from "./FieldInfo";

export const resultArrayToArray = (plain: Record<string, any>, selectedFields: Map<string, IFieldInfo>) => {
    for (const [fieldName, info] of selectedFields.entries()) {
        if ( allowsMultipleValues(info) && plain[fieldName]?.['results']) {
            plain[fieldName] = plain[fieldName]['results'];
        }
    }
}

export const setNullArrays = <ItemType extends ListItemBase>(item: ItemType, propertyFields: Map<keyof ItemType, IFieldInfo>) => {
    const source = item.source as Record<string, any>;

    for (const [propertyName, info] of propertyFields) {
        if (allowsMultipleValues(info) || info.FieldTypeKind === FieldTypes.MultiChoice) {
            if (!item[propertyName] || undefined !== source?.[info.InternalName]?.['__deferred']) {
                if (undefined !== source?.[info.InternalName]?.['__deferred']) {
                    console.warn(`setNullArrays .${propertyName} don't know what to do with deferred, set ${propertyName}=empty array !!`, { item, propertyName });
                }
                item[propertyName] = new Array() as any;
            }
        }
    }
}


export const fixSingleTaxonomyFields = <ItemType extends ListItem>(item: ItemType, propertyFields: Map<keyof ItemType, IFieldInfo>) => {
    for (const [propertyName, field] of propertyFields.entries()) {
        if (FieldTypes.Invalid === field.FieldTypeKind && 'TaxonomyFieldType' === field.TypeAsString) {
            const metaTerm = (item[propertyName] as MetaTerm);
            if (metaTerm) {
                if (item.taxCatchAll) {
                    const id = Number.parseInt(metaTerm.label);
                    const catchAll = item.taxCatchAll.find(prospect => id === prospect.id);
                    metaTerm.label = catchAll.term;
                } else {
                    console.error(`[${item.id}].fixSingleTaxonomyFields ${propertyName} ${metaTerm?.label} no catchAll`, { metaTerm, item });
                }
            }
        }
    }
}

const toSubmitArray = <SubmitType>(array: Array<SubmitType>): { results: Array<SubmitType> } =>
    ({ results: array });

const toTermStringSP = (terms: Array<MetaTermSP>): string =>
    terms
        .map(term => `-1;#${term.Label}|${term.TermGuid};`)
        .join('#');

/** Returns the fieldname for a multi metadata field, required to update the item. */
const getHiddenMetadataField = (normalName: string, allFields: Map<string, IFieldInfo>, fieldInfo?: IFieldInfo): string => {
    if (fieldInfo && isKeyword( fieldInfo) ) {
        return 'TaxKeywordTaxHTField';
    }
    const hiddenFieldTitle = normalName + '_0';

    for (const [fieldName, info] of allFields.entries()) {
        if (hiddenFieldTitle === info.Title) {
            return fieldName;
        }
    }
    throw new Error(`getHiddenMetadataField(${normalName}) can't find hidden field. Is this really a multi-value metadata field?`);
}


const toTaxonomyFieldTypeMulti = async (submitRecord: Record<string, any>, propertyName: string, terms: Array<MetaTermSP>, fieldInfo: IFieldInfo, allFields: Map<string, IFieldInfo>) => {
    console.warn(`[${submitRecord['ID']}].toTaxonomyFieldTypeMulti() NOT QUITE IMPLEMENTED YET ${propertyName}`, { submitRecordNow: { ...submitRecord }, propertyName, termsNow: terms ? { ...terms } : terms });

    if (terms && terms.length) {
        const hiddenFieldName = getHiddenMetadataField(propertyName, allFields, fieldInfo);

        for (const term of terms) {
            if (EmptyGuid === term.TermGuid) {
                await addTerm( getTermSetId( fieldInfo ), term);
            }
        }
        const termsString = toTermStringSP(terms);
        submitRecord[hiddenFieldName] = termsString;
        delete submitRecord[propertyName];
        console.warn(
            `[${submitRecord['ID']}].toTaxonomyFieldTypeMulti() NOT QUITE IMPLEMENTED YET ${propertyName}`,
            {
                hiddenFieldName, termsString,
                submitRecordNow: { ...submitRecord }
            }
        );
    } else {
        console.warn(`[${submitRecord['ID']}].toTaxonomyFieldTypeMulti() NOT QUITE IMPLEMENTED YET DELETE EMPTY ${propertyName}`, { submitRecordNow: { ...submitRecord }, propertyName, termsNow: terms ? { ...terms } : terms });
        delete submitRecord[propertyName];
    }
}

const toTaxonomyFieldType = (submitRecord: Record<string, any>, propertyName: string, term: MetaTermSP) => {
    if (term) {
        submitRecord[propertyName] = {
            "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
            ...term,
            WssId: '-1'
        };
        console.log(`[${submitRecord['ID']}].toTaxonomyFieldType() ${propertyName}`, { originalValue: term ? { ...term } : term, convertedValue: submitRecord[propertyName] ? { ...submitRecord[propertyName] } : submitRecord[propertyName] });
    }
}

export const toSubmit = async (jsRecord: ListItemBase, selectedFields: Map<string, IFieldInfo>, allFields: Map<string, IFieldInfo>) => {
    const submitRecord = classToPlain(jsRecord, { excludeExtraneousValues: true });
    console.debug(`[${jsRecord.id}].toSubmit()`, { jsRecord: { ...jsRecord }, submitRecord: { ...submitRecord } });

    for (const propertyName in submitRecord) {
        const propertyValue = submitRecord[propertyName];
        const fieldInfo = selectedFields.get(propertyName);

        if (undefined === fieldInfo) {
            if (['Attachments', 'TaxKeyword', 'TaxCatchAll'].findIndex(optional => optional === propertyName) >= 0) {
                console.warn(`[${jsRecord.id}].toSubmit() delete field ${propertyName}`, { jsRecord, submitRecord });
                delete submitRecord[propertyName];
            } else {
                throw new Error(`[${jsRecord.id}].toSubmit() '${propertyName}' not in FieldInfo, maybe add to hardcoded optional fields above`);
            }
        } else {
            if ('ID' !== fieldInfo.InternalName && fieldInfo.ReadOnlyField) {
                console.warn(`[${jsRecord.id}].toSubmit() ignore readOnly field ${propertyName}`, { jsRecord });
            } else {
                const multiValue = allowsMultipleValues(fieldInfo);

                switch (fieldInfo.FieldTypeKind) {
                    case FieldTypes.Invalid:
                        switch (fieldInfo.TypeAsString) {
                            case 'TaxonomyFieldType': toTaxonomyFieldType(submitRecord, propertyName, propertyValue);
                                break;
                            case 'TaxonomyFieldTypeMulti': await toTaxonomyFieldTypeMulti(submitRecord, propertyName, propertyValue, fieldInfo, allFields);
                                break;
                            default:
                                throw new Error(`[${jsRecord.id}].toSubmit() ignore NOT IMPLEMENTED YET ${propertyName}[${fieldInfo.TypeAsString}, ${fieldInfo.FieldTypeKind}]`);
                        }
                        break;
                    case FieldTypes.Attachments:
                        console.warn(`[${jsRecord.id}].toSubmit() ignore NOT IMPLEMENTED YET ATTACHMENTS ${propertyName}[${fieldInfo.TypeAsString}, ${fieldInfo.FieldTypeKind}]`);
                        break;
                    case FieldTypes.DateTime:
                        submitRecord[propertyName] = 'string' === typeof (propertyValue) ? propertyValue : (propertyValue ? propertyValue.toISOString() : propertyValue);
                        break;
                    case FieldTypes.MultiChoice:
                        {
                            submitRecord[propertyName] = toSubmitArray(propertyValue);
                        }
                        break;
                    case FieldTypes.Lookup:
                    case FieldTypes.User:
                        {
                            const newFieldName = propertyName + 'Id';
                            let newValue;

                            if (multiValue) {
                                if ((propertyValue as Array<{ ID: number }>).some(item => (undefined === item['ID']))) {
                                    throw new Error(`[${jsRecord.id}].toSubmit() don't know how to convert multi value ${propertyName}, it doesn't contain ID`);
                                } else {
                                    newValue = toSubmitArray(
                                        (propertyValue as Array<{ ID: number }>).map(itemRef => itemRef.ID)
                                    );
                                }
                            } else if (undefined === propertyValue) {
                                newValue = propertyValue;
                            } else if ('ID' in propertyValue) {
                                newValue = propertyValue['ID'];
                            } else {
                                throw new Error(`[${jsRecord.id}].toSubmit() don't know how to convert single value ${propertyName}, it doesn't contain ID`);
                            }
                            delete submitRecord[propertyName];
                            submitRecord[newFieldName] = newValue;
                        }
                        break;
                }
            }
        }
    }

    console.debug(`[${jsRecord.id}].toSubmit() done`, { jsRecord, submitRecord });

    return submitRecord;
}
