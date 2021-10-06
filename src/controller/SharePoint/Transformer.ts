import { FieldTypes, IFieldInfo } from "@pnp/sp/fields";
import { EmptyGuid } from '@pnp/spfx-controls-react';
import { classToPlain } from "class-transformer";
import { ListItem } from "../../models/ListItem";
import { ListItemBase } from "../../models/ListItemBase";
import { MetaTerm, MetaTermSP } from "../../models/MetaTerm";
import { addTerm } from "../Taxonomy";
import { allowsMultipleValues, getTermSetId, isKeyword } from "./FieldInfo";

enum NewFieldTypes {
    Image = 34
}
export type AllFieldTypes = FieldTypes | NewFieldTypes;

export interface ResultsArray {
    [key: string]: { results?: [] } | [];
}

/**
 * Deletes all properties with a value null.
 * Used as preparation for plainToClass when retrieving SharePoint items.
 * This ensures that the target class default values are used, e.g. an empty array and not set values are actually undefined.
 * @param item, e.g. {a: 1, b: [], c: null}
 * @returns the item without properties that were null, e.g.: {a: 1, b:[]}
 */
export const removeNullValues = (item: Record<string, null | unknown>): Record<string, unknown> => {
    for(const property in item) {
        if(null === item[property]) {
            delete item[property];
        }
    }
    return item;
}

/**
 * Parses the JSON string returned for image fields.
 * @param item the plain item returned from pnp.list.items.get
 * @param selectedFields fieldnames mapped to fieldinfo to determine Iamge fields
 * @return the plain object with all Image fields string value replaced by the parsed object
 */
export const parseImages = (plain: Record<string, string | unknown>, selectedFields: Map<string, IFieldInfo>): Record<string, unknown> => {
    for( const [fieldName, info] of selectedFields.entries()) {
        if( NewFieldTypes.Image === (info.FieldTypeKind as unknown as NewFieldTypes)) {
            const fieldValue = plain[fieldName];
            if(fieldValue) {
                if('string' === typeof fieldValue) {
                    console.debug(`Transformer:parseImages() found image field ${fieldName}`, {fieldValue, plainNow: {...plain}, plain});
                    const imageObject = JSON.parse(fieldValue);
                    plain[fieldName] = imageObject;
                    console.debug(`Transformer:parseImages() parsed image field ${fieldName}`, {fieldValue, imageObject, plainNow: {...plain}, plain});
                } else {
                    console.error(`Transformer:parseImages() found image field ${fieldName} of type '${typeof fieldValue}' should be 'string'`, {fieldValue, plainNow: {...plain}, plain});
                }
            } else {
                console.debug(`Transformer:parseImages() found empty image field ${fieldName}`, {plainNow: {...plain}, plain});
            }
        }
    }
    return plain;
}

export const resultArrayToArray = (plain: ResultsArray, selectedFields: Map<string, IFieldInfo>): void => {
    for (const [fieldName, info] of selectedFields.entries()) {
        const fieldValue = plain[fieldName];
        if (allowsMultipleValues(info) && fieldValue && 'results' in fieldValue) {
            console.debug(`resultArrayToArray[${fieldName}]`, { fieldValue });
            plain[fieldName] = fieldValue.results;
        }
    }
}

export interface DeferredContainer {
    [key: string]: { '__deferred': unknown } | [];
}

export const setNullArrays = <ItemType extends ListItemBase>(item: ItemType, propertyFields: Map<keyof ItemType, IFieldInfo>): void => {
    const source = item.source as DeferredContainer | undefined;

    for (const [propertyName, info] of propertyFields) {
        if (allowsMultipleValues(info) || info.FieldTypeKind === FieldTypes.MultiChoice) {
            const sourceValue = source?.[info.InternalName];
            if (!item[propertyName] || (sourceValue && '__deferred' in sourceValue)) {
                if (sourceValue && '__deferred' in sourceValue) {
                    console.warn(`setNullArrays .${propertyName} don't know what to do with deferred, set ${propertyName}=empty array !!`, { itemNow: {...item}, item, propertyName });
                } else {
                    console.error(`setNullArrays .${propertyName} arrays should be initialised, setting ${propertyName}=[] !!`, { itemNow: {...item}, item, propertyName });
                }
                (item[propertyName] as unknown) = [];
            }
        }
    }
}


export const fixSingleTaxonomyFields = <ItemType extends ListItem>(item: ItemType, propertyFields: Map<keyof ItemType, IFieldInfo>): void => {
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
    if (fieldInfo && isKeyword(fieldInfo)) {
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


const toTaxonomyFieldTypeMulti = async (submitRecord: Record<string, string>, propertyName: string, terms: Array<MetaTermSP>, fieldInfo: IFieldInfo, allFields: Map<string, IFieldInfo>) => {
    console.warn(`[${submitRecord['ID']}].toTaxonomyFieldTypeMulti() NOT QUITE IMPLEMENTED YET ${propertyName}`, { submitRecordNow: { ...submitRecord }, propertyName, termsNow: terms ? { ...terms } : terms });

    if (terms && terms.length) {
        const hiddenFieldName = getHiddenMetadataField(propertyName, allFields, fieldInfo);

        for (const term of terms) {
            if (EmptyGuid === term.TermGuid) {
                const termSetId = getTermSetId(fieldInfo);

                if (false === termSetId) throw new Error(`[${submitRecord['ID']}].toTaxonomyFieldTypeMulti() can't get TermSetId from FieldInfo ${fieldInfo}`);

                await addTerm(termSetId, term);
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

interface TaxonomySubmitField extends MetaTermSP {
        "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
        WssId: -1 // Was '-1'
}

interface TaxonomySubmitRecord {
    [key: string]: TaxonomySubmitField;
}

const toTaxonomyFieldType = (submitRecord: TaxonomySubmitRecord, propertyName: string, term: MetaTermSP) => {
    if (term) {
        submitRecord[propertyName] = {
            "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
            ...term,
            WssId: -1
        };
        console.log(`[${submitRecord['ID']}].toTaxonomyFieldType() ${propertyName}`, { originalValue: term ? { ...term } : term, convertedValue: submitRecord[propertyName] ? { ...submitRecord[propertyName] } : submitRecord[propertyName] });
    }
}

export const toSubmit = async (jsRecord: ListItemBase, selectedFields: Map<string, IFieldInfo>, allFields: Map<string, IFieldInfo>): Promise<Record<string, unknown>> => {
    const submitRecord = classToPlain(jsRecord, { excludeExtraneousValues: true });

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
