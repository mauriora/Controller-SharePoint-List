import { FieldTypes, IFieldInfo } from '@pnp/sp/presets/all';
export { IFieldInfo } from '@pnp/sp/presets/all';

export interface IFieldInfoWithAllowMultipleValues extends IFieldInfo {
    AllowMultipleValues: boolean;
}

export const hasAllowMultipleValues = (fieldInfo: IFieldInfo | IFieldInfoWithAllowMultipleValues): fieldInfo is IFieldInfoWithAllowMultipleValues => 
    (fieldInfo as IFieldInfoWithAllowMultipleValues).AllowMultipleValues !== undefined && 
    'boolean' === typeof((fieldInfo as IFieldInfoWithAllowMultipleValues).AllowMultipleValues);

export const allowsMultipleValues = (fieldInfo: IFieldInfo | IFieldInfoWithAllowMultipleValues): boolean =>
    hasAllowMultipleValues( fieldInfo) && fieldInfo.AllowMultipleValues;

export interface IFieldInfoWithLookupList extends IFieldInfo {
    LookupList: string;
}

export const hasLookupList = (fieldInfo: IFieldInfo | IFieldInfoWithLookupList): fieldInfo is IFieldInfoWithLookupList =>{
    return [FieldTypes.Lookup, FieldTypes.User ].includes(fieldInfo.FieldTypeKind) &&
    (fieldInfo as IFieldInfoWithLookupList).LookupList !== undefined &&
    'string' === typeof((fieldInfo as IFieldInfoWithLookupList).LookupList);
}
export const getLookupList = (fieldInfo: IFieldInfo | IFieldInfoWithLookupList) =>
    hasLookupList(fieldInfo) &&
    fieldInfo.LookupList;

export interface IFieldInfoWithIsKeyword extends IFieldInfo {
    IsKeyword: boolean;
}

export const hasIsKeyword = (fieldInfo: IFieldInfo | IFieldInfoWithIsKeyword): fieldInfo is IFieldInfoWithIsKeyword => 
    (fieldInfo as IFieldInfoWithIsKeyword).IsKeyword !== undefined && 
    'boolean' === typeof((fieldInfo as IFieldInfoWithIsKeyword).IsKeyword);

export const isKeyword = (fieldInfo: IFieldInfo | IFieldInfoWithIsKeyword): boolean =>
    hasIsKeyword( fieldInfo) && fieldInfo.IsKeyword;
    
export interface IFieldInfoWithTermSetId extends IFieldInfo {
    TermSetId: string;
}

export const hasTermSetId = (fieldInfo: IFieldInfo | IFieldInfoWithTermSetId): fieldInfo is IFieldInfoWithTermSetId =>
    (fieldInfo as IFieldInfoWithTermSetId).TermSetId !== undefined &&
    'string' === typeof((fieldInfo as IFieldInfoWithTermSetId).TermSetId);

export const getTermSetId = (fieldInfo: IFieldInfo | IFieldInfoWithTermSetId) =>
    hasTermSetId(fieldInfo) &&
    fieldInfo.TermSetId;

export interface IFieldInfoWithChoices extends IFieldInfo {
    Choices: Array<string>;
}

export const hasChoices = (fieldInfo: IFieldInfo | IFieldInfoWithChoices): fieldInfo is IFieldInfoWithChoices => 
    fieldInfo.FieldTypeKind in [FieldTypes.Choice, FieldTypes.MultiChoice ] &&
    (fieldInfo as IFieldInfoWithChoices).Choices !== undefined &&
    Array.isArray((fieldInfo as IFieldInfoWithChoices).Choices);

export const getChoices = (fieldInfo: IFieldInfo | IFieldInfoWithChoices) =>
    hasChoices(fieldInfo) &&
    fieldInfo.Choices;

export interface IFieldInfoWithDisplayFormat extends IFieldInfo {
    DisplayFormat: number;
}

export const hasDisplayFormat = (fieldInfo: IFieldInfo | IFieldInfoWithDisplayFormat): fieldInfo is IFieldInfoWithDisplayFormat =>
    (fieldInfo as IFieldInfoWithDisplayFormat).DisplayFormat !== undefined &&
    'number' === typeof((fieldInfo as IFieldInfoWithDisplayFormat).DisplayFormat);

export const getDisplayFormat = (fieldInfo: IFieldInfo | IFieldInfoWithDisplayFormat) =>
    hasDisplayFormat(fieldInfo) &&
    fieldInfo.DisplayFormat;

export interface IFieldInfoWithFillInChoice extends IFieldInfo {
    FillInChoice: boolean;
}

export const hasFillInChoice = (fieldInfo: IFieldInfo | IFieldInfoWithFillInChoice): fieldInfo is IFieldInfoWithFillInChoice => 
    (fieldInfo as IFieldInfoWithFillInChoice).FillInChoice !== undefined && 
    'boolean' === typeof((fieldInfo as IFieldInfoWithFillInChoice).FillInChoice);

export const isFillInChoice = (fieldInfo: IFieldInfo | IFieldInfoWithFillInChoice): boolean =>
    hasFillInChoice( fieldInfo) && fieldInfo.FillInChoice;

export interface IFieldInfoWithRichtText extends IFieldInfo {
    RichtText: boolean;
}

export const hasRichtText = (fieldInfo: IFieldInfo | IFieldInfoWithRichtText): fieldInfo is IFieldInfoWithRichtText => 
    (fieldInfo as IFieldInfoWithRichtText).RichtText !== undefined && 
    'boolean' === typeof((fieldInfo as IFieldInfoWithRichtText).RichtText);

export const isRichtText = (fieldInfo: IFieldInfo | IFieldInfoWithRichtText): boolean =>
    hasRichtText( fieldInfo) && fieldInfo.RichtText;

export interface IFieldInfoWithMaximumValue extends IFieldInfo {
    MaximumValue: number;
}

export const hasMaximumValue = (fieldInfo: IFieldInfo | IFieldInfoWithMaximumValue): fieldInfo is IFieldInfoWithMaximumValue =>
    (fieldInfo as IFieldInfoWithMaximumValue).MaximumValue !== undefined &&
    'number' === typeof((fieldInfo as IFieldInfoWithMaximumValue).MaximumValue);

export const getMaximumValue = (fieldInfo: IFieldInfo | IFieldInfoWithMaximumValue) =>
    hasMaximumValue(fieldInfo) &&
    fieldInfo.MaximumValue;

export interface IFieldInfoWithMinimumValue extends IFieldInfo {
    MinimumValue: number;
}

export const hasMinimumValue = (fieldInfo: IFieldInfo | IFieldInfoWithMinimumValue): fieldInfo is IFieldInfoWithMinimumValue =>
    (fieldInfo as IFieldInfoWithMinimumValue).MinimumValue !== undefined &&
    'number' === typeof((fieldInfo as IFieldInfoWithMinimumValue).MinimumValue);

export const getMinimumValue = (fieldInfo: IFieldInfo | IFieldInfoWithMinimumValue) =>
    hasMinimumValue(fieldInfo) &&
    fieldInfo.MinimumValue;

export interface IFieldInfoWithTimeFormat extends IFieldInfo {
    TimeFormat: number;
}

export const hasTimeFormat = (fieldInfo: IFieldInfo | IFieldInfoWithTimeFormat): fieldInfo is IFieldInfoWithTimeFormat =>
    (fieldInfo as IFieldInfoWithTimeFormat).TimeFormat !== undefined &&
    'number' === typeof((fieldInfo as IFieldInfoWithTimeFormat).TimeFormat);

export const getTimeFormat = (fieldInfo: IFieldInfo | IFieldInfoWithTimeFormat) =>
    hasTimeFormat(fieldInfo) &&
    fieldInfo.TimeFormat;
    
