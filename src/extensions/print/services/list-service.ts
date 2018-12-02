import {
    sp
} from "@pnp/sp";
import IListField from './list-field';
import * as strings from 'PrintCommandSetStrings';
import { Log } from '@microsoft/sp-core-library';
import { IListFieldResult } from "./list-field-result";
import * as moment from 'moment';
const LOG_SOURCE: string = 'PrintCommandSet';
export default class ListService {

    private _listFields: Array<IListFieldResult>;
    private _fieldTypesToIgnore: Array<string>;
    private _fieldsToIgnore: Array<string>;
    constructor() {
        this.buildExclusions();
    }

    public async GetItemById(listId: string, itemId: number):Promise<any>{
        return new Promise<any>((resolve: (values:any) => void, reject: (error: any) => void): void => {
            this.ensureListSchema(listId) //Go get the field information
            .then((listFields: Array<IListFieldResult>): void => {
                //Get an array of the internal field names for the select along with any necessary expand fields
                let fieldNames: Array<string> = new Array<string>();
                let expansions: Array<string> = new Array<string>();
                listFields.forEach((field: IListFieldResult) => {
                    switch (field.TypeAsString) {
                        case 'User':
                        case 'UserMulti':
                        case 'Lookup':
                        case 'LookupMulti':
                            fieldNames.push(field.InternalName + '/Id');
                            expansions.push(field.InternalName);
                            break;
                        default:
                            fieldNames.push(field.InternalName);
                    }
                });
                //Get the item values
                sp.web.lists.getById(listId).items.getById(itemId).select(...fieldNames).expand(...expansions).get<Array<any>>()
                    .then((result: any) => {
                        //Copy just the fields we care about and provide some adjustments for certain field types
                        let item: any = {};
                        listFields.forEach((field: IListFieldResult) => {
                            switch (field.TypeAsString) {
                                case 'User':
                                case 'Lookup':
                                    //These items need to be the underlying Id and their names have to have Id appended to them
                                    item[field.InternalName] = result[field.InternalName]['Id'];
                                    break;
                                case 'UserMulti':
                                case 'LookupMulti':
                                    //These items need to be an array of the underlying Ids and the array has to be called results
                                    // their names also have to have Id appended to them
                                    item[field.InternalName] = {
                                        results: new Array<Number>()
                                    };
                                    result[field.InternalName].forEach((prop: any) => {
                                        item[field.InternalName].results.push(prop['Id']);
                                    });
                                    break;
                                case "TaxonomyFieldTypeMulti":
                                    //These doesn't need to be included, since the hidden Note field will take care of these
                                    // in fact, including these will cause problems
                                    break;
                                case "MultiChoice":
                                    //These need to be in an array of the selected choices and the array has to be called results
                                    item[field.InternalName] = {
                                        results: result[field.InternalName]
                                    };
                                    break;
                                case "DateTime":
                                    item[field.InternalName] = moment(result[field.InternalName]).format("YYYY/MM/DD");
                                    break;
                                default:
                                    //Everything else is just a one for one match
                                    item[field.InternalName] = result[field.InternalName];
                            }
                        });
                        resolve(item);
                    })
                    .catch((error: any): void => {
                        Log.error(LOG_SOURCE, error);
                        this.safeLog(error);
                        reject(error);
                    });
            })
            .catch((error: any): void => {
                Log.error(LOG_SOURCE, error);
                this.safeLog(error);
                reject(error);
            });
        });
    }
    /**
     * GetTemplatesByListId
     */
    public async GetTemplatesByListId(listId: string): Promise<any[]> {
        return sp.web.lists.getByTitle('PrintSettings').items.filter(`ListId eq '${listId}'`).select('Id', 'Title', 'Header', 'Footer', 'Columns', 'ListId').get().then(items => {
            return items;
        }).catch(error => {
            Log.error(LOG_SOURCE, error);
            return error.message;
        });
    }

    /**
     * GetFieldsByListId
     */
    public async GetFieldsbyListId(listId: string): Promise<Array<IListField>> {
        return sp.web.lists.getById(listId).fields.select('Id', 'Title', 'InternalName','TypeAsString','IsDependentLookup').get().then((results: any) => {
            //Setup the list fields
            const _listFields = new Array<IListField>();
            // This includes any field of a type we don't want (such as computed)
            // This also includes several internal fields that won't make sense to clone (such as the creation date)
            // Finally, no dependent lookup columns (projected fields)
            for (let field of results) {
                const { InternalName, TypeAsString, Title, IsDependentLookup, Id } = field;
                if (this._fieldTypesToIgnore.indexOf(TypeAsString) == -1 && this._fieldsToIgnore.indexOf(InternalName) == -1 && !IsDependentLookup) {

                    _listFields.push({
                        InternalName,
                        Title: Title,
                        Id: Id,
                        Type: 'Field'
                    });

                }
            }
            return _listFields;
        }).catch(error => {
            Log.error(LOG_SOURCE, error);
            return error.message;
        });
    }

    public async AddTemplate(template: any): Promise<any> {
        console.log(template);
        return sp.web.lists.getByTitle('PrintSettings').items.add(template).then(({ data }) => data).catch(error => {
            Log.error(LOG_SOURCE, error);
            return error.message;
        });
    }

    public async UpdateTemplate(id: number, template: any): Promise<boolean> {
        return sp.web.lists.getByTitle('PrintSettings').items.getById(id).update(template).then(e => true).catch(error => {
            Log.error(LOG_SOURCE, error);
            return error.message;
        });
    }

    public async removeTempate(id: number): Promise<boolean> {
        return sp.web.lists.getByTitle('PrintSettings').items.getById(id).delete().then(e => true).catch(error => {
            Log.error(LOG_SOURCE, error);
            return error.message;
        });
    }

    /** Retrieves all the fields for the list */
    private ensureListSchema(listId: string): Promise<Array<IListFieldResult>> {
        return new Promise<Array<IListFieldResult>>((resolve: (listFields: Array<IListFieldResult>) => void, reject: (error: any) => void): void => {

            if (this._listFields) {
                //Looks like we already got it, so just return that
                resolve(this._listFields);

            } else {
                //Go get all the fields for the list
                sp.web.lists.getById(listId).fields.select('InternalName', 'TypeAsString', 'IsDependentLookup').get<IListFieldResult[]>()
                    .then((results: IListFieldResult[]) => {

                        //Setup the list fields
                        this._listFields = new Array<IListFieldResult>();

                        //Filter out all the extra fields we don't want to clone
                        // This includes any field of a type we don't want (such as computed)
                        // This also includes several internal fields that won't make sense to clone (such as the creation date)
                        // Finally, no dependent lookup columns (projected fields)
                        for (let field of results) {
                            if (this._fieldTypesToIgnore.indexOf(field.TypeAsString) == -1 && this._fieldsToIgnore.indexOf(field.InternalName) == -1 && !field.IsDependentLookup) {

                                this._listFields.push({
                                    InternalName: field.InternalName,
                                    TypeAsString: field.TypeAsString
                                });

                            }
                        }
                        resolve(this._listFields);
                    })
                    .catch((error: any): void => {
                        reject(error);
                    });
            }
        });
    }
    /** Builds the fieldTypes and fields to ignore arrays */
    private buildExclusions(): void {
        this._fieldTypesToIgnore = new Array<string>(
            strings.typeCounter,
            strings.typeContentType,
            strings.typeAttachments,
            strings.typeModStat,
            strings.typeComputed
        );

        this._fieldsToIgnore = new Array<string>(
            strings.field_ContentTypeId,
            strings.field_HasCopyDestinations,
            strings.field_CopySource,
            strings.fieldowshiddenversion,
            strings.fieldWorkflowVersion,
            strings.field_UIVersion,
            strings.field_UIVersionString,
            strings.field_ModerationComments,
            strings.fieldInstanceID,
            strings.field_ComplianceAssetId,
            strings.fieldGUID,
            strings.fieldWorkflowInstanceID,
            strings.fieldFileRef,
            strings.fieldFileDirRef,
            strings.fieldLast_x0020_Modified,
            strings.fieldCreated_x0020_Date,
            strings.fieldFSObjType,
            strings.fieldSortBehavior,
            strings.fieldFileLeafRef,
            strings.fieldUniqueId,
            strings.fieldSyncClientId,
            strings.fieldProgId,
            strings.fieldScopeId,
            strings.fieldFile_x0020_Type,
            strings.fieldMetaInfo,
            strings.field_Level,
            strings.field_IsCurrentVersion,
            strings.fieldItemChildCount,
            strings.fieldRestricted,
            strings.fieldOriginatorId,
            strings.fieldNoExecute,
            strings.fieldContentVersion,
            strings.field_ComplianceFlags,
            strings.field_ComplianceTag,
            strings.field_ComplianceTagWrittenTime,
            strings.field_ComplianceTagUserId,
            strings.fieldAccessPolicy,
            strings.field_VirusStatus,
            strings.field_VirusVendorID,
            strings.field_VirusInfo,
            strings.fieldAppAuthor,
            strings.fieldAppEditor,
            strings.fieldSMTotalSize,
            strings.fieldSMLastModifiedDate,
            strings.fieldSMTotalFileStreamSize,
            strings.fieldSMTotalFileCount,
            strings.fieldFolderChildCount,
            strings.fieldOrder
        );
    }

    /** Logs messages to the console if the console is available */
    private safeLog(message: any): void {
        if (window.console && window.console.log) {
            window.console.log(message);
        }
    }

}