import {
    sp
} from "@pnp/sp";
import IListField from './list-field';
import * as strings from 'PrintCommandSetStrings';
import { Log } from '@microsoft/sp-core-library';

const LOG_SOURCE: string = 'PrintCommandSet';
export default class ListService {

    private _listFields: Array<IListField>;
    private _fieldTypesToIgnore: Array<string>;
    private _fieldsToIgnore: Array<string>;
    constructor() {
        this.buildExclusions();
    }

    /**
     * GetItemById
     */
    public GetItemById(id: number) {

    }

    /**
     * GetTemplatesByListId
     */
    public async GetTemplatesByListId(listId: string): Promise<any[]> {
        return sp.web.lists.getByTitle('PrintSettings').items.filter(`ListId eq '${listId}'`).select('Title').get().then(items => {
            return items;
        }).catch(error => {
            Log.error(LOG_SOURCE, error);
            return error.message;
        });
    }

    /**
     * GetFieldsByListId
     */
    public async GetFieldsByListId(listId: string): Promise<Array<IListField>> {
        return sp.web.lists.getById(listId).fields.select('Title', 'InternalName', 'TypeAsString', 'IsDependentLookup').get().then((results: IListField[]) => {
            //Setup the list fields
            this._listFields = new Array<IListField>();
            // This includes any field of a type we don't want (such as computed)
            // This also includes several internal fields that won't make sense to clone (such as the creation date)
            // Finally, no dependent lookup columns (projected fields)
            for (let field of results) {
                const {InternalName,TypeAsString,Title,IsDependentLookup} = field;
                if (this._fieldTypesToIgnore.indexOf(TypeAsString) == -1 && this._fieldsToIgnore.indexOf(InternalName) == -1 && !IsDependentLookup) {

                    this._listFields.push({
                        InternalName,
                        TypeAsString,
                        Title:Title
                    });

                }
            }
            return this._listFields;
        }).catch(error => {
            Log.error(LOG_SOURCE, error);
            return error.message;
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
            strings.fieldFolderChildCount
        );
    }

}