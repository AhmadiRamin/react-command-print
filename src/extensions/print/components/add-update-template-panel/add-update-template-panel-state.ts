import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
export default interface AddUpdateTemplatePanelState{
    helperItems:any[];
    fields:any[];
    selectionDetails?: string;
    columns: IColumn[];
    itemColumns:IColumn[];
    isColumnReorderEnabled: boolean;
    frozenColumnCountFromStart: string;
    frozenColumnCountFromEnd: string;    
    templateColumns:any[];
    listId:string;
    section: string;
    showColorPicker: boolean;
    color:string;
}