import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import ITemplate from './template';
export default interface AddUpdateTemplatePanelState{
    items:any[];
    helperItems:any[];
    fields:any[];
    selectionDetails?: string;
    columns: IColumn[];
    itemColumns:IColumn[];
    isColumnReorderEnabled: boolean;
    frozenColumnCountFromStart: string;
    frozenColumnCountFromEnd: string;
    template: ITemplate;
    section: string;
}