export default interface IPrintDialogContentState {
    loading: boolean;
    loadingMessage: string;
    printTemplates: Array<string>;
    items:any[];
    showPanel:boolean;
}