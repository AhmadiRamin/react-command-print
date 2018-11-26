
export default interface AddUpdateTemplatePanelProps{
    showTemplatePanel:boolean;
    listId:string;
    itemId?:number;
    setShowTemplatePanel: (showPanel: boolean)=> () => void ;
}