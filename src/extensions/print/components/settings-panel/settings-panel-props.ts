export default interface ISettingsPanelProps{
    showPanel:boolean;
    listId:string;
    setShowPanel: (showPanel: boolean)=> () => void ;
}