import * as React from 'react';
import { initializeIcons } from '@uifabric/icons';
import ReactToPrint from "react-to-print";

import styles from './print-dialog.module.scss';
import {
    DialogContent, IDropdownOption
} from 'office-ui-fabric-react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import IPrintDialogContentProps from './print-dialog-content-props';
import IPrintDialogContentState from './print-dialog-content-state';
import PrintTemplateContent from '../print-dialog-template-content/print-template-content';
import SettingsPanel from '../settings-panel/settings-panel';
import { DetailsList, DetailsListLayoutMode, IColumn, CheckboxVisibility } from 'office-ui-fabric-react/lib/DetailsList';
import ListHelper from '../../util/list-helper';
import {
    Dropdown
} from 'office-ui-fabric-react';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import ListService from '../../services/list-service';
import { isArray } from '@pnp/common';
import ITemplateItem from '../../services/template-item';
const _items: any[] = [];
export default class PrintDialogContent extends React.Component<IPrintDialogContentProps, IPrintDialogContentState> {
    private componentRef;
    private listService: ListService;
    private _columns: IColumn[] = [
        {
            key: 'NameColumn',
            name: 'Name',
            fieldName: 'Name',
            minWidth: 100,
            maxWidth: 200
        },
        {
            key: 'ValueColumn',
            name: 'Value',
            fieldName: 'Value',
            minWidth: 100,
            maxWidth: 200
        }
    ];
    constructor(props) {
        super(props);

        if (_items.length === 0) {
            for (let i = 0; i < 10; i++) {
                _items.push({
                    key: i,
                    name: 'Item ' + i,
                    value: i
                });
            }
        }

        this.state = {
            hideLoading: false,
            loadingMessage: "Loading...",
            templates: [],
            items: _items,
            showPanel: false,
            hideTemplateLoading: true,
            printTemplate:null,
            selectedTemplateIndex:-1,
            itemContent:{}
        };
        this.listService = new ListService();
        this._onTemplateAdded = this._onTemplateAdded.bind(this);
        this._onTemplateUpdated = this._onTemplateUpdated.bind(this);
        this._onTemplateRemoved = this._onTemplateRemoved.bind(this);
        this.getItemContent=this.getItemContent.bind(this);
        // Initialize icons
        initializeIcons();
    }

    public componentDidMount() {

        // Validate and create Print Settings list
        this.initializeSettings();

        // Get select item values
        this.getItemContent();
    }

    public render(): JSX.Element {
        const templates = this.state.templates;
        return <div className={styles.PrintDialogContent}>
            <DialogContent
                title='Print List Item'
                onDismiss={this.props.close}
                showCloseButton={true}
            >

                <div className="ms-grid-row">
                    <Spinner hidden={this.state.hideLoading} size={SpinnerSize.large} label={this.state.loadingMessage} ariaLive="assertive" />
                </div>
                <div className="ms-Grid" dir="ltr" hidden={!this.state.hideLoading}>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg8">
                            <Dropdown
                                placeHolder="Select your template..."
                                options={this.state.templates.map(t => ({ key: t.Id, text: t.Title }))}
                                onChanged={this._onDropDownChanged}
                            />
                        </div>
                        <div className={styles.printIcons + " ms-Grid-col ms-sm6 ms-md4 ms-lg14"}>
                            <ReactToPrint
                                trigger={() => <IconButton iconProps={{ iconName: 'Print' }} title="Print" ariaLabel="Print" />}
                                content={() => this.componentRef}
                            />
                            <IconButton iconProps={{ iconName: 'Mail' }} title="Mail" ariaLabel="Mail" />
                            <IconButton iconProps={{ iconName: 'PDF' }} title="PDF" ariaLabel="PDF" />
                            <IconButton iconProps={{ iconName: 'ExcelLogo' }} title="Export to excel" ariaLabel="ExcelLogo" />
                            <IconButton iconProps={{ iconName: 'Settings' }} title="Settings" ariaLabel="Settings" onClick={this._setShowPanel(true)} />
                        </div>
                    </div>
                    <div className={`${styles.loadingMargin} ms-grid-row`}>
                        <Spinner hidden={this.state.hideTemplateLoading} size={SpinnerSize.large} label={this.state.loadingMessage} ariaLive="assertive" />
                    </div>
                    <div hidden={!this.state.hideTemplateLoading} className={`${styles.templateContent} ms-grid-row`}>
                        <PrintTemplateContent itemId={this.props.itemId} template={this.state.printTemplate} ref={el => (this.componentRef = el)} />
                    </div>
                </div>
                <SettingsPanel onTemplateAdded={this._onTemplateAdded}
                    onTemplateRemoved={this._onTemplateRemoved}
                    onTemplateUpdated={this._onTemplateUpdated}
                    templates={isArray(templates) ? templates : []} showPanel={this.state.showPanel} setShowPanel={this._setShowPanel} listId={this.props.listId} />
            </DialogContent>
        </div>;
    }

    private _onTemplateAdded(template: ITemplateItem) {
        this.setState(prevState => (
            {
                templates: prevState.templates.concat(template)
            }
        ));
    }


    private _onDropDownChanged = (option: IDropdownOption, index?: number) => {        
        
        this.loadTemplate(index);
    }

    private _onTemplateUpdated(index: number, template: ITemplateItem) {
        const newTemplatesList = [...this.state.templates];
        newTemplatesList[index] = {...template,Columns:JSON.stringify(template.Columns)};
        this.setState({
            templates: newTemplatesList
        });
        if(this.state.selectedTemplateIndex === index)
            this.loadTemplate(index);
        
    }

    private loadTemplate(index:number){
        const template = this.state.templates[index];

        this.setState({
            hideTemplateLoading:false
        });

        const columns: any[] = JSON.parse(template.Columns);
        let table: any[] = [];
        const content: any[] = [];
        if (columns.length > 0) {
            for (var i = 0; i < columns.length; i++) {
                const item = columns[i];
                if (item.Type === "Section") {
                    if (table.length > 0) {
                        content.push(
                            <DetailsList
                                items={table}
                                columns={this._columns}
                                isHeaderVisible={false}
                                className={styles.templateTable}
                                setKey="set"
                                layoutMode={DetailsListLayoutMode.fixedColumns}
                                checkboxVisibility={CheckboxVisibility.hidden}
                                selectionPreservedOnEmptyClick={true}
                            />
                        );
                    }
                    content.push(<div className={styles.templateSection}><span>{item.Title}</span></div>);
                    table = [];
                }
                if (item.Type === "Field") {
                    table.push({
                        Name: item.Title,
                        Value: this.state.itemContent[item.InternalName]
                    });
                }
                if(i+1 === columns.length){
                    if (table.length > 0) {
                        content.push(
                            <DetailsList
                                items={table}
                                columns={this._columns}
                                isHeaderVisible={false}
                                setKey="set"
                                className={styles.templateTable}
                                layoutMode={DetailsListLayoutMode.fixedColumns}
                                checkboxVisibility={CheckboxVisibility.hidden}
                                selectionPreservedOnEmptyClick={true}
                            />
                        );
                    }
                }
            }
        }
        this.setState({
            printTemplate:{
                header: template.Header,
                footer: template.Footer,
                content
            },
            selectedTemplateIndex:index,
            hideTemplateLoading:true
        });  
    }

    private async _onTemplateRemoved(id: number, template: ITemplateItem) {
        const removedItem = await this.listService.removeTempate(id);
        if (removedItem)
            this.setState(prevState => ({
                templates: prevState.templates.filter(el => el != template)
            }));
    }

    public _setShowPanel = (showPanel: boolean): (() => void) => {
        return (): void => {
            this.setState({ showPanel });
        };
    }

    private async initializeSettings() {
        const listHelper = new ListHelper(this.props.webUrl);

        listHelper.ValidatePrintSettingsList().then(_ => {
            this.setState({
                hideLoading: true
            });
        }).catch(e => {
            // Print Settings list already exists
            this.getTemplates();
        });
    }

    private async getTemplates() {
        const templates = await this.listService.GetTemplatesByListId(this.props.listId);
        this.setState({
            templates,
            hideLoading: true
        });
    }

    private async getItemContent(){
        const {listId,itemId} = this.props;
        const itemContent = await this.listService.GetItemById(listId,itemId);
        
        this.setState({
            itemContent
        });
    }
}