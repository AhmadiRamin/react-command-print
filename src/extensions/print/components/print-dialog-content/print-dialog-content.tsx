import * as React from 'react';
import * as ReactElementToString  from 'react-element-to-string';
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
import ListHelper from '../../util/list-helper';
import ReactHtmlParser from 'react-html-parser';
import {
    Dropdown
} from 'office-ui-fabric-react';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import ListService from '../../services/list-service';
import { isArray } from '@pnp/common';
import ITemplateItem from '../../models/template-item';
import { style } from 'typestyle';
import printStyles from '../print-dialog-template-content/print-template-content.module.scss';
import { sp, EmailProperties } from '@pnp/sp';

const _items: any[] = [];
export default class PrintDialogContent extends React.Component<IPrintDialogContentProps, IPrintDialogContentState> {
    private componentRef;
    private listService: ListService;
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
            printTemplate: null,
            selectedTemplateIndex: -1,
            itemContent: {},
            isSiteAdmin:false
        };
        this.listService = new ListService();
        this._onTemplateAdded = this._onTemplateAdded.bind(this);
        this._onTemplateUpdated = this._onTemplateUpdated.bind(this);
        this._onTemplateRemoved = this._onTemplateRemoved.bind(this);
        this.getItemContent = this.getItemContent.bind(this);
        this._makeEmailBody=this._makeEmailBody.bind(this);
        this._sendAsEmail = this._sendAsEmail.bind(this);
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
                title={`Print ${this.props.title}`}
                onDismiss={this.props.close}
                showCloseButton={true}
            >

                <div className="ms-grid-row">
                    <Spinner hidden={this.state.hideLoading} size={SpinnerSize.large} label={this.state.loadingMessage} ariaLive="assertive" />
                </div>
                <div className="ms-Grid" dir="ltr" hidden={!this.state.hideLoading}>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm8 ms-md8 ms-lg8">
                            <Dropdown
                                placeHolder="Select your template..."
                                options={this.state.templates.map(t => ({ key: t.Id, text: t.Title }))}
                                onChanged={this._onDropDownChanged}
                            />
                        </div>
                        <div className={styles.printIcons + " ms-Grid-col ms-sm4 ms-md4 ms-lg4"}>
                            <ReactToPrint
                                trigger={() => <IconButton iconProps={{ iconName: 'Print' }} title="Print" ariaLabel="Print" />}
                                content={() => this.componentRef}
                            />
                            <span hidden={true}>
                            <IconButton iconProps={{ iconName: 'Mail' }} title="Mail" ariaLabel="Mail" onClick={this._sendAsEmail} />
                            <IconButton iconProps={{ iconName: 'PDF' }} title="PDF" ariaLabel="PDF" />
                            <IconButton iconProps={{ iconName: 'ExcelLogo' }} title="Export to excel" ariaLabel="ExcelLogo" />
                            </span>
                            <span hidden={!this.state.isSiteAdmin}><IconButton iconProps={{ iconName: 'Settings' }} title="Settings" ariaLabel="Settings" onClick={this._setShowPanel(true)} /></span>
                        </div>
                    </div>
                    <div className={`${styles.loadingMargin} ms-grid-row`}>
                        <Spinner hidden={this.state.hideTemplateLoading} size={SpinnerSize.large} label={this.state.loadingMessage} ariaLive="assertive" />
                    </div>
                    <div hidden={!this.state.printTemplate} className={`${styles.templateContent} ms-grid-row`}>
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


    private _sendAsEmail() {
        
        if (this.state.printTemplate) {
            const Body = ReactElementToString(this._makeEmailBody());
            console.log(Body);
            const email: EmailProperties = {
                To: ["ramin.ahmadi@live.com"],
                Body,
                Subject: "Test"
            };
            sp.utility.sendEmail(email).then();
        }
        
    }

    private _makeEmailBody(): any {
        return <div className={printStyles.Print}>
            {this.state.printTemplate &&
                <div className={printStyles.Print}>
                    <div className={printStyles.printHeader}>
                        {ReactHtmlParser(this.state.printTemplate.header)}
                    </div>
                    <div className={printStyles.printContent}>
                        {
                            this.state.printTemplate.content
                        }
                    </div>
                    <div className={printStyles.printFooter}>
                        {ReactHtmlParser(this.state.printTemplate.footer)}
                    </div>

                </div>
            }
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
        const template = this.state.templates[index];
        const content = this.loadTemplate(template);
        this.setState({
            printTemplate: {
                header: template.Header,
                footer: template.Footer,
                content
            },
            selectedTemplateIndex: index,
            hideTemplateLoading: true
        });

    }

    private _onTemplateUpdated(index: number, template: ITemplateItem) {
        const newTemplatesList = [...this.state.templates];
        newTemplatesList[index] = { ...template, Columns: JSON.stringify(template.Columns) };
        this.setState({
            templates: newTemplatesList
        });
        if (this.state.selectedTemplateIndex === index)
            this.loadTemplate(index);

    }

    private loadTemplate(template: any): any[] {
        this.setState({
            hideTemplateLoading: false
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
                            <table className={styles.templateTable}>
                                {
                                    table.map(el => <tr>
                                        <td className={styles.nameColumn}>
                                            {el.Name}
                                        </td>
                                        <td className={styles.valueColumn}>
                                            {el.Value}
                                        </td>
                                    </tr>)
                                }
                            </table>
                        );
                    }
                    const { BackgroundColor, FontColor } = item;
                    const className = style({ backgroundColor: BackgroundColor, color: FontColor });
                    content.push(<div className={`${styles.templateSection} ${className}`}><span>{item.Title}</span></div>);
                    table = [];
                }
                if (item.Type === "Field") {
                    if (template.SkipBlankColumns) {
                        if (this.state.itemContent[item.InternalName].length > 0)
                            table.push({
                                Name: item.Title,
                                Value: this.state.itemContent[item.InternalName]
                            });
                    }
                    else {
                        table.push({
                            Name: item.Title,
                            Value: this.state.itemContent[item.InternalName]
                        });
                    }

                }
                if (i + 1 === columns.length) {
                    if (table.length > 0) {
                        content.push(
                            <table className={styles.templateTable}>
                                {
                                    table.map(el => <tr>
                                        <td className={styles.nameColumn}>
                                            {el.Name}
                                        </td>
                                        <td className={styles.valueColumn}>
                                            {el.Value}
                                        </td>
                                    </tr>)
                                }
                            </table>
                        );
                    }
                }
            }
        }

        return content;
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

    private async getItemContent() {
        const { listId, itemId } = this.props;
        const itemContent = await this.listService.GetItemById(listId, itemId);
        const isSiteAdmin = await this.listService.IsCurrentUserSiteAdmin();
        this.setState({
            itemContent,
            isSiteAdmin
        });
    }
}