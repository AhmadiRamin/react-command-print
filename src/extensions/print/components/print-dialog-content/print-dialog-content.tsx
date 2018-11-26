import * as React from 'react';
import { initializeIcons } from '@uifabric/icons';
import ReactToPrint from "react-to-print";

import styles from './print-dialog.module.scss';
import {
    DialogContent
} from 'office-ui-fabric-react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import IPrintDialogContentProps from './print-dialog-content-props';
import IPrintDialogContentState from './print-dialog-content-state';
import PrintTemplateContent from '../print-dialog-template-content/print-template-content';
import SettingsPanel from '../settings-panel/settings-panel';

import ListHelper from '../../util/list-helper';
import {
    Dropdown
} from 'office-ui-fabric-react';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
const _items: any[] = [];
export default class PrintDialogContent extends React.Component<IPrintDialogContentProps, IPrintDialogContentState> {
    private componentRef;
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
            loading: true,
            loadingMessage: "Loading...",
            printTemplates: [],
            items: _items,
            showPanel:false
        };

        // Initialize icons
        initializeIcons();

        // Validate and create Print Settings list
        this.initializeSettings();
    }

    public render(): JSX.Element {

        return <div className={styles.PrintDialogContent}>
            <DialogContent
                title='Print List Item'
                onDismiss={this.props.close}
                showCloseButton={true}
            >
                {
                    this.state.loading ?
                        <div className="ms-grid-row">
                            <Spinner size={SpinnerSize.large} label={this.state.loadingMessage} ariaLive="assertive" />
                        </div>
                        :
                        <div className="ms-Grid" dir="ltr">
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg8">
                                    <Dropdown
                                        placeHolder="Select your template..."
                                        options={[
                                            { key: 'A', text: 'Template 1' },
                                            { key: 'B', text: 'Template 2' },
                                            { key: 'C', text: 'Template 3' }
                                        ]}
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
                            <div className={`${styles.detailsListMargin} ms-grid-row`}>
                                <PrintTemplateContent items={this.state.items} ref={el => (this.componentRef = el)} />
                            </div>
                        </div>
                        
                }
                <SettingsPanel showPanel={this.state.showPanel} setShowPanel={this._setShowPanel} listId={this.props.listId}/>
            </DialogContent>
        </div>;
    }

    public _setShowPanel = (showPanel: boolean): (() => void) => {
        return (): void => {
            this.setState({ showPanel });
        };
    }

    private initializeSettings() {
        const listHelper = new ListHelper(this.props.webUrl);
        listHelper.ValidatePrintSettingsList().then(_ => {
            this.setState({
                loadingMessage: 'Initializing settings',
                loading: true
            });
        }).catch(e => {
            console.log("Print Settings list already exist!");
            this.setState({
                loading: false
            });
        });
    }
}