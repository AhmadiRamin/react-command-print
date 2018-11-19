import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@uifabric/icons';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import styles from './print-dialog.module.scss';
import {
    Dropdown,
    DialogContent
} from 'office-ui-fabric-react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import {
    SPHttpClient,
    SPHttpClientResponse
} from '@microsoft/sp-http';
import {
    sp,
    ListEnsureResult
} from "@pnp/sp";

import IPrintDialogContentProps from './IPrintDialogContentProps';
import IPrintDialogContentState from './IPrintDialogContentState';

export default class PrintDialogContent extends React.Component<IPrintDialogContentProps, IPrintDialogContentState> {
    constructor(props) {
        super(props);
        initializeIcons();
        this.state = {
            loading: true,
            loadingMessage: "Loading...",
            printTemplates: []
        };
        // Getting configuration
        this.initializeSettings();
    }

    public render(): JSX.Element {
        return <div className={styles.PrintDialogContent}>
            <DialogContent
                title='Print List Item'
                onDismiss={this.props.close}
                showCloseButton={true}
            >
                <div className="ms-Grid" dir="ltr">
                    {
                        this.state.loading ?
                            <div className="ms-grid-row">
                                <Spinner size={SpinnerSize.large} label={this.state.loadingMessage} ariaLive="assertive" />
                            </div>
                            :
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg8">
                                    <span className={styles.templateText}>Template:</span>
                                    <Dropdown
                                        placeHolder="Select..."
                                        options={[
                                            { key: 'A', text: 'Template 1' },
                                            { key: 'B', text: 'Template 2' },
                                            { key: 'C', text: 'Template 3' }
                                        ]}
                                    />
                                </div>
                                <div className={styles.printIcons + " ms-Grid-col ms-sm6 ms-md4 ms-lg14"}>
                                    <IconButton iconProps={{ iconName: 'Print' }} title="Print" ariaLabel="Print" />
                                    <IconButton iconProps={{ iconName: 'Mail' }} title="Mail" ariaLabel="Mail" />
                                    <IconButton iconProps={{ iconName: 'PDF' }} title="PDF" ariaLabel="PDF" />
                                    <IconButton iconProps={{ iconName: 'ExcelLogo' }} title="Export to excel" ariaLabel="ExcelLogo" />
                                    <IconButton iconProps={{ iconName: 'Settings' }} title="Settings" ariaLabel="Settings" />
                                </div>
                            </div>
                    }
                </div>
            </DialogContent>;
        </div>;
    }

    public initializeSettings() {        
        // Check if Print Settings list exists, otherwise we are going to create it
        this.props.httpClient.get(this.props.webUrl + `/_api/web/lists/GetByTitle('Print Settings')/items`, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    // Perfect! Print Settings list exists
                }
                else {
                    // We need to add the Print Settings list
                    sp.web.lists.add('Print Settings', 'List of templates for Print Command Set extension', 100).then(result => {                        
                        //result.list.fields.addText('ListID');
                        //result.list.fields.addMultilineText('Header');
                        result.list.fields.createFieldAsXml('<Field Type="Note" DisplayName="Header" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" NumLines="6" RichText="FALSE" RichTextMode="Compatible" IsolateStyles="FALSE" Sortable="FALSE" ID="{6e971d30-8d31-4333-8735-b2c455432e03}" SourceID="{4f07c112-1d4c-4133-8539-9732fe069a9f}" StaticName="Header" Name="Header" CustomFormatter="" RestrictedMode="TRUE" AppendOnly="FALSE" UnlimitedLengthInDocumentLibrary="FALSE"></Field>')
                    });
                }
            }).catch(e => {
                console.log('error messsage');
            });
    }
}