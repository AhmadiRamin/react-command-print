import * as React from 'react';
import { initializeIcons } from '@uifabric/icons';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import styles from './print-dialog.module.scss';
import {
    Dropdown,
    DialogContent
} from 'office-ui-fabric-react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import IPrintDialogContentProps from './IPrintDialogContentProps';
import IPrintDialogContentState from './IPrintDialogContentState';
import ListHelper from '../util/list-helper';

export default class PrintDialogContent extends React.Component<IPrintDialogContentProps, IPrintDialogContentState> {
    constructor(props) {
        super(props);
        
        this.state = {
            loading: true,
            loadingMessage: "Loading...",
            printTemplates: []
        };

        // Initialize icons
        initializeIcons();

        // Validate and create Print Settings list
        const listHelper = new ListHelper(this.props.webUrl);
        listHelper.ValidatePrintSettingsList().then(_ => {
            this.setState({
                loadingMessage: 'Initializing settings',
                loading:true
            });
        }).catch(e => {
            console.log("Print Settings list already exist!");
            this.setState({
                loading:false
            });
        });
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
}