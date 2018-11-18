import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { initializeIcons } from '@uifabric/icons';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import styles from './print-dialog.module.scss';
import {
    autobind,
    PrimaryButton,
    Button,
    Dropdown,
    DialogFooter,
    DialogContent
} from 'office-ui-fabric-react';

interface IPrintDialogContentProps {
    close: () => void;
    submit: (color: string) => void;
}

class PrintDialogContent extends React.Component<IPrintDialogContentProps, {}> {
    private _pickedColor: string;

    constructor(props) {
        super(props);
        initializeIcons();
    }

    public render(): JSX.Element {
        return <div className={styles.PrintDialogContent}>
            <DialogContent
                title='Print List Item'
                onDismiss={this.props.close}
                showCloseButton={true}
            >
                <div className="ms-Grid" dir="ltr">
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg8">
                            <span className={styles.templateText}>Template:</span>
                        <Dropdown
                                placeHolder="Select..."
                                options={[
                                    { key: 'A', text: 'Option a' },
                                    { key: 'B', text: 'Option b' },
                                    { key: 'C', text: 'Option c' },
                                    { key: 'D', text: 'Option d' },
                                    { key: 'E', text: 'Option e' },
                                    { key: 'F', text: 'Option f' },
                                    { key: 'G', text: 'Option g' }
                                ]}
                            />
                        </div>
                        <div className={styles.printIcons +  " ms-Grid-col ms-sm6 ms-md4 ms-lg14"}>
                            <IconButton iconProps={{ iconName: 'Print' }} title="Print" ariaLabel="Print" />
                            <IconButton iconProps={{ iconName: 'Mail' }} title="Mail" ariaLabel="Mail" />
                            <IconButton iconProps={{ iconName: 'PDF' }} title="PDF" ariaLabel="PDF" />
                            <IconButton iconProps={{ iconName: 'ExcelLogo' }} title="ExcelLogo" ariaLabel="ExcelLogo" />
                        </div>
                    </div>
                </div>

                <DialogFooter>
                    <Button text='Cancel' title='Cancel' onClick={this.props.close} />
                    <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this._pickedColor); }} />
                </DialogFooter>
            </DialogContent>;
        </div>;
    }
}

export default class PrintDialog extends BaseDialog {
    public message: string;
    public colorCode: string;

    public render(): void {
        ReactDOM.render(<PrintDialogContent
            close={this.close}
            submit={this._submit}
        />, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: false
        };
    }

    @autobind
    private _submit(color: string): void {
        this.colorCode = color;
        this.close();
    }
}