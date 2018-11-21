import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import PrintDialogContent from './print-dialog-content';

class PrintDialog extends BaseDialog {
    public message: string;
    public webUrl: string;
    public render(): void {        
        ReactDOM.render(<PrintDialogContent
            close={this.close}
            webUrl={this.webUrl}
        />, this.domElement);
    }
}

export{
    PrintDialog
};