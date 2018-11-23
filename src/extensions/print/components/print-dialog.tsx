import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import PrintDialogContent from './print-dialog-content';

class PrintDialog extends BaseDialog {    
    public webUrl: string;
    public listId: string;
    public render(): void {        
        ReactDOM.render(<PrintDialogContent
            close={this.close}
            webUrl={this.webUrl}
            listId={this.listId}
        />, this.domElement);
    }
}

export{
    PrintDialog
};