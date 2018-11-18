import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import PrintDialogContent from './components/print-dialog';
import * as strings from 'PrintCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPrintCommandSetProperties {
  // This is an example; replace with your own properties
  printText: string;
}

const LOG_SOURCE: string = 'PrintCommandSet';

export default class PrintCommandSet extends BaseListViewCommandSet<IPrintCommandSetProperties> {

  private _colorCode: string;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized PrintCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const printCommand: Command = this.tryGetCommand('COMMAND_Print');
    if (printCommand) {
      // This command should be hidden unless exactly one row is selected.
      printCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_Print':
      const dialog: PrintDialogContent = new PrintDialogContent();
      dialog.show();
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
