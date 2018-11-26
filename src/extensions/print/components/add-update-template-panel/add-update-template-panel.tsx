import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import ReactQuill from 'react-quill';
import 'react-quill/dist/quill.snow.css';
import styles from './add-update-template.module.scss';
import { modules, formats } from './editor-toolbar';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import AddUpdateTemplatePanelState from './add-update-template-panel-state';
import AddUpdateTemplatePanelProps from './add-update-template-panel-props';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { DetailsList, Selection } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import ListService from '../../services/list-service';

export default class AddUpdateTemplate extends React.Component<AddUpdateTemplatePanelProps, AddUpdateTemplatePanelState> {
    private listService: ListService;
    constructor(props) {
        super(props);
        this.props.setShowTemplatePanel(this._onClosePanel.bind(this));
        this.listService = new ListService();
    }

    public async componentDidMount() {
        let { listId } = this.props;
        let fields: any[] = await this.listService.GetFieldsByListId(listId);
        console.log(fields);
    }

    public render() {
        return (
            <Panel
                isOpen={this.props.showTemplatePanel}
                type={PanelType.largeFixed}
                onDismiss={this._onClosePanel}
                isFooterAtBottom={true}
                headerText="Add/Update template"
                closeButtonAriaLabel="Close"
                onRenderFooterContent={this._onRenderFooterContent}
            >
                <div className={`${styles.AddUpdateTemplate} ms-Grid}`}>
                    <TextField label="Name" />
                    <Label>Header</Label>
                    <div className={styles.editorContainer}>
                        <ReactQuill modules={modules} formats={formats} className={styles.quillEditor} />
                    </div>
                    <Label>Footer</Label>
                    <div className={styles.editorContainer}>
                        <ReactQuill modules={modules} formats={formats} className={styles.quillEditor} />
                    </div>
                    <Label>Columns</Label>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">

                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                        </div>
                    </div>
                </div>
            </Panel>
        );
    }
    public _onClosePanel = () => {
        this.props.setShowTemplatePanel(false);
    }

    private _onRenderFooterContent = (): JSX.Element => {
        return (
            <div>
                <PrimaryButton onClick={this.props.setShowTemplatePanel(false)} style={{ marginRight: '8px' }}>
                    Save
            </PrimaryButton>
                <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
            </div>
        );
    }
}