import * as React from 'react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { List } from 'office-ui-fabric-react/lib/List';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { CommandBarButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import ListService from '../../services/list-service';
import ISettingsPanelState from './settings-panel-state';
import ISettingsPanelProps from './settings-panel-props';
import styles from './settings-panel.module.scss';
import AddUpdateTemplate from '../add-update-template-panel/add-update-template-panel';

export default class SettingsPanel extends React.Component<ISettingsPanelProps, ISettingsPanelState>{
    private listService: ListService;

    constructor(props) {
        super(props);
        this.listService = new ListService();
        this.state = {
            templates: [],
            showTemplatePanel:false
        };
    }

    public async componentDidMount() {
        let templates: any[] = await this.listService.GetTemplatesByListId(this.props.listId);
        this.setState(
            {
                templates
            }
        );
    }

    public render() {
        
        return (
            <div className={styles.SettingsPanel}>
                <Panel isOpen={this.props.showPanel} onDismiss={this.props.setShowPanel(false)} type={PanelType.medium} headerText="Print Settings">
                    <h3>Print Templates:</h3>
                    <div style={{ display: 'flex', alignItems: 'stretch', height: '40px', marginBottom: '10px' }}>
                        <CommandBarButton
                            data-automation-id="test"
                            iconProps={{ iconName: 'Add' }}
                            text="Create template"
                            onClick={this._setShowTemplatePanel(true)}
                        />
                    </div>
                    <FocusZone direction={FocusZoneDirection.vertical}>
                        <List items={this.state.templates} onRenderCell={this._onRenderCell} />
                    </FocusZone>
                    <AddUpdateTemplate listId={this.props.listId} showTemplatePanel={this.state.showTemplatePanel} setShowTemplatePanel={this._setShowTemplatePanel} />
                </Panel>
            </div>
        );
    }

    public _setShowTemplatePanel = (showTemplatePanel: boolean): (() => void) => {
        return (): void => {
            this.setState({ showTemplatePanel });
        };
    }

    private _onRenderCell(item: any, index: number): JSX.Element {
        return (
            <div className={styles.SettingsPanel} data-is-focusable={true}>
                <div className={`${styles.itemCell} ${index % 2 === 0 && styles.itemCellEven}`} >
                    <div className={styles.itemTitle}>{item.Title}</div>
                    <div className={styles.cellIcons}>
                        <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit" />
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" />
                    </div>
                </div>
            </div>
        );
    }
}