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
import { IDragDropEvents } from 'office-ui-fabric-react/lib/utilities/dragdrop/interfaces';
import { DetailsList, IColumn, Selection, DetailsListLayoutMode, IDetailsRowProps, SelectionMode, DetailsRow } from 'office-ui-fabric-react/lib/DetailsList';
import { IColumnReorderOptions } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import ListService from '../../services/list-service';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { ColorPicker } from 'office-ui-fabric-react/lib/ColorPicker';
let _draggedItem: any = null;
let _draggedIndex = -1;

export default class AddUpdateTemplate extends React.Component<AddUpdateTemplatePanelProps, AddUpdateTemplatePanelState> {
    private listService: ListService;
    private _fieldSelection: Selection;
    private _itemSelection: Selection;
    private _columns: IColumn[] = [
        {
            key: 'Title',
            name: 'Field',
            fieldName: 'Title',
            minWidth: 120,
            isResizable: true,
            ariaLabel: 'Operations for Field'
        }];
    private _itemColumns: IColumn[] = [
        {
            key: 'Title',
            name: 'Field',
            fieldName: 'Title',
            minWidth: 90,
            isResizable: true,
            ariaLabel: 'Operations for Field'
        }, {
            key: 'manage',
            name: 'Manage',
            fieldName: '',
            minWidth: 50,
            isResizable: false
        }];
    private _defautState: any;
    private _defaultColor: string;

    constructor(props) {
        super(props);
        this.props.setShowTemplatePanel(this._onClosePanel.bind(this));
        this.listService = new ListService();
        this._defautState = {
            helperItems: [{
                Title: 'Drag your fields here'
            }],
            listId: this.props.listId,
            templateColumns: [],
            section: '',
            columns: this._columns,
            itemColumns: this._itemColumns,
            isColumnReorderEnabled: false,
            frozenColumnCountFromStart: '1',
            frozenColumnCountFromEnd: '0',
            showColorPicker: false,
            color: '#eeeeee'
        };
        this.state = {
            ...this._defautState,
            fields: []
        };
        this._defaultColor = '#eeeeee';
        this._fieldSelection = new Selection();
        this._itemSelection = new Selection();
        this._renderItemColumn = this._renderItemColumn.bind(this);
        this._onRemoveItem = this._onRemoveItem.bind(this);
        this._closeColorPicker=this._closeColorPicker.bind(this);
        this._onColorChange=this._onColorChange.bind(this);
        this._onColorSelected=this._onColorSelected.bind(this);
        this._openColorPicker=this._openColorPicker.bind(this);
    }

    public async componentDidMount() {
        let fields: any[] = await this.listService.GetFieldsbyListId(this.props.listId);
        this.setState({
            fields
        });
    }

    public render() {
        const { fields, columns, itemColumns, helperItems } = this.state;
        const items = this.props.template.Columns;

        return (
            <div>
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
                        <TextField value={this.props.template.Title} label="Name" onChanged={(name) => this.props.onTemplateChanged({ ...this.props.template, Title: name })} />
                        <Label>Columns (Drag fields from the left table to the right one)</Label>
                        <div className="ms-Grid-row">
                            <div className={`ms-Grid-col ms-sm6 ms-md6 ms-lg6`}>
                                <MarqueeSelection selection={this._fieldSelection}>
                                    <DetailsList
                                        className={styles.detailsList}
                                        isHeaderVisible={false}
                                        layoutMode={DetailsListLayoutMode.fixedColumns}
                                        setKey={'fields'}
                                        items={fields}
                                        columns={columns}
                                        selection={this._fieldSelection}
                                        selectionPreservedOnEmptyClick={true}
                                        dragDropEvents={this._getFieldsDragEvents()}
                                        ariaLabelForSelectionColumn="Toggle selection"
                                        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                    />
                                </MarqueeSelection>
                            </div>
                            <div className={`ms-Grid-col ms-sm6 ms-md6 ms-lg6`}>
                                {
                                    items.length > 0 ?
                                        <MarqueeSelection selection={this._itemSelection}>
                                            <DetailsList
                                                className={styles.detailsList}
                                                isHeaderVisible={false}
                                                layoutMode={DetailsListLayoutMode.justified}
                                                setKey={'items'}
                                                items={items}
                                                columns={itemColumns}
                                                selection={this._itemSelection}
                                                selectionPreservedOnEmptyClick={true}
                                                onRenderItemColumn={this._renderItemColumn}
                                                onRenderRow={this._renderRow}
                                                dragDropEvents={this._getDragDropEvents()}
                                                columnReorderOptions={this.state.isColumnReorderEnabled ? this._getColumnReorderOptions() : undefined}
                                                ariaLabelForSelectionColumn="Toggle selection"
                                                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                            />
                                        </MarqueeSelection>
                                        :
                                        <DetailsList
                                            className={styles.detailsList}
                                            isHeaderVisible={false}
                                            items={helperItems}
                                            columns={columns}
                                            selectionMode={SelectionMode.none}
                                            selectionPreservedOnEmptyClick={false}
                                            dragDropEvents={this._getDragDropEvents()}
                                        />
                                }

                            </div>
                        </div>
                        <Label>Add section</Label>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm10 ms-md10 ms-lg10">
                                <TextField onChanged={(value) => this.setState({ section: value })} value={this.state.section} />

                            </div>
                            <div className="ms-Grid-col ms-sm1 ms-md2 ms-lg2">
                                <IconButton iconProps={{ iconName: 'Color' }} title="Change color" ariaLabel="Change Color" onClick={this._openColorPicker} />
                                <IconButton iconProps={{ iconName: 'Accept' }} title="Accept" ariaLabel="Accept" onClick={this._addSection} />
                            </div>
                        </div>
                        <Label>Header</Label>
                        <div className={styles.editorContainer}>
                            <ReactQuill modules={modules} formats={formats} className={styles.quillEditor} value={this.props.template.Header} onChange={(Header) => this.props.onTemplateChanged({ ...this.props.template, Header })} />
                        </div>
                        <Label>Footer</Label>
                        <div className={styles.editorContainer}>
                            <ReactQuill modules={modules} formats={formats} className={styles.quillEditor} value={this.props.template.Footer} onChange={(Footer) => this.props.onTemplateChanged({ ...this.props.template, Footer })} />
                        </div>

                    </div>
                    <Dialog
                    onDismissed={this._closeColorPicker}
                    isClickableOutsideFocusTrap={true}                    
                    isOpen={this.state.showColorPicker}
                    ignoreExternalFocusing={true}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Color picker',
                        showCloseButton:true
                    }}
                    modalProps={{
                        titleAriaId: 'myLabelId',
                        ignoreExternalFocusing:true,

                        subtitleAriaId: 'mySubTextId',
                        isBlocking: true,
                        containerClassName: 'ms-dialogMainOverride'
                    }}
                >
                        <ColorPicker color={this.state.color} onColorChanged={this._onColorChange} />
                    <DialogFooter>
                        <PrimaryButton onClick={this._onColorSelected} text="OK" />
                        <DefaultButton onClick={this._closeColorPicker} text="Cancel" />
                    </DialogFooter>
                </Dialog>
                </Panel>
                
            </div>

        );
    }

    private _openColorPicker(){
        this.setState({
            showColorPicker: true
        });
    }
    private _onColorChange(color: string): void {
        this._defaultColor = color;
    }

    private _closeColorPicker() {
        this.setState({
            showColorPicker: false
        });
    }

    private _onColorSelected() {
        this.setState({
            showColorPicker: false,
            color: this._defaultColor
        });
    }

    private _addSection = () => {
        const newSection = {
            Title: this.state.section,
            Type: 'Section',
            Id: this.state.section
        };
        this.setState({
            section: ''
        });
        this.props.onTemplateChanged(
            {
                ...this.props.template,
                Columns: this.props.template.Columns.concat(newSection)
            }
        );

    }

    public _onClosePanel = () => {
        this.setState({ ...this._defautState });
        this.props.setShowTemplatePanel(false)();
    }

    private _onRenderFooterContent = (): JSX.Element => {
        return (
            <div>
                <PrimaryButton onClick={this.props.onTemplateSaved} style={{ marginRight: '8px' }}>Save</PrimaryButton>
                <DefaultButton onClick={() => this._onClosePanel()}>Cancel</DefaultButton>
            </div>
        );
    }

    private _renderRow = (props: IDetailsRowProps, defaultRender?: any) => {

        if (props.item.Type === 'Section')
            return <DetailsRow {...props} className={styles.sectionRow} />;
        else
            return <DetailsRow {...props} />;
    }

    private _renderItemColumn(item: any, index: number, column: IColumn) {
        const fieldContent = item[column.fieldName || ''];
        switch (column.key) {
            case 'manage':
                return <IconButton className={styles.removeIconContainer} onClick={() => this._onRemoveItem(item)} iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" />;
            default:
                return <span>{fieldContent}</span>;
        }
    }

    private _onRemoveItem = (item: any) => {
        this.props.onTemplateChanged({
            ...this.props.template,
            Columns: this.props.template.Columns.filter(el => el != item)
        });

    }
    // Details list methods

    private _handleColumnReorder = (draggedIndex: number, targetIndex: number) => {
        const draggedItems = this.state.columns[draggedIndex];
        const newColumns: IColumn[] = [...this.state.columns];

        // insert before the dropped item
        newColumns.splice(draggedIndex, 1);
        newColumns.splice(targetIndex, 0, draggedItems);
        this.setState({ columns: newColumns });
    }

    private _getColumnReorderOptions(): IColumnReorderOptions {
        return {
            frozenColumnCountFromStart: parseInt(this.state.frozenColumnCountFromStart, 10),
            frozenColumnCountFromEnd: parseInt(this.state.frozenColumnCountFromEnd, 10),
            handleColumnReorder: this._handleColumnReorder
        };
    }

    private _getFieldsDragEvents(): IDragDropEvents {
        return {
            canDrop: () => {
                return false;
            },
            canDrag: () => {
                return true;
            },
            onDragEnter: () => {
                return 'dragEnter';
            }, // return string is the css classes that will be added to the entering element.
            onDragLeave: () => {
                return;
            },
            onDragStart: (item?: any, itemIndex?: number, selectedItems?: any[], event?: MouseEvent) => {
                _draggedItem = item;
                _draggedIndex = itemIndex!;
            },
            onDragEnd: (item?: any, event?: DragEvent) => {
                _draggedItem = null;
                _draggedIndex = -1;
            }
        };
    }

    private _getDragDropEvents(): IDragDropEvents {
        return {
            canDrop: () => {
                return true;
            },
            canDrag: () => {
                return true;
            },
            onDragEnter: () => {
                return 'dragEnter';
            }, // return string is the css classes that will be added to the entering element.
            onDragLeave: () => {
                return;
            },
            onDrop: (item?: any, event?: DragEvent) => {
                if (_draggedItem) {
                    this._insertBeforeItem(item);
                }
            },
            onDragStart: (item?: any, itemIndex?: number, selectedItems?: any[], event?: MouseEvent) => {

                _draggedItem = item;
                _draggedIndex = itemIndex!;
            },
            onDragEnd: (item?: any, event?: DragEvent) => {

                _draggedItem = null;
                _draggedIndex = -1;
            }
        };
    }

    private _insertBeforeItem(item: any): void {
        const draggedItems = this._fieldSelection.isIndexSelected(_draggedIndex) ? this._fieldSelection.getSelection() : [_draggedItem];

        const items: any[] = this.props.template.Columns.filter((i: number) => draggedItems.indexOf(i) === -1);
        let insertIndex = items.indexOf(item);
        // if dragging/dropping on itself, index will be 0.
        if (insertIndex === -1) {
            insertIndex = 0;
        }
        items.splice(insertIndex, 0, ...draggedItems);
        this._fieldSelection.setItems([]);
        this.props.onTemplateChanged({
            ...this.props.template,
            Columns: items
        });
    }
}