import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, CheckboxVisibility } from 'office-ui-fabric-react/lib/DetailsList';
import IPrintTemplateProps from './print-template-props';
const _columns: IColumn[] = [
    {
        key: 'column1',
        name: 'Name',
        fieldName: 'name',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        ariaLabel: 'Operations for name'
    },
    {
        key: 'column2',
        name: 'Value',
        fieldName: 'value',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        ariaLabel: 'Operations for value'
    }
];
class PrintTemplateContent extends React.Component<IPrintTemplateProps,{}>{

    constructor(props){
        super(props);
    }
    public render(){
        return (
            <DetailsList
                items={this.props.items}
                columns={_columns}
                isHeaderVisible={false}
                setKey="set"
                layoutMode={DetailsListLayoutMode.fixedColumns}
                checkboxVisibility={CheckboxVisibility.hidden}
                selectionPreservedOnEmptyClick={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            />
        );
    }
}

export default PrintTemplateContent;