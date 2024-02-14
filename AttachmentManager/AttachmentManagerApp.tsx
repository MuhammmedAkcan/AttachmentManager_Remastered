import * as React from 'react';

import { DefaultButton, Stack, ProgressIndicator, Icon } from 'office-ui-fabric-react';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import {
    DetailsList,
    DetailsListLayoutMode,
    IDetailsHeaderProps,
    Selection,
    IColumn,
    ConstrainMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { IRenderFunction } from 'office-ui-fabric-react/lib/Utilities';
import { ScrollablePane, ScrollbarVisibility } from 'office-ui-fabric-react/lib/ScrollablePane';
import { Sticky, StickyPositionType } from 'office-ui-fabric-react/lib/Sticky';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { IFileItem, ItemList } from './ItemList';
import { classNames } from './ComponentStyles';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { AttachmentManager } from '.';

export interface IAttachmentProps {
    regardingObjectId: string;
    regardingEntityName: string;
    files: IFileItem[];
    onAttach: (selectedFiles: IFileItem[]) => Promise<void>;
    isControlLoading: boolean;
}

export interface IAttachmentState {
    files: IFileItem[];
    columns: IColumn[];
    minimalColumns: IColumn[];
    hiddenModal: boolean;
    isInProgress: boolean;
    isLoading: boolean;
}

export class AttachmentManagerApp extends React.Component<IAttachmentProps, IAttachmentState> {
    private selection: Selection;
    private allFiles: ItemList;
    

    constructor(props: IAttachmentProps) {
        super(props)

        initializeIcons();

        this.allFiles = new ItemList();

        this.selection = new Selection();

        this.allFiles.setItems(this.props.files);
        
        this.state = {
            files: this.allFiles.getItems(),
            hiddenModal: false,
            isInProgress: false,
            columns: this.allFiles.getColumns(),
            minimalColumns: this.allFiles.getMinimalColumns(),
            isLoading: true
        };

        this.attachFilesClicked = this.attachFilesClicked.bind(this);
        this.onFilterChanged = this.onFilterChanged.bind(this);
        this.onAttachClicked = this.onAttachClicked.bind(this);
        this.hideDialog = this.hideDialog.bind(this);
        this.resetProgress = this.resetProgress.bind(this);
    }

    public render(): React.JSX.Element { 
        const { hiddenModal: hiddenDialog, files, columns, minimalColumns, isLoading } = this.state;
        return (
            <div>
                {files.length == 0 ? (
                <Spinner size={SpinnerSize.large} label="Loading..." ariaLive="assertive" labelPosition="right" />
                ) : (
                    <div className={classNames.wrapper}>
                        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                            <Sticky stickyPosition={StickyPositionType.Header}>
                                <Stack horizontal tokens={{childrenGap: 20, padding:10}}>
                                    <Stack.Item>
                                        <DefaultButton text="Attach" onClick={this.onAttachClicked} />
                                    </Stack.Item>
                                    <Stack.Item grow align="stretch">
                                        <SearchBox styles={{ root: { width: '100%' } }} placeholder="Search file" onChange={this.onFilterChanged} />
                                    </Stack.Item>
                                </Stack>
                                <Stack>
                                    { this.state.isInProgress && <ProgressIndicator label="In progress" description="Copying files from SharePoint to an email" /> }
                                </Stack>
                            </Sticky>
                            <MarqueeSelection selection={this.selection}>
                                <DetailsList
                                    items={files}
                                    columns={minimalColumns}
                                    setKey="set"
                                    layoutMode={DetailsListLayoutMode.justified}
                                    constrainMode={ConstrainMode.unconstrained}
                                    onRenderItemColumn={renderItemColumn}
                                    onRenderDetailsHeader={onRenderDetailsHeader}
                                    selection={this.selection}
                                    selectionPreservedOnEmptyClick={true}
                                    ariaLabelForSelectionColumn="Toggle selection"
                                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                    onItemInvoked={this.onItemInvoked}
                                />
                            </MarqueeSelection>
                        </ScrollablePane>
                    </div>
                )}
            </div>
        );
    }
    

    

    private getItems = () => {
        return [
            {
                key: 'attachFile',
                name: 'Attach Files',
                cacheKey: 'myCacheKey',
                iconProps: {
                    iconName: 'Attach'
                },
                ariaLabel: 'Attach Files',
                onClick: this.attachFilesClicked
            }
        ]
    }

    private attachFilesClicked(): void {
        this.setState({ hiddenModal: false, isInProgress : false });
    }

    private onAttachClicked(): void {
        this.setState({isInProgress : true});
        this.props.onAttach(this.getSelectedFiles()).then(this.resetProgress);
    }

    private resetProgress(): void {
        console.log('Resetting progress');
        this.setState({isInProgress : false});
    }

    private hideDialog(): void {
        this.setState({hiddenModal:true});
    }

    private onItemInvoked(item: IFileItem): void {
        console.log('Item invoked: ' + item.fileName);
    }

    private onFilterChanged(ev?: React.ChangeEvent<HTMLInputElement>, text?: string): void {
        this.setState({
            files: text ? this.state.files.filter((item: IFileItem) => 
            hasText(item, text)) : this.state.files
        });
    }

    private getSelectedFiles(): IFileItem[] {
        let selectedFiles: IFileItem[] = [];

        for(let i = 0; i < this.selection.getSelectedCount(); i++) {
            selectedFiles.push((this.selection.getSelection()[i] as IFileItem));
        }

        return selectedFiles;
    }
}

function hasText(item: IFileItem, text: string): boolean {
    return `${item.id}|${item.fileName}|${item.fileType}`.indexOf(text) > -1;
}

function onRenderDetailsHeader(
    props?: IDetailsHeaderProps,
    defaultRender?: IRenderFunction<IDetailsHeaderProps>
): React.JSX.Element {
    return (
        <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced={true}>
            {defaultRender && defaultRender({ ...props! })}
        </Sticky>
    );
}

function renderItemColumn(item: IFileItem, index?: number, column?: IColumn) {
    if (column) {
        const fieldContent = item[column.fieldName as keyof IFileItem] as string;

        //console.log('Rendering column: ' + column.key + ' with content: ' + fieldContent);

        switch (column.key) {
            case 'iconclassname':
                return <Icon iconName={fieldContent} className={classNames.fileIcon}></Icon>;
            case 'createdOn':
            case 'sharepointcreatedon':
            case 'lastModifiedOn': {
                const dateField = item[column.fieldName as keyof IFileItem] as Date;
                //return <div>{dateField.toLocaleDateString('de-de')} {dateField.toLocaleTimeString('de-de')}</div>;
                const options = { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' } as Intl.DateTimeFormatOptions;
                const formatter = new Intl.DateTimeFormat('de-DE', options);
                return <div>{formatter.format(dateField)}</div>;
            }
            case 'fileName':{
                const dateField = item["lastModifiedOn" as keyof IFileItem] as Date;
                //return <div>{dateField.toLocaleDateString('de-de')} {dateField.toLocaleTimeString('de-de')}</div>;
                const options = { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' } as Intl.DateTimeFormatOptions;
                const formatter = new Intl.DateTimeFormat('de-DE', options);

                if(item["subject" as keyof IFileItem] as string != "---"){
                    return <div><b>{fieldContent}</b><div>{formatter.format(dateField)} - {item["subject" as keyof IFileItem] as string}</div></div>
                } else {
                    return <div><b>{fieldContent}</b><div>{formatter.format(dateField)}</div></div>
                }
            }
            case 'fileType':
            case 'lastModifiedBy':
            case 'subject':
            case 'version':
                return <div>{fieldContent}</div>;
            case 'directioncode':
                return <div>{fieldContent ? 'Incoming' : 'Outgoing'}</div>;
            default:{
                console.log('Unknown column: ' + column.key);
                break;
            }
        }
    }


    
}