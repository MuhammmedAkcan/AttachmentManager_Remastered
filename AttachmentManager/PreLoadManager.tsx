import * as React from "react";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { IAttachmentProps } from "./AttachmentManagerApp";
import { DefaultButton, ScrollablePane, ScrollbarVisibility, SearchBox, ShimmeredDetailsList, Stack, Sticky, StickyPositionType, MessageBar, MessageBarType } from "office-ui-fabric-react";
import { classNames } from './ComponentStyles';

export class PreLoadManager extends React.Component<IAttachmentProps> {
    constructor(props: IAttachmentProps) {
        super(props)
        
    }

    //Spinner effect
    public render(): React.JSX.Element { 
        console.log("PreLoadManager render: " + this.props.noFilesFound + " " + this.props.notSavedYet);
        console.log(this.props);

        if(this.props.notSavedYet){
            return this.NotSavedYet();
        }

        if(this.props.noFilesFound){
            return this.NoFilesFound();
        }
        else {
            return this.Spinner();
        }
    }

    public NoFilesFound(): React.JSX.Element {
        return (
            <div>
                <MessageBar messageBarType={MessageBarType.warning}>
                    No files found in this incident.
                </MessageBar>
            </div>
        )
    }

    public NotSavedYet(): React.JSX.Element {
        return (
            <div>
                <MessageBar messageBarType={MessageBarType.info}>
                    Save the E-Mail to load the documents of the incident.
                </MessageBar>
            </div>
        )
    }

    public Spinner(): React.JSX.Element {
        return (
            <div className={classNames.wrapper}>
                <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                    <Sticky stickyPosition={StickyPositionType.Header}>
                        <Stack horizontal tokens={{childrenGap: 20, padding:10}}>
                            <Stack.Item>
                                <DefaultButton text="Attach" />
                            </Stack.Item>
                            <Stack.Item grow align="stretch">
                                <SearchBox styles={{ root: { width: '100%' } }} placeholder="Search file" />
                            </Stack.Item>
                        </Stack>
                    </Sticky>

                    <Spinner size={SpinnerSize.large} label="Loading..." ariaLive="polite" labelPosition="right" />
                </ScrollablePane>
            </div>
            
        )
    }
}