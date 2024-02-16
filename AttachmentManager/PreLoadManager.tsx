import * as React from "react";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { IAttachmentProps } from "./AttachmentManagerApp";
import { DefaultButton, ScrollablePane, ScrollbarVisibility, SearchBox, ShimmeredDetailsList, Stack, Sticky, StickyPositionType } from "office-ui-fabric-react";
import { classNames } from './ComponentStyles';

export class PreLoadManager extends React.Component<IAttachmentProps> {
    constructor(props: IAttachmentProps) {
        super(props)

    }

    //Spinner effect
    public render(): React.JSX.Element { 
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