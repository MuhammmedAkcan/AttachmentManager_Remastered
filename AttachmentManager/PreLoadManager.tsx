import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { EntityReference, PrimaryEntity, isInHarness, SharePointHelper } from "./PCFHelper";
import { http } from "./http";
import { IFileItem } from "./ItemList";
import { IconMapper } from "./IconMapper";
import { Email, SharePointDocument, ActivityMimeAttachment } from "./Entity";
import { TIMEOUT } from "dns";
import { isNullOrUndefined } from "util";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { IAttachmentProps } from "./AttachmentManagerApp";
import { ShimmeredDetailsList } from "office-ui-fabric-react";

export class PreLoadManager extends React.Component<IAttachmentProps> {

    constructor(props: IAttachmentProps) {
        super(props)

    }

    //Spinner effect
    public render(): React.JSX.Element { 
        return (
            <ShimmeredDetailsList
                enableShimmer={true}
                items={[]}
                columns={[
                    { key: 'iconclassname',
                    name: '',
                    fieldName: 'iconclassname',
                    minWidth: 20,
                    maxWidth: 40,
                    isResizable: false },
                    { key: 'fileName',
                    name: 'Document',
                    fieldName: 'fileName',
                    minWidth: 100,
                    maxWidth: 200,
                    isResizable: true }
                ]}
            />
        )
    }
}