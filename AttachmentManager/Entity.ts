import { IInputs } from "./generated/ManifestTypes";
import { FetchXML, EntityReference } from "./PCFHelper";

export const ActivityMimeAttachment = {
    EntityName: "activitymimeattachment",
    create: async(content: any, regarding: EntityReference, name: string, context: ComponentFramework.Context<IInputs>) => {
        
        var attachment = {} as any;
		
        attachment.body = content;
        attachment["objectid_activitypointer@odata.bind"] = `activitypointers(${regarding.id})`;
        attachment["objecttypecode"] = regarding.typeName;
        attachment["filename"] = name;

		// console.log("Regarding ID of Mime" + regarding.id);
		// console.log("Regarding Type of Mime" + regarding.typeName);
		// console.log("Attachment: " + name);
		// console.log("Content: " + content);

        await context.webAPI.createRecord(ActivityMimeAttachment.EntityName, attachment);
    }
}

export const Email = {
    EntityName: "email",
    RegardingObject: "regardingobjectid",
    getById: async(id: string, context: ComponentFramework.Context<IInputs>) => {
		if(id == null || id == undefined || id == "") return null;
        const email = await context.webAPI.retrieveRecord(Email.EntityName, id);
        return email;
    },

	//Replace {0} with regardingobjectid
	FetchXML:`
	<fetch distinct="false" mapping="logical" returntotalrecordcount="true" page="1" count="100" no-lock="false">
		<entity name="email">
			<attribute name="createdby" />
			<attribute name="createdon" />
			<attribute name="directioncode" />
			<attribute name="subject" />
			<attribute name="activityid" />
			<filter>
				<condition attribute="regardingobjectid" operator="eq" value="{0}" />
			</filter>
		</entity>
	</fetch>`,
	
	Subject: "subject",
	Directioncode: "directioncode",

	/* Das könnte gehen! */
	getAllEmails: async(context: ComponentFramework.Context<IInputs>, id:string) => {
		const fetchXml: string = Email.FetchXML.replace("{0}", id);
		const emails = await context.webAPI.retrieveMultipleRecords(Email.EntityName, FetchXML.prepareOptions(fetchXml));
		return emails;
	}
}



export const SharePointDocument = {
    EntityName: "sharepointdocument",
    FetchXml: `
		<fetch distinct="false" mapping="logical" returntotalrecordcount="true" page="1" count="100" no-lock="false">
			<entity name="sharepointdocument">
				<attribute name="regardingobjectid" />
				<attribute name="version" />
				<attribute name="versionnumber" />
				<attribute name="absoluteurl" />
				<attribute name="author" />
				<attribute name="businessunitid" />
				<attribute name="checkedoutto" />
				<attribute name="checkincomment" />
				<attribute name="childfoldercount" />
				<attribute name="childitemcount" />
				<attribute name="contenttype" />
				<attribute name="contenttypeid" />
				<attribute name="copysource" />
				<attribute name="createdon" />
				<attribute name="documentid" />
				<attribute name="documentlocationtype" />
				<attribute name="edit" />
				<attribute name="editurl" />
				<attribute name="exchangerate" />
				<attribute name="filesize" />
				<attribute name="filetype" />
				<attribute name="fullname" />
				<attribute name="iconclassname" />
				<attribute name="ischeckedout" />
				<attribute name="isfolder" />
				<attribute name="isrecursivefetch" />
				<attribute name="itvt_weclapp_last_modified_date" />
				<attribute name="locationid" />
				<attribute name="locationname" />
				<attribute name="modified" />
				<attribute name="modifiedon" />
				<attribute name="msft_datastate" />
				<attribute name="organizationid" />
				<attribute name="ownerid" />
				<attribute name="owningbusinessunit" />
				<attribute name="readurl" />
				<attribute name="relativelocation" />
				<attribute name="servicetype" />
				<attribute name="sharepointcreatedon" />
				<attribute name="sharepointdocumentid" />
				<attribute name="sharepointmodifiedby" />
				<attribute name="title" />
				<attribute name="transactioncurrencyid" />
				<order attribute="relativelocation" descending="false"/>
				<filter>
					<condition attribute="isrecursivefetch" operator="eq" value="0"/>
				</filter>
				<link-entity name="{entityName}" from="{entityNameID}id" to="regardingobjectid" alias="bb">
					<filter type="and">
						<condition attribute="{entityNameID}id" operator="eq" uitype="{entityNameID}id" value="{id}"/>
					</filter>
				</link-entity>
			</entity>
        </fetch>`,
        FullName : "fullname",
        AbsoluteUrl : "absoluteurl",
        FileType: "filetype",
        IconClassName: "iconclassname",
        LastModifiedOn: "modified",
		LastModifiedBy: "sharepointmodifiedby", 
		Version: "version",
		CreatedOn: "sharepointcreatedon",
		RegardingID: "regardingobjectid",

        getByRegarding: async(id: string, entityName: string, context: ComponentFramework.Context<IInputs>) => {
            const fetchXml: string = (SharePointDocument.FetchXml as string).split("{entityName}").join(entityName).split("{entityNameID}").join(entityName).split("{id}").join(id);

            const documents = await context.webAPI.retrieveMultipleRecords(SharePointDocument.EntityName, FetchXML.prepareOptions(fetchXml));
            return documents.entities;
        },

		getByEmailRegarding: async(id: string, context: ComponentFramework.Context<IInputs>) => {
			const fetchXml: string = (SharePointDocument.FetchXml as string).split("{entityName}").join("email").split("{entityNameID}").join("activity").split("{id}").join(id);

            const documents = await context.webAPI.retrieveMultipleRecords(SharePointDocument.EntityName, FetchXML.prepareOptions(fetchXml));
            return documents.entities;
		}

}