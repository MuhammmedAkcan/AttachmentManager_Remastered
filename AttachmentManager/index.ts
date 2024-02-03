import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { AttachmentManagerApp, IAttachmentProps } from "./AttachmentManagerApp";
import { EntityReference, PrimaryEntity, isInHarness, SharePointHelper } from "./PCFHelper";
import { http } from "./http";
import { IFileItem } from "./ItemList";
import { IconMapper } from "./IconMapper";
import { Email, SharePointDocument, ActivityMimeAttachment } from "./Entity";
import { TIMEOUT } from "dns";

export class AttachmentManager implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private container: HTMLDivElement;
	private context: ComponentFramework.Context<IInputs>;

	private primaryEntity: PrimaryEntity;
	private regardingId: string;
	private notifyOutputChanged: () => void;

	private iconMapper: IconMapper;
	private spHelper: SharePointHelper;

	

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	private timeout(ms:number){
		return new Promise(resolve => setTimeout(resolve, ms));
	}

	private async onAttach(selectedFiles: IFileItem[]) {
		let apiUrl: string;
		for (let i = 0; i < selectedFiles.length; i++) {
			const fileUrl = selectedFiles[i].fileUrl;
			console.log(fileUrl);

			apiUrl = this.spHelper.makeApiUrl(fileUrl);

			console.log(apiUrl);

			const data = await http(apiUrl);

			ActivityMimeAttachment.create(data["Content"],
			this.primaryEntity.Entity, fileUrl.substr(fileUrl.lastIndexOf('/')), this.context);
		}

		//email|NoRelationship|Form|Mscrm.SavePrimary21-button
		try {
			let attachButton = document.getElementById("email|NoRelationship|Form|Mscrm.SavePrimary20-button");
			if(!attachButton) attachButton = document.getElementById("email|NoRelationship|Form|Mscrm.SavePrimary211-button");

			if (attachButton) {
				await this.timeout(500);
				attachButton.click();
				console.log("attachButton clicked");
			} else {
				console.log("attachButton not found");
			}
		} catch (error) {
			//console.log(error);
			
		}
		
	}

	private renderControl(ec: ComponentFramework.WebApi.Entity[], ecEmails: ComponentFramework.WebApi.Entity[], isFileFromEmailArray: boolean[]): void {
		console.log("renderControl");


		let props: IAttachmentProps = {} as IAttachmentProps;
		props.files = [];
		props.onAttach = this.onAttach.bind(this);

		let emailIndex = 0;

		for (let i = 0; i < ec.length; i++) {
			console.log(ec[i]);

			let file:IFileItem;
			if(!isFileFromEmailArray[i]) {
				console.log("Email not found");

				file = {
					key: i,
					id: i.toString(),
					fileName: ec[i][SharePointDocument.FullName],
					fileUrl: ec[i][SharePointDocument.AbsoluteUrl],
					fileType: ec[i][SharePointDocument.FileType],
					iconclassname: this.iconMapper.getBySharePointIcon(ec[i][SharePointDocument.IconClassName]),
					lastModifiedOn: new Date(ec[i][SharePointDocument.LastModifiedOn]),
					lastModifiedBy: ec[i][SharePointDocument.LastModifiedBy],
					sharepointcreatedon: new Date(ec[i][SharePointDocument.CreatedOn]),
					version: ec[i][SharePointDocument.Version],
					subject: "---",
					directioncode: false
				};
			} else {
				console.log("Email found");
				console.log(ecEmails[emailIndex]);

				file = {
					key: i,
					id: i.toString(),
					fileName: ec[i][SharePointDocument.FullName],
					fileUrl: ec[i][SharePointDocument.AbsoluteUrl],
					fileType: ec[i][SharePointDocument.FileType],
					iconclassname: this.iconMapper.getBySharePointIcon(ec[i][SharePointDocument.IconClassName]),
					lastModifiedOn: new Date(ec[i][SharePointDocument.LastModifiedOn]),
					lastModifiedBy: ec[i][SharePointDocument.LastModifiedBy],
					sharepointcreatedon: new Date(ec[i][SharePointDocument.CreatedOn]),
					version: ec[i][SharePointDocument.Version],
					subject: ecEmails[emailIndex][Email.Subject],
					directioncode: ecEmails[emailIndex][Email.Directioncode]
				};

				console.log("ACHTUNG");
				console.log(file);

				//emailIndex++; this needs fixing!!!
			}

			props.files.push(file);
		}

		ReactDOM.render(
			React.createElement(AttachmentManagerApp, props)
			, this.container
		);
	}

	private renderControlWithMockData(): void {
		console.log("renderControl");
		let props: IAttachmentProps = {} as IAttachmentProps;
		ReactDOM.render(
			React.createElement(AttachmentManagerApp, props)
			, this.container
		);
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		this.context = context;
		this.container = container;
		this.notifyOutputChanged = notifyOutputChanged;

		this.primaryEntity = new PrimaryEntity(this.context);
		this.iconMapper = new IconMapper();

		this.spHelper = new SharePointHelper(this.context.parameters.SharePointSiteURLs.raw as string, this.context.parameters.FlowURL.raw as string);
	}



	public async updateView(context: ComponentFramework.Context<IInputs>): Promise<void> {
		this.context = context;
		this.primaryEntity = new PrimaryEntity(this.context);
	
		if (isInHarness()) {
			this.renderControlWithMockData();
		} else {
			let aggregatedFiles: ComponentFramework.WebApi.Entity[] = [];
			let aggregatedEmails: ComponentFramework.WebApi.Entity[] = [];
			let isFileFromEmailArray: boolean[] = [];
			let regardingEntity: ComponentFramework.WebApi.Entity = {} as ComponentFramework.WebApi.Entity;

			console.log("Getting Regarding");
			await Email.getById(this.primaryEntity.Entity.id, this.context).then(
				async (e) => {
					if (!e) return;
					const regarding: EntityReference = EntityReference.get(e, Email.RegardingObject)
					regardingEntity = regarding;
	
					const ec = await SharePointDocument.getByRegarding(regarding.id, regarding.typeName, this.context);
					console.log(`No. of documents in Regarding ${ec.length}`);
	
					if (ec.length > 0) {
						aggregatedFiles = [...aggregatedFiles, ...ec];
						ec.forEach(() => {
							isFileFromEmailArray.push(false);
						});
					}
				}
			)
	
			console.log("Getting Emails");
			await Email.getAllEmails(this.context, regardingEntity.id).then(
				async (emails) => {
					console.log(`No. of emails ${emails.entities.length}`);
	
					const promises = emails.entities.map(e =>
						SharePointDocument.getByEmailRegarding(e["activityid"], this.context).then(
							(ec) => {
								console.log(`No. of documents in E-Mail ${ec.length}`);
								if (ec.length > 0) {
									aggregatedFiles = [...aggregatedFiles, ...ec];
									aggregatedEmails = [...aggregatedEmails, e];
									ec.forEach(() => {
										isFileFromEmailArray.push(true);
									});
								}
							}
						)
					);
					await Promise.all(promises);
				}
			)
	
			//await Promise.all([sharePointDocPromise, emailDocPromise]);
	
			console.log("All promises resolved");
			console.log(`No. of documents in total ${aggregatedFiles.length}`);
			console.log(`No. of emails in total ${aggregatedEmails.length}`);
			console.log(`No. of isFileFromEmailArray in total ${isFileFromEmailArray.length}`);

			if(isFileFromEmailArray.length == aggregatedFiles.length - aggregatedEmails.length) {
				console.log("isFileFromEmailArray is correct");
			} else {
				console.log("isFileFromEmailArray is not correct");
			}
	
			if (aggregatedFiles.length > 0) {
				this.renderControl(aggregatedFiles, aggregatedEmails, isFileFromEmailArray);
			}
		}
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		this.context = context;

		this.primaryEntity = new PrimaryEntity(this.context);

		

		if (isInHarness()) {
			this.renderControlWithMockData();
		} else {
			let aggregatedFiles:ComponentFramework.WebApi.Entity[] = [];
			let aggregatedEmails:ComponentFramework.WebApi.Entity[] = [];
			let isFileFromEmailArray: boolean[] = [];

			//Get SharePoint Document from "Regarding" (in our case it's incident/case)
			const sharePointDocPromise = Email.getById(this.primaryEntity.Entity.id, this.context).then(
				(e) => {
					if(!e) return;
					const regarding: EntityReference = EntityReference.get(e, Email.RegardingObject)

					SharePointDocument.getByRegarding(regarding.id, regarding.typeName, this.context).then(
						(ec) => {
							console.log(`No. of documents in SP ${ec.length}`);
							
							if(ec.length > 0) {
								aggregatedFiles = [...aggregatedFiles, ...ec];

								ec.forEach(element => {
									isFileFromEmailArray.push(false);
								});
							}
						}
					);
				}
			)

			//Get All Emails
			const emailDocPromise = Email.getAllEmails(this.context, this.primaryEntity.Entity.id).then(
				(emails) => {
					console.log(`No. of emails ${emails.entities.length}`);
					
					const promises = emails.entities.map(e => {
						return SharePointDocument.getByEmailRegarding(e["activityid"], this.context).then(
							(ec) => {
								console.log(`No. of documents in E-Mails ${ec.length}`);
								if(ec.length > 0) {
									aggregatedFiles = [...aggregatedFiles, ...ec];
									aggregatedEmails = [...aggregatedEmails, e];

									ec.forEach(element => {
										isFileFromEmailArray.push(true);
									});
								}
							}
						)
					});
					Promise.all(promises).then(() => {});

					/*
					Promise.all(promises).then(() => {
						console.log(`No. of documents in total ${aggregatedFiles.length}`);

						if(aggregatedFiles.length > 0) {
							this.renderControl(aggregatedFiles, aggregatedEmails, isFileFromEmailArray);
						}
					});
					
				}
			)

			Promise.all([sharePointDocPromise, emailDocPromise]).then(() => {
				console.log("All promises resolved");
				console.log(`No. of documents in total ${aggregatedFiles.length}`);
				console.log(`No. of emails in total ${aggregatedEmails.length}`);
				console.log(`No. of isFileFromEmailArray in total ${isFileFromEmailArray.length}`);

				if(aggregatedFiles.length > 0) {
					this.renderControl(aggregatedFiles, aggregatedEmails, isFileFromEmailArray);
				}
			});
		}
	}
	*/

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {
			RegardingId: this.regardingId
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		ReactDOM.unmountComponentAtNode(this.container);
	}
}