import {
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodePropertyOptions,
	INodeType,
	INodeTypeDescription,
	IRequestOptions,
} from 'n8n-workflow';


async function makeMicrosoftRequest(thisRef: IExecuteFunctions | ILoadOptionsFunctions, resource: string, options: IRequestOptions = {}): Promise<any> {
	let url = resource;
	
	if(!url.startsWith('https://')) {
		url = "https://graph.microsoft.com/v1.0/" + resource;
	}

	const reqOptions: IRequestOptions = {
		method: 'GET',
		uri: url,
		headers: {
			"Content-Type": "application/json",
		},
		body: {},
		json: true,
		...options
	};

	return await thisRef.helpers.requestOAuth2.call(thisRef, 'microsoftSharepointOAuth2Api', reqOptions);
}

async function MSGetSites(thisRef: IExecuteFunctions | ILoadOptionsFunctions): Promise<any> {
	return await makeMicrosoftRequest(thisRef, 'sites', {
		qs: {
			search: '*',
		}
	});
}

async function MSGetSiteDrives(thisRef: IExecuteFunctions | ILoadOptionsFunctions, siteId: string): Promise<any> {
	return await makeMicrosoftRequest(thisRef, `sites/${siteId}/drives`);
}


export class Sharepoint implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Sharepoint',
		name: 'sharepoint',
		icon: 'file:Sharepoint.svg',
		group: ['transform'],
		version: 1,
		description: 'Interact with Sharepoint',
		defaults: {
			name: 'Sharepoint',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [
			{
				name: 'microsoftSharepointOAuth2Api',
				required: true,
			},
		],
		properties: [
			// ---- Grouping resources we can interact with (File, Folder, Site)
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'File',
						value: 'file',
					},
					{
						name: 'Site',
						value: 'site',
					},
				],
				default: 'file',
			},
			// File Operations
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				displayOptions: {
					show: {
						resource: ['file'],
					},
				},
				options: [
					{
						name: 'Get items in a folder',
						action: 'Get items in folder',
						value: 'getItemsInFolder',
					},
					{
						name: 'Get File',
						action: 'Get File',
						value: 'getFile',
					},
					{
						name: 'Upload File',
						action: 'Upload File',
						value: 'uploadFile',
					},
				],
				default: 'getFile',
				required: true,
				noDataExpression: true,
				description: 'The operation to perform on the file',
			},
			// --------------- Site Actions ------------------
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				displayOptions: {
					show: {
						resource: ['site'],
					},
				},
				options: [
					{
						name: 'Get Sites',
						action: 'Get Sites',
						value: 'getSites',
					},
				],
				default: 'getSites',
				noDataExpression: true,
				required: true,
			},

			// ---------------- Parameters -------------------
			{
				displayName: 'Site ID',
				name: 'siteId',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getSites',
				},
				default: '',
				required: true,
				displayOptions: {
					show: {
						operation: ['getFile', 'uploadFile', 'getItemsInFolder'],
					},
				},
			},
			{
				displayName: 'Document Library ID',
				name: 'libraryId',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getSiteDrives',
					loadOptionsDependsOn: ['siteId'],
				},
				default: '',
				required: true,
				displayOptions: {
					show: {
						operation: ['getFile', 'uploadFile', 'getItemsInFolder'],
					},
				},
			},
			{
				displayName: 'File Path',
				name: 'filePath',
				type: 'string',
				default: '',
				required: true,
				displayOptions: {
					show: {
						operation: ['getFile', 'uploadFile', 'getItemsInFolder'],
					},
				},
			},
			{
				displayName: 'File name',
				name: 'fileName',
				type: 'string',
				default: '',
				required: true,
				displayOptions: {
					show: {
						operation: ['uploadFile'],
					},
				},
			},
			{
				displayName: 'Binary Property',
				name: 'binaryPropertyName',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: {
					show: {
						operation: ['uploadFile'],
					},
				},
				description: 'Name of the binary property which contains the data for the file to be uploaded',
			},
		],
	};

	methods = {
		loadOptions: {
			async getSites(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const sites = await MSGetSites(this);
				return sites.value.map((site: any) => {
					return {
						name: site.displayName,
						value: site.id,
					};
				});
			},

			async getSiteDrives(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const siteId = this.getCurrentNodeParameter('siteId') as string;
				const drives = await MSGetSiteDrives(this, siteId);
				return drives.value.map((drive: any) => {
					return {
						name: drive.name,
						value: drive.id,
					};
				});
			}
		}
	};


	// The function below is responsible for actually doing whatever this node
	// is supposed to do. In this case, we're just appending the `myString` property
	// with whatever the user has entered.
	// You can make async calls and use `await`.
	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const operation = this.getNodeParameter('operation', 0) as string;

		if(operation === 'getSites'){
			const output = await MSGetSites(this);

			return this.prepareOutputData(
				output.value.map((site: any) => ({ json: site }))
			);
		}

		if(operation === 'uploadFile'){
			const output: INodeExecutionData[] = [];

			for (let i = 0; i < items.length; i++) {
				const siteId = this.getNodeParameter('siteId', i) as string;
				const libraryId = this.getNodeParameter('libraryId', i) as string;
				const filePath = this.getNodeParameter('filePath', i) as string;
				const fileName = this.getNodeParameter('fileName', i) as string;

				const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;
				this.helpers.assertBinaryData(i, binaryPropertyName);
        		const buffer = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);

				// Figure out folder ID
				this.logger.info('Fetching folder ID...');
				const folder = await makeMicrosoftRequest(this, `drives/${libraryId}/root:/${filePath}:/`);
				if(!folder || !folder.id){
					// Do something
					throw new Error("Could not find folder. Is your path correct?");
				}
				this.logger.info('Got folder ID ' + folder.id);
				
				const res = await makeMicrosoftRequest(this, `sites/${siteId}/drive/items/${folder.id}:/${fileName}:/content`, {
					method: 'PUT',
					body: buffer,
				});

				output.push({ json: res });
			}

			return this.prepareOutputData(output);
		}

		if(operation === 'getItemsInFolder'){
			const output: INodeExecutionData[] = [];
			for (let i = 0; i < items.length; i++) {
				const siteId = this.getNodeParameter('siteId', i) as string;
				const libraryId = this.getNodeParameter('libraryId', i) as string;
				const filePath = this.getNodeParameter('filePath', i) as string;

				// URL needed: /sites/{siteId}/drives/{driveId}/root:/{folder-path}:/children
				const res = await makeMicrosoftRequest(this, `sites/${siteId}/drives/${libraryId}/root:/${filePath}:/children`);
				output.push({ json: res.value });
			}

			return this.prepareOutputData(output);
		}

		if(operation === 'getFile'){
			const output: INodeExecutionData[] = [];

			for(let i = 0; i < items.length; i++){
				const siteId = this.getNodeParameter('siteId', i) as string;
				const libraryId = this.getNodeParameter('libraryId', i) as string;
				const filePath = this.getNodeParameter('filePath', i) as string;

				// Get file metadata
				const resFileDetails = await makeMicrosoftRequest(this, `sites/${siteId}/drives/${libraryId}/root:/${filePath}`);

				// Download the file
				const resFileDownload = await makeMicrosoftRequest(this,  resFileDetails['@microsoft.graph.downloadUrl'], {
					headers: {}, // Don't send the default Content-Type header
					encoding: null, // Don't decode the response body, return a Buffer
				});

				const binaryData = await this.helpers.prepareBinaryData(
					resFileDownload as Buffer,
					resFileDetails.name,
					resFileDetails.file.mimeType,
				);

				output.push({
					json: resFileDetails,
					binary: {
						file: binaryData,
					},
				});
			}

			return this.prepareOutputData(output);
		}

		return this.prepareOutputData([
			{ json: { no: "hi"}},
		]);
	}
}
