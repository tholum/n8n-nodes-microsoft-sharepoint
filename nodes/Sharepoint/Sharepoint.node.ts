import {
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodePropertyOptions,
	INodeType,
	INodeTypeDescription,
	IRequestOptions,
	NodeOperationError,
} from 'n8n-workflow';


async function makeMicrosoftRequest(thisRef: IExecuteFunctions | ILoadOptionsFunctions, resource: string, options: IRequestOptions = {}): Promise<any> {
	const reqOptions: IRequestOptions = {
		method: 'GET',
		uri: "https://graph.microsoft.com/v1.0/" + resource,
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
		displayName: 'SharePoint',
		name: 'sharePoint',
		icon: 'file:Sharepoint.svg',
		group: ['transform'],
		version: 1,
		description: 'Interact with SharePoint',
		defaults: {
			name: 'SharePoint',
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
			// --------------- Main Actions ------------------
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				options: [
					{
						name: 'Get File',
						value: 'getFile',
					},
					{
						name: 'Upload File',
						value: 'uploadFile',
					},
					{
						name: 'Get Sites',
						value: 'getSites',
					},
				],
				default: 'getFile',
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
						operation: ['getFile', 'uploadFile'],
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
						operation: ['getFile', 'uploadFile'],
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
						operation: ['getFile', 'uploadFile'],
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
						operation: ['getFile', 'uploadFile'],
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
				this.logger.warn("MS upload: " + JSON.stringify(res));

				return this.prepareOutputData([
					{ json: { no: "hi2"}},
				]);
			}
		}


		return this.prepareOutputData([
			{ json: { no: "hi"}},
		]);

			// const reqOptions: IRequestOptions = {
			// 	method: 'GET',
			// 	uri: "https://graph.microsoft.com/v1.0/sites?search=*",
			// 	headers: {
			// 		"Content-Type": "application/json",
			// 	},
			// 	body: {},
			// 	json: true,
			// };
			// const response = await this.helpers.requestOAuth2.call(this, 'microsoftSharepointOAuth2Api', reqOptions);
			// this.logger.warn('Sharepoint response:' + JSON.stringify(response));

		

		

		let item: INodeExecutionData;
		let myString: string;

		// Iterates over all input items and add the key "myString" with the
		// value the parameter "myString" resolves to.
		// (This could be a different value for each item in case it contains an expression)
		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				myString = this.getNodeParameter('myString', itemIndex, '') as string;
				item = items[itemIndex];

				item.json['myString'] = myString;
			} catch (error) {
				// This node should never fail but we want to showcase how
				// to handle errors.
				if (this.continueOnFail()) {
					items.push({ json: this.getInputData(itemIndex)[0].json, error, pairedItem: itemIndex });
				} else {
					// Adding `itemIndex` allows other workflows to handle this error
					if (error.context) {
						// If the error thrown already contains the context property,
						// only append the itemIndex
						error.context.itemIndex = itemIndex;
						throw error;
					}
					throw new NodeOperationError(this.getNode(), error, {
						itemIndex,
					});
				}
			}
		}

		return this.prepareOutputData(items);
	}
}
