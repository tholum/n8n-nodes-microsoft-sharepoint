import {
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodePropertyOptions,
	INodeType,
	INodeTypeDescription,
	NodeApiError,
} from 'n8n-workflow';
import * as file from './actions/file/File.resource';
import * as site from './actions/site/Site.resource';
import * as folder from './actions/folder/Folder.resource';

import { MSGetSiteDrives, MSGetSites } from './helpers/misc';

export class MicrosoftSharepoint implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft SharePoint',
		name: 'microsoftSharepoint',
		icon: 'file:MicrosoftSharepoint.svg',
		group: ['transform'],
		version: 1,
		description: 'Interact with Microsoft SharePoint',
		defaults: {
			name: 'Microsoft SharePoint',
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
						name: 'Folder',
						value: 'folder',
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
						name: 'Get File',
						action: 'Get file',
						value: 'getFile',
					},
					{
						name: 'Upload File',
						action: 'Upload file',
						value: 'uploadFile',
					},
					{
						name: 'Move File',
						action: 'Move file',
						value: 'moveFile',
					},
				],
				default: 'getFile',
				required: true,
				noDataExpression: true,
			},

			// --------------- Folder Actions ------------------
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				displayOptions: {
					show: {
						resource: ['folder'],
					},
				},
				options: [
					{
						name: 'Get Items in a Folder',
						action: 'Get items in folder',
						value: 'getItemsInFolder',
					},
					{
						name: 'Create Folder',
						action: 'Create folder',
						value: 'createFolder',
					},
				],
				default: 'getItemsInFolder',
				noDataExpression: true,
				required: true,
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
						action: 'Get sites',
						value: 'getSites',
					},
				],
				default: 'getSites',
				noDataExpression: true,
				required: true,
			},

			// ---------------- Parameters -------------------
			{
				displayName: 'Site Name or ID',
				name: 'siteId',
				type: 'options',
				description: 'Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>',
				typeOptions: {
					loadOptionsMethod: 'getSites',
				},
				default: '',
				required: true,
				displayOptions: {
					show: {
						operation: ['getFile', 'uploadFile', 'moveFile', 'getItemsInFolder', 'createFolder'],
					},
				},
			},
			{
				displayName: 'Document Library Name or ID',
				name: 'libraryId',
				type: 'options',
				description: 'Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>',
				typeOptions: {
					loadOptionsMethod: 'getSiteDrives',
					loadOptionsDependsOn: ['siteId'],
				},
				default: '',
				required: true,
				displayOptions: {
					show: {
						operation: ['getFile', 'uploadFile', 'moveFile', 'getItemsInFolder', 'createFolder'],
					},
				},
			},
			{
				displayName: 'File',
				name: 'fileLocator',
				type: 'resourceLocator',
				required: true,
				default: { mode: 'path' },
				modes: [
					{
						displayName: 'Path',
						name: 'path',
						type: 'string',
						placeholder: 'e.g. /sites/test/Documents/test.txt',
					},
					{
						displayName: 'ID',
						name: 'id',
						type: 'string',
						placeholder: 'e.g. 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ',
					}
				],
				displayOptions: {
					show: {
						operation: ['getFile', 'moveFile'],
					}
				}
			},
			{
				// Used to select the parent folder while uploading a new file
				displayName: 'Parent Folder',
				name: 'parentLocator',
				type: 'resourceLocator',
				required: true,
				default: { mode: 'path' },
				modes: [
					{
						displayName: 'Path',
						name: 'path',
						type: 'string',
						placeholder: 'e.g. /sites/test/Documents/',
					},
					{
						displayName: 'ID',
						name: 'id',
						type: 'string',
						placeholder: 'e.g. 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ',
					}
				],
				displayOptions: {
					show: {
						operation: ['uploadFile'],
					}
				}
			},
			{
				displayName: 'File Path',
				name: 'filePath',
				type: 'string',
				default: '',
				required: true,
				displayOptions: {
					show: {
						operation: ['getItemsInFolder', 'createFolder'],
					},
				},
			},
			{
				displayName: 'Target Folder',
				name: 'targetFolderLocator',
				type: 'resourceLocator',
				default: { mode: 'path' },
				required: true,
				modes: [
					{
						displayName: 'Path',
						name: 'path',
						type: 'string',
						placeholder: 'e.g. /sites/test/Documents/',
					},
					{
						displayName: 'ID',
						name: 'id',
						type: 'string',
						placeholder: 'e.g. 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ',
					}
				],
				displayOptions: {
					show: {
						operation: ['moveFile'],
					},
				}
			},
			{
				displayName: 'File Name',
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
			{
				displayName: 'Options',
				name: 'options',
				type: 'collection',
				displayOptions: {
					show: {
						operation: ['getFile'],
						resource: ['file'],
					},
				},
				default: {},
				placeholder: 'Add Option',
				options: [
					{
						displayName: 'Include Extra Fields',
						name: 'includeExtraFields',
						type: 'boolean',
						default: false,
						description: 'Whether to include input data in the output',
					},
				],
			},
			{
				displayName: 'Options',
				name: 'options',
				type: 'collection',
				displayOptions: {
					show: {
						operation: ['createFolder'],
						resource: ['folder'],
					},
				},
				default: {},
				placeholder: 'Add Option',
				options: [
					{
						displayName: 'Create Intermediate Folders',
						name: 'createIntermedateFolders',
						type: 'boolean',
						default: false,
						description: 'Whether to create intermediate directories (similar to mkdir -p)',
					},
				],
			},
			{
				displayName: 'Options',
				name: 'optionsGetItemsInFolder',
				type: 'collection',
				displayOptions: {
					show: {
						operation: ['getItemsInFolder'],
						resource: ['folder'],
					},
				},
				default: {},
				placeholder: 'Add Option',
				options: [
					{
						displayName: 'Return All',
						name: 'returnAll',
						type: 'boolean',
						default: false,
						description: 'Whether to return all results or only up to a given limit',
					},
					{
						displayName: 'Recursive',
						name: 'recursive',
						type: 'boolean',
						default: false,
						description: 'Whether to list items in subfolders recursively',
					},
				],
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

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		const operationMapping: any = {
			'uploadFile': file.upload,
			'getFile': file.get,
			'moveFile': file.move,
			'getSites': site.getAll,
			'getItemsInFolder': folder.list,
			'createFolder': folder.create,
		};
		
		// Execute the operation!
		for(let i = 0; i < items.length; i++){
			const operation = this.getNodeParameter('operation', i) as string;

			if(!operationMapping[operation]){
				throw new NodeApiError(this.getNode(), {}, {
					message: 'Unsupported operation called',
				});
			}

			returnData.push(...await operationMapping[operation].execute.call(this, i));
		}

		return [returnData];
	}
}
