import { IExecuteFunctions, INodeExecutionData } from "n8n-workflow";
import { makeMicrosoftRequest } from "../../helpers/makeMicrosoftRequest";
import { MSGetItemDetailsByPath } from "../../helpers/misc";

/**
 * Uploads new file to Sharepoint in a given Site and Library
 * 
 * https://learn.microsoft.com/en-us/graph/api/driveitem-update
 * 
 * @param this
 * @param i 
 * @returns 
 */
export async function execute(this: IExecuteFunctions, i: number): Promise<INodeExecutionData[]> {
    const siteId = this.getNodeParameter('siteId', i) as string;
    const libraryId = this.getNodeParameter('libraryId', i) as string;
    const fileName = this.getNodeParameter('fileName', i) as string;
    const parentLocator = this.getNodeParameter('parentLocator', i) as any;

    const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;
    this.helpers.assertBinaryData(i, binaryPropertyName);
    const buffer = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);

    // Figure out folder ID
    let parentId = parentLocator.value;
    if(parentLocator.mode === 'path'){
        const folder = await MSGetItemDetailsByPath(this, libraryId, parentLocator.value);
        parentId = folder.id;
    }

    const res = await makeMicrosoftRequest(this, `sites/${siteId}/drive/items/${parentId}:/${fileName}:/content`, {
        method: 'PUT',
        body: buffer,
    });

    return [{
        json: res 
    }];
}