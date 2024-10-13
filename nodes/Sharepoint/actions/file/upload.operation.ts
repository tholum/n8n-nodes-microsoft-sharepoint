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
    const filePath = this.getNodeParameter('filePath', i) as string;
    const fileName = this.getNodeParameter('fileName', i) as string;

    const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;
    this.helpers.assertBinaryData(i, binaryPropertyName);
    const buffer = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);

    // Figure out folder ID
    this.logger.info('Fetching folder ID...');
    const folder = await MSGetItemDetailsByPath(this, libraryId, filePath);
    this.logger.info('Got folder ID ' + folder.id);
    
    const res = await makeMicrosoftRequest(this, `sites/${siteId}/drive/items/${folder.id}:/${fileName}:/content`, {
        method: 'PUT',
        body: buffer,
    });

    return [{
        json: res 
    }];
}