import { IExecuteFunctions, INodeExecutionData } from "n8n-workflow";
import { makeMicrosoftRequest } from "../../helpers/makeMicrosoftRequest";
import { MSGetItemDetailsByPath } from "../../helpers/misc";

/**
 * Move a DriveItem to a new folder
 * https://learn.microsoft.com/en-us/graph/api/driveitem-move?view=graph-rest-1.0&tabs=http
 * 
 * @param this
 * @param i 
 * @returns 
 */
export async function execute(this: IExecuteFunctions, i: number): Promise<INodeExecutionData[]> {
    const siteId = this.getNodeParameter('siteId', i) as string;
    const libraryId = this.getNodeParameter('libraryId', i) as string;
    const filePath = this.getNodeParameter('filePath', i) as string;
    const newPath = this.getNodeParameter('targetFolderPath', i) as string;

    // Get details of the driveItem that needs to be moved
    const driveItem = await MSGetItemDetailsByPath(this, libraryId, filePath);
    const targetDriveItem = await MSGetItemDetailsByPath(this, libraryId, newPath);
    
    this.logger.debug(`Moving driveItem ${driveItem.id} to ${targetDriveItem.id}`);

    const res = await makeMicrosoftRequest(this, `sites/${siteId}/drive/items/${driveItem.id}`, {
        method: 'PATCH',
        body: {
            // Target ID
            "parentReference": {
                "id": targetDriveItem.id
            },
        }
    });

    return [{
        json: res,
    }];
    // // Get file metadata
    // const resFileDetails = await makeMicrosoftRequest(this, `sites/${siteId}/drives/${libraryId}/root:/${filePath}`);

    // // Download the file
    // const resFileDownload = await makeMicrosoftRequest(this,  resFileDetails['@microsoft.graph.downloadUrl'], {
    //     headers: {}, // Don't send the default Content-Type header
    //     encoding: null, // Don't decode the response body, return a Buffer
    // });

    // const binaryData = await this.helpers.prepareBinaryData(
    //     resFileDownload as Buffer,
    //     resFileDetails.name,
    //     resFileDetails.file.mimeType,
    // );

    // return [{
    //     json: resFileDetails,
    //     binary: {
    //         file: binaryData,
    //     },
    // }];
}