import { IExecuteFunctions, INodeExecutionData } from "n8n-workflow";
import { makeMicrosoftRequest } from "../../helpers/makeMicrosoftRequest";

/**
 * Gets metadata of a given file and downloads the binary
 * 
 * Get item: https://learn.microsoft.com/en-us/graph/api/driveitem-get?view=graph-rest-1.0&tabs=http
 * Download file: https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0&tabs=http
 * 
 * @param this
 * @param i 
 * @returns 
 */
export async function execute(this: IExecuteFunctions, i: number): Promise<INodeExecutionData[]> {
    const siteId = this.getNodeParameter('siteId', i) as string;
    const libraryId = this.getNodeParameter('libraryId', i) as string;
    const fileLocator = this.getNodeParameter('fileLocator', i) as any;

    let fileDetails = null;

    if(fileLocator.mode === "path"){
        fileDetails = await makeMicrosoftRequest(this, `sites/${siteId}/drives/${libraryId}/root:/${fileLocator.path}`);
    }

    if(fileLocator.mode === "id"){
        fileDetails = await makeMicrosoftRequest(this, `sites/${siteId}/drive/items/${fileLocator.value}`);
    }

    // Download the file
    const resFileDownload = await makeMicrosoftRequest(this,  fileDetails['@microsoft.graph.downloadUrl'], {
        headers: {}, // Don't send the default Content-Type header
        encoding: null, // Don't decode the response body, return a Buffer
    });

    const binaryData = await this.helpers.prepareBinaryData(
        resFileDownload as Buffer,
        fileDetails.name,
        fileDetails.file.mimeType,
    );

    return [{
        json: fileDetails,
        binary: {
            file: binaryData,
        },
    }];
}