import { IExecuteFunctions, INodeExecutionData } from "n8n-workflow";
import { makeMicrosoftRequest } from "../../helpers/makeMicrosoftRequest";

/**
 * Returns a list of children of a folder in Sharepoint.
 * 
 * List children of a driveItem:
 * https://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0&tabs=http
 *  
 * @param this
 * @param i 
 * @returns 
 */
export async function execute(this: IExecuteFunctions, i: number): Promise<INodeExecutionData[]> {
    const siteId = this.getNodeParameter('siteId', i) as string;
    const libraryId = this.getNodeParameter('libraryId', i) as string;
    const filePath = this.getNodeParameter('filePath', i) as string;
    const options = this.getNodeParameter('optionsGetItemsInFolder', i, {}) as any;
    const returnAll = options.returnAll || false;

    const output = [];

    // URL needed: /sites/{siteId}/drives/{driveId}/root:/{folder-path}:/children
    let requestUrl: string|null = `sites/${siteId}/drives/${libraryId}/root:/${filePath}:/children`;

    // Keep looping while we have a requestUrl. This will be changed to support
    // paging and set to null when we reach the end.
    while(requestUrl !== null){
        const res = await makeMicrosoftRequest(this, requestUrl);
        output.push(...res.value);

        if(returnAll === false){
            break;
        }

        requestUrl = res['@odata.nextLink'] || null;
    }

    return output.map((item: any) => ({ json: item }));
}
