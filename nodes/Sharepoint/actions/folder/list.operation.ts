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

    // URL needed: /sites/{siteId}/drives/{driveId}/root:/{folder-path}:/children
    const res = await makeMicrosoftRequest(this, `sites/${siteId}/drives/${libraryId}/root:/${filePath}:/children`);
    return res.value.map((item: any) => ({ json: item }));
}