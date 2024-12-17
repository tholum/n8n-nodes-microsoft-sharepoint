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
    const fileLocator = this.getNodeParameter('fileLocator', i) as any;
    const targetFolderLocator = this.getNodeParameter('targetFolderLocator', i) as any;


    // Get details of the driveItem that needs to be moved
    let fileId = fileLocator.value;
    if(fileLocator.mode === "path"){
        const file = await MSGetItemDetailsByPath(this, libraryId, fileLocator.value);
        fileId = file.id;
    }

    // Get details of the new parent
    let parentId = targetFolderLocator.value;
    if(targetFolderLocator.mode === "path"){
        const folder = await MSGetItemDetailsByPath(this, libraryId, targetFolderLocator.value);
        parentId = folder.id;
    }

    const res = await makeMicrosoftRequest(this, `sites/${siteId}/drive/items/${fileId}`, {
        method: 'PATCH',
        body: {
            // Target ID
            "parentReference": {
                "id": parentId,
            },
        }
    });

    return [{
        json: res,
    }];
}