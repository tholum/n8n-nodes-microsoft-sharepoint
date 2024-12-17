import { IExecuteFunctions, INodeExecutionData, NodeApiError } from "n8n-workflow";
import { MSGetItemDetailsByPath } from "../../helpers/misc";
import { makeMicrosoftRequest } from "../../helpers/makeMicrosoftRequest";

/**
 * Creates a new folder in a SharePoint site. By default, it will not create
 * nested folders (similar to default mkdir behaviour). This can be enabled 
 * with the "createIntermediateFolders" option (similar to mkdir -p).
 * 
 * https://learn.microsoft.com/en-us/graph/api/driveitem-post-children?view=graph-rest-1.0&tabs=http  
 * @param this
 * @param i 
 * @returns 
 */
export async function execute(this: IExecuteFunctions, i: number): Promise<INodeExecutionData[]> {
    const siteId = this.getNodeParameter('siteId', i) as string;
    const libraryId = this.getNodeParameter('libraryId', i) as string;
    const options = this.getNodeParameter('options', i, {});
    
    // Strip trailing slashes from folder path provided by user
    const folderPath = (this.getNodeParameter('filePath', i) as string).replace(/\/$/, '');
    let directories = folderPath.split('/');

    // Place to store the ID of the DriveItem under which we should create the
    // next child. By default, we start at the root of the document library.
    let lastParentId = libraryId;

    // Keep track of the last created folder item, so we can return that.
    let lastCreatedItem = null;

    // If the user doesn't want to create intermediate folders, only attempt
    // to create the last folder
    const createIntermedateFolders = (options.createIntermedateFolders as boolean) || false;
    if(createIntermedateFolders === false){
        // Construct path of all folders, except last level
        const parent = directories.slice(0, -1).join('/');
        const last = directories.slice(-1);

        // Set the top most directory as the parent
        const res = await MSGetItemDetailsByPath(this, siteId, libraryId, parent);
        lastParentId = res.id;

        // The only directory that needs to be created now is the last one
        directories = [
            ...last,
        ];
    }

    // Loop over all nested directories that need to be created and create them!
    for(const [idx, folderName] of directories.entries()){
        // When creating nested folders, the first one needs to be created at
        // root level of the document library and requires a different API
        // endpoint
        const createInRoot = idx == 0 && createIntermedateFolders === true;
        lastCreatedItem = await createFolder(this, siteId, lastParentId, folderName, createInRoot);

        // Keep track of the created folder ID so we can use it as parent for
        // the next folder.
        lastParentId = lastCreatedItem.id;
    }

    return [{ json : lastCreatedItem }];
}

/**
 * Helper function that creates a new folder under a given site and parent.
 * When the folder already exists, we fetch and return the data of that folder.
 *  
 * @param thisRef 
 * @param siteId 
 * @param parentId 
 * @param folderName 
 * @returns 
 */
async function createFolder(thisRef: IExecuteFunctions, siteId: string, parentId: string, folderName: string, inRoot: boolean = false) : Promise<any> {
    // We need a slightly different endpoint when creating items at root level.
    const url = inRoot 
                    ? `sites/${siteId}/drive/root/children` 
                    : `sites/${siteId}/drive/items/${parentId}/children`;
    
    try{
        const res = await makeMicrosoftRequest(
            thisRef, 
            url, 
            {
                method: "POST",
                body: {
                    name: folderName,
                    folder: {},
                    // When we try to create a folder that already exists, fail the request
                    '@microsoft.graph.conflictBehavior': 'fail'
                },
            }
        );
    
        return res;
    }catch(error){
        // If we got another error than 409 - CONFLICT, throw it again
        if(error.status !== 409){
            throw new NodeApiError(thisRef.getNode(), error);
        }

        // Folder already exists. Fetch its details so we can return that.
        const res = await makeMicrosoftRequest(
            thisRef, 
            url,
        );

        return res.value.find((el: any) => el.name === folderName);
    }
}
