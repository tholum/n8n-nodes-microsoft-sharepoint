import { IExecuteFunctions, INodeExecutionData, NodeApiError } from "n8n-workflow";
import { MSGetItemDetailsByPath } from "../../helpers/misc";
import { makeMicrosoftRequest } from "../../helpers/makeMicrosoftRequest";

/**
 * The first question is: what happens when you create a folder which already exists?
 *  Google Drive detects this, and adds a number to your folder name, bleh!
 * 
 * https://learn.microsoft.com/en-us/graph/api/driveitem-post-children?view=graph-rest-1.0&tabs=http  
 * @param this
 * @param i 
 * @returns 
 */
export async function execute(this: IExecuteFunctions, i: number): Promise<INodeExecutionData[]> {
    const siteId = this.getNodeParameter('siteId', i) as string;
    const libraryId = this.getNodeParameter('libraryId', i) as string;
    const folderPath = this.getNodeParameter('filePath', i) as string;
    const options = this.getNodeParameter('options', i, {});

    // TODO: strip trailing slash!
    
    let directories = folderPath.split('/');

    // Input: folder1/folder2/folder3 without parameters
    //    -> Fetch ID of folder1/folder2
    //    -> Create folder3 underneath it
    // Input: folder1 without parameters
    //    -> Fetch ID of root folder
    //    -> Upload straight away

    // Place to store the ID of the DriveItem under which we should create the
    // next child. By default, we start at the root of the document library.
    let lastParentId = libraryId;
    let lastCreatedItem = null;

    // If the user doesn't want to create intermediate folders, only attempt
    // to create the last folder
    const createIntermedateFolders = (options.createIntermedateFolders as boolean) || false;
    if(createIntermedateFolders === false){
        const parent = directories.slice(0, -1).join('/');
        const last = directories.slice(-1);

        // Set the top most directory as the parent
        this.logger.info('User does not want recursive creation. Fetching ID for ' + parent);
        const res = await MSGetItemDetailsByPath(this, libraryId, parent);
        lastParentId = res.id;

        this.logger.info('New parent ID: ' + lastParentId);
        // Change the 
        directories = [
            ...last,
        ];
    }

    for(const [idx, folderName] of directories.entries()){
        this.logger.info('Creating folder ' + folderName + ' under ID ' + lastParentId);
        const createInRoot = idx == 0 && createIntermedateFolders === true;
        this.logger.info('idx ' + idx +  ' create inroot ' + createInRoot);
        lastCreatedItem = await createFolder(this, siteId, lastParentId, folderName, createInRoot);
        lastParentId = lastCreatedItem.id;
    }

    return [{ json : lastCreatedItem }];


    // Loop over each folder we need to create in a nested way, starting at the
    // top level.
    let lastParent = null;
    for(const folderName of directories){
        if(lastParent === null){
            this.logger.info('Creating top level folder, first');
            lastParent = await MSGetItemDetailsByPath(this, libraryId, folderName);
            continue;
        }

        // If we get here, create a new folder with the lastParent as parent
        this.logger.info('Creating nested folder ' + folderName + ' under ID' + lastParent.id);
        lastParent = await createFolder(this, siteId, lastParent.id, folderName);
    }

    return [{ json: lastParent }];
 
    // Another shortcut we could consider
    // for(const dirName of directories) {
        
    // }
    // // URL needed: /sites/{siteId}/drives/{driveId}/root:/{folder-path}:/children
    // const res = await makeMicrosoftRequest(this, `sites/${siteId}/drives/${libraryId}/root:/${filePath}:/children`);
    // return res.value.map((item: any) => ({ json: item }));
    return [];
}

/**
 * Helper function that creates a new folder under a given site and parent.
 * When the folder already exists, we return the data of that existing folder.
 *  
 * @param thisRef 
 * @param siteId 
 * @param parentId 
 * @param folderName 
 * @returns 
 */
async function createFolder(thisRef: IExecuteFunctions, siteId: string, parentId: string, folderName: string, inRoot: boolean = false) : Promise<any> {
    const url = inRoot ? `sites/${siteId}/drive/root/children` : `sites/${siteId}/drive/items/${parentId}/children`;
    try{
        thisRef.logger.info('Create folder ' + folderName + ' inRoot? ' + inRoot);
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
    
        thisRef.logger.info('created folder: ' + JSON.stringify(res));
        return res;
    }catch(error){
        // If we got another error than 409 - CONFLICT, throw it again
        if(error.status !== 409){
            throw new NodeApiError(thisRef.getNode(), error);
        }

        thisRef.logger.info('Conflict, folder already exists. Fetching details..');

        const res = await makeMicrosoftRequest(
            thisRef, 
            url,
        );

        thisRef.logger.info('Got children:' + JSON.stringify(res));

        // Return the 
        return res.value.find((el: any) => el.name === folderName);
    }
}