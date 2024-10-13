import { IExecuteFunctions, ILoadOptionsFunctions, NodeApiError } from "n8n-workflow";
import { makeMicrosoftRequest } from "./makeMicrosoftRequest";

export async function MSGetSites(thisRef: IExecuteFunctions | ILoadOptionsFunctions): Promise<any> {
	return await makeMicrosoftRequest(thisRef, 'sites', {
		qs: {
			search: '*',
		}
	});
}

export async function MSGetSiteDrives(thisRef: IExecuteFunctions | ILoadOptionsFunctions, siteId: string): Promise<any> {
	return await makeMicrosoftRequest(thisRef, `sites/${siteId}/drives`);
}

export async function MSGetItemDetailsByPath(thisRef: IExecuteFunctions, libraryId: string, filePath: string): Promise<any> {
    const folder = await makeMicrosoftRequest(thisRef, `drives/${libraryId}/root:/${filePath}:/`);
    if(!folder || !folder.id){
        // Do something
        throw new NodeApiError(thisRef.getNode(), folder, {
            message: "Could not find folder. Is your path correct?"
        });
    }

    return folder;
}