import { IExecuteFunctions, INodeExecutionData } from "n8n-workflow";
import { MSGetSites } from "../../helpers/misc";

/**
 * Returns a list of all Sharepoint sites in the tenant.
 * Uses the search feature with a wildcard to get all sites.
 * 
 * Search for sites: https://learn.microsoft.com/en-us/graph/api/site-search?view=graph-rest-1.0&tabs=http
 * 
 * @param this
 * @param i 
 * @returns 
 */
export async function execute(this: IExecuteFunctions, i: number): Promise<INodeExecutionData[]> {
    const output = await MSGetSites(this);

    return output.value.map((site: any) => ({ json: site }));
}