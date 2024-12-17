import { IExecuteFunctions, ILoadOptionsFunctions } from "n8n-workflow";
import { makeMicrosoftRequest } from "./makeMicrosoftRequest";

export async function MSGetItemDetailsByPath(
	thisRef: IExecuteFunctions | ILoadOptionsFunctions,
	siteId: string,
	driveId: string,
	path: string
): Promise<any> {
	return await makeMicrosoftRequest(
		thisRef,
		`sites/${siteId}/drives/${driveId}/root:/${path}`
	);
}

export async function MSGetSites(thisRef: IExecuteFunctions | ILoadOptionsFunctions): Promise<any> {
	let allSites: any[] = [];
	let nextLink: string | undefined;

	do {
		const response = nextLink 
			? await makeMicrosoftRequest(thisRef, nextLink)
			: await makeMicrosoftRequest(thisRef, 'sites', {
					qs: {
						search: '*',
						$top: 100,
						$select: 'id,name,displayName,webUrl,siteCollection'
					}
				});

		if (response.value) {
			// Add more descriptive names including the site URL
			const enhancedSites = response.value.map((site: any) => ({
				...site,
				displayName: `${site.displayName} (${site.webUrl})`,
			}));
			allSites = allSites.concat(enhancedSites);
		}

		nextLink = response['@odata.nextLink'];
		if (nextLink) {
			// Remove the base URL from nextLink as makeMicrosoftRequest will add it
			nextLink = nextLink.replace('https://graph.microsoft.com/v1.0/', '');
		}
	} while (nextLink);

	// Log the sites for debugging
	thisRef.logger.info('Available SharePoint sites:', { sites: allSites });

	return { value: allSites };
}

export async function MSGetSiteDrives(thisRef: IExecuteFunctions | ILoadOptionsFunctions, siteId: string): Promise<any> {
	let allDrives: any[] = [];
	let nextLink: string | undefined;

	do {
		const response = nextLink
			? await makeMicrosoftRequest(thisRef, nextLink)
			: await makeMicrosoftRequest(thisRef, `sites/${siteId}/drives`, {
					qs: {
						$top: 100,
						$select: 'id,name,driveType,webUrl'
					}
				});

		if (response.value) {
			// Add more descriptive names including the drive URL
			const enhancedDrives = response.value.map((drive: any) => ({
				...drive,
				name: `${drive.name} (${drive.webUrl})`,
			}));
			allDrives = allDrives.concat(enhancedDrives);
		}

		nextLink = response['@odata.nextLink'];
		if (nextLink) {
			nextLink = nextLink.replace('https://graph.microsoft.com/v1.0/', '');
		}
	} while (nextLink);

	// Log the drives for debugging
	thisRef.logger.info('Available drives for site:', { drives: allDrives });

	return { value: allDrives };
}
