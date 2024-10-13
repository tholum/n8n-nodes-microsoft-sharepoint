import { IExecuteFunctions, ILoadOptionsFunctions, IRequestOptions } from "n8n-workflow";

export async function makeMicrosoftRequest(thisRef: IExecuteFunctions | ILoadOptionsFunctions, resource: string, options: IRequestOptions = {}): Promise<any> {
	let url = resource;
	
	if(!url.startsWith('https://')) {
		url = "https://graph.microsoft.com/v1.0/" + resource;
	}

	const reqOptions: IRequestOptions = {
		method: 'GET',
		uri: url,
		headers: {
			"Content-Type": "application/json",
		},
		body: {},
		json: true,
		...options
	};

	return await thisRef.helpers.requestOAuth2.call(thisRef, 'microsoftSharepointOAuth2Api', reqOptions);
}