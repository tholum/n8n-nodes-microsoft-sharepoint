import { IExecuteFunctions, ILoadOptionsFunctions, IN8nHttpFullResponse, IRequestOptions } from "n8n-workflow";

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
		...options,
		resolveWithFullResponse: true,
	};

	const output = await thisRef.helpers.requestOAuth2.call(thisRef, 'microsoftSharepointOAuth2Api', reqOptions) as IN8nHttpFullResponse;

	// Handle throttled responses. Microsoft Graph will return 429 (Too many
	// requests) with a "Retry-After" header that indicates how many seconds
	// we should wait for the next request.
	if(output.statusCode === 429 && output.headers['Retry-After']) {
		const retryAfter = parseInt(output.headers['Retry-After'] as string, 10) || 10;

		thisRef.logger.warn("Sharepoint requests throttled. Waiting " + retryAfter + " seconds to make next request");
		await new Promise(resolve => setTimeout(resolve, retryAfter * 1000));

		return await makeMicrosoftRequest(thisRef, resource, options);
	}

	return output.body;
}