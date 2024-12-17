import type { ICredentialType, INodeProperties } from 'n8n-workflow';

export class MicrosoftSharepointOAuth2Api implements ICredentialType {
	name = 'microsoftSharepointOAuth2Api';

	extends = ['microsoftOAuth2Api'];

	displayName = 'Microsoft SharePoint OAuth2 API';

	documentationUrl = 'https://docs.n8n.io/integrations/builtin/credentials/microsoft/';

	properties: INodeProperties[] = [
		{
			displayName: 'Grant Type',
			name: 'grantType',
			type: 'hidden',
			default: 'authorizationCode',
		},
		{
			displayName: 'Authorization URL',
			name: 'authUrl',
			type: 'hidden',
			default: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
		},
		{
			displayName: 'Access Token URL',
			name: 'accessTokenUrl',
			type: 'hidden',
			default: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
		},
		{
			displayName: 'Auth URI Query Parameters',
			name: 'authQueryParameters',
			type: 'hidden',
			default: 'response_type=code&prompt=consent',
		},
		{
			displayName: 'Authentication',
			name: 'authentication',
			type: 'hidden',
			default: 'header',
		},
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			default: 'openid offline_access https://graph.microsoft.com/.default',
		}
	];
}
