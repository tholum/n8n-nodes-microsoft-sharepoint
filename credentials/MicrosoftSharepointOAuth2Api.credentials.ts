import type { ICredentialType, INodeProperties } from 'n8n-workflow';

export class MicrosoftSharepointOAuth2Api implements ICredentialType {
	name = 'microsoftSharepointOAuth2Api';

	extends = ['microsoftOAuth2Api'];

	displayName = 'Microsoft Sharepoint OAuth2 API';

	documentationUrl = 'https://learn.microsoft.com/en-us/graph/auth/auth-concepts';

	properties: INodeProperties[] = [
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			default: 'openid offline_access User.Read.All Sites.Read.All',
		},
	];
}
