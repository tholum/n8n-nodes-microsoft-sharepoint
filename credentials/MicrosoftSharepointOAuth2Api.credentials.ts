import type { ICredentialType, INodeProperties } from 'n8n-workflow';

export class MicrosoftSharepointOAuth2Api implements ICredentialType {
	name = 'microsoftSharepointOAuth2Api';

	extends = ['microsoftOAuth2Api'];

	displayName = 'Microsoft Sharepoint OAuth2 API';

	documentationUrl = 'microsoft';

	properties: INodeProperties[] = [
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			// TODO: This should not be this broad!
			default: 'openid offline_access User.ReadWrite.All Group.ReadWrite.All Chat.ReadWrite Sites.ReadWrite.All',
		},
	];
}