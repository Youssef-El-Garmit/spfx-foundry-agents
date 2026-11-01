/**
 * TokenCredential implementation that uses SPFx AAD token provider.
 * Azure AI Foundry expects token audience https://ai.azure.com (see 401 error message).
 */
import type { AccessToken, GetTokenOptions, TokenCredential } from '@azure/core-auth';

/** Resource/audience for SPFx getToken - must be https://ai.azure.com for Azure AI Foundry */
export const AZURE_AI_SCOPE = 'https://ai.azure.com';

export type GetTokenFn = (resource: string) => Promise<string>;

/**
 * Wraps SPFx getToken (resource => token) as Azure SDK TokenCredential.
 */
export class SpFxTokenCredential implements TokenCredential {
  constructor(private getTokenFn: GetTokenFn) {}

  async getToken(
    scopes: string | string[],
    _options?: GetTokenOptions
  // eslint-disable-next-line @rushstack/no-new-null
  ): Promise<AccessToken | null> {
    const scope = Array.isArray(scopes) ? scopes[0] : scopes;
    const resource = scope && scope.indexOf('/.default') !== -1 ? scope.replace('/.default', '') : scope || 'https://ai.azure.com';
    const token = await this.getTokenFn(resource);
    if (!token) return null; // eslint-disable-line @rushstack/no-new-null -- TokenCredential requires null
    return {
      token,
      expiresOnTimestamp: Date.now() + 3600 * 1000
    };
  }
}
