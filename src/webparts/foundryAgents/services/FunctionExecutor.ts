/**
 * SPFx — generic function_call executor. Contract never changes.
 * Arguments JSON: { request: { path, method, query?, body?, headers? } }
 * audience: spo_* → https://${tenant}.sharepoint.com, else → https://graph.microsoft.com
 */

export type GetTokenFn = (resource: string) => Promise<string>;

/** request shape from data.item.arguments (agent contract) */
export interface IFunctionRequest {
  path: string;
  method: string;
  query?: Record<string, string>;
  body?: unknown;
  headers?: Record<string, string>;
}

function parseArguments(argsJson: string): { request: IFunctionRequest } | undefined {
  try {
    const data = argsJson && argsJson.trim() ? JSON.parse(argsJson) : {};
    const request = data.request;
    if (!request || typeof request.path !== 'string' || typeof request.method !== 'string') {
      return undefined;
    }
    return { request };
  } catch {
    return undefined;
  }
}

/**
 * Execute one function_call: request-based contract only.
 * tenantHost: for spo_* calls, the current SharePoint tenant origin (e.g. https://contoso.sharepoint.com), derived from the page context. No custom tenant.
 */
export async function executeFromFunctionCall(
  functionName: string,
  argumentsJson: string,
  getUserToken: GetTokenFn,
  tenantHost: string
): Promise<unknown> {
  const parsed = parseArguments(argumentsJson);
  if (parsed === undefined) {
    return { error: 'Invalid arguments: expected { request: { path, method, query?, body?, headers? } }' };
  }

  const { request } = parsed;
  const tenant = (tenantHost ?? '').trim();
  const audience = functionName.startsWith('spo_')
    ? (tenant.startsWith('http') ? tenant : tenant.includes('.') ? `https://${tenant}` : `https://${tenant}.sharepoint.com`)
    : 'https://graph.microsoft.com';

  let token: string;
  try {
    token = await getUserToken(audience);
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return { error: `Token failed: ${message}` };
  }

  const qs = request.query
    ? '?' + new URLSearchParams(request.query).toString()
    : '';

  const url = `${audience}/v1.0${request.path.startsWith('/') ? request.path : '/' + request.path}${qs}`;

  let res: Response;
  try {
    res = await fetch(url, {
      method: request.method,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...(request.headers ?? {})
      },
      body: request.body !== undefined && request.body !== null ? JSON.stringify(request.body) : undefined
    });
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return { error: message };
  }

  let result: unknown;
  try {
    const text = await res.text();
    result = text ? (JSON.parse(text) as unknown) : {};
  } catch {
    result = { error: `HTTP ${res.status}`, code: res.status };
  }

  if (!res.ok) {
    const errMsg = (result as { error?: { message?: string }; message?: string })?.error?.message
      ?? (result as { message?: string })?.message
      ?? `HTTP ${res.status}`;
    return { error: errMsg, code: res.status };
  }

  return result;
}
