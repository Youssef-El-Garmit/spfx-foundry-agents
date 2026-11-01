/**
 * Context values for resolving additional information template variables.
 */
export interface IAdditionalInformationContext {
  userDisplayName?: string;
  userMail?: string;
  userPrincipalName?: string;
  currentDateTime?: string;
  currentTimezone?: string;
  currentRegion?: string;
}

/**
 * Replaces {variableName} in the template with values from context.
 * Unknown variables are left as-is (e.g. {unknownVar}).
 */
export function resolveAdditionalInformationTemplate(
  template: string | undefined,
  context: IAdditionalInformationContext
): string {
  if (!template || !template.trim()) return '';
  let out = template;
  const keys: (keyof IAdditionalInformationContext)[] = [
    'userDisplayName',
    'userMail',
    'userPrincipalName',
    'currentDateTime',
    'currentTimezone',
    'currentRegion'
  ];
  for (const key of keys) {
    const value = context[key];
    const placeholder = `{${key}}`;
    out = out.split(placeholder).join(value !== undefined && value !== null ? String(value) : '');
  }
  return out.trim();
}

/**
 * Build default context from browser/Intl (datetime, timezone, region).
 * Call from the client when resolving; merge with user context from Graph/pageContext.
 */
export function getDefaultDateTimeContext(): Pick<
  IAdditionalInformationContext,
  'currentDateTime' | 'currentTimezone' | 'currentRegion'
> {
  const now = new Date();
  try {
    const resolved = Intl.DateTimeFormat().resolvedOptions();
    return {
      currentDateTime: now.toISOString(),
      currentTimezone: resolved.timeZone ?? '',
      currentRegion: resolved.locale ?? ''
    };
  } catch {
    return {
      currentDateTime: now.toISOString(),
      currentTimezone: '',
      currentRegion: ''
    };
  }
}
