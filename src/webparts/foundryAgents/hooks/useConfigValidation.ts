import type { IAgentConfig } from '../models';
import { validateAgent } from '../utils/agentValidation';

/** Localized strings for validation messages. */
export interface IConfigValidationStrings {
  ConfigValidationDisplayNameRequired: string;
  ConfigValidationAgentIdRequired: string;
  ConfigValidationEndpointRequired: string;
  ConfigValidationEndpointInvalid: string;
}

/**
 * Validates that at least one agent is configured and has required fields.
 * Used to show "config required" UI or open property pane.
 * Pass loc strings so firstError is localized.
 */
export function useConfigValidation(
  agents: IAgentConfig[] | undefined,
  locStrings?: IConfigValidationStrings
): {
  isValid: boolean;
  isEmpty: boolean;
  firstError: string | undefined;
} {
  if (!agents || agents.length === 0) {
    return { isValid: false, isEmpty: true, firstError: undefined };
  }
  const replace0 = (s: string, n: number): string => (s && s.includes('{0}') ? s.replace(/\{0\}/g, String(n)) : s);
  for (let i = 0; i < agents.length; i++) {
    const errors = validateAgent(agents[i]);
    const n = i + 1;
    if (errors.displayName) {
      return {
        isValid: false,
        isEmpty: false,
        firstError: locStrings ? replace0(locStrings.ConfigValidationDisplayNameRequired, n) : undefined
      };
    }
    if (errors.agentId) {
      return {
        isValid: false,
        isEmpty: false,
        firstError: locStrings ? replace0(locStrings.ConfigValidationAgentIdRequired, n) : undefined
      };
    }
    if (errors.endpointUrl === 'required') {
      return {
        isValid: false,
        isEmpty: false,
        firstError: locStrings ? replace0(locStrings.ConfigValidationEndpointRequired, n) : undefined
      };
    }
    if (errors.endpointUrl === 'invalid_url') {
      return {
        isValid: false,
        isEmpty: false,
        firstError: locStrings ? replace0(locStrings.ConfigValidationEndpointInvalid, n) : undefined
      };
    }
  }
  return { isValid: true, isEmpty: false, firstError: undefined };
}
