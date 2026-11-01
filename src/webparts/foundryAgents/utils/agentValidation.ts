import type { IAgentConfig } from '../models';

export type AgentFieldError = 'required' | 'invalid_url' | 'duplicate_name';

export interface IAgentFieldErrors {
  displayName?: AgentFieldError;
  agentId?: AgentFieldError;
  endpointUrl?: AgentFieldError;
}

const ENDPOINT_URL_RE = /^https?:\/\/[^\s]+$/i;

/**
 * Validates a single agent's required fields.
 * Returns an object whose keys are field names and values are error codes.
 * An empty object means the agent is valid.
 */
export function validateAgent(a: IAgentConfig): IAgentFieldErrors {
  const errors: IAgentFieldErrors = {};
  if (!a.displayName?.trim()) errors.displayName = 'required';
  if (!a.agentId?.trim()) errors.agentId = 'required';
  if (!a.endpointUrl?.trim()) {
    errors.endpointUrl = 'required';
  } else if (!ENDPOINT_URL_RE.test(a.endpointUrl.trim())) {
    errors.endpointUrl = 'invalid_url';
  }
  return errors;
}

/**
 * Validates a single agent in the context of the full agent list.
 * Checks all the basic field rules AND ensures the display name is unique
 * among the other agents (i.e. every agent except the one being edited).
 *
 * @param a            The agent being created / edited.
 * @param otherAgents  All agents whose display names must not conflict
 *                     (typically the full list minus the agent being edited).
 */
export function validateAgentInContext(
  a: IAgentConfig,
  otherAgents: IAgentConfig[]
): IAgentFieldErrors {
  const errors = validateAgent(a);

  if (!errors.displayName && a.displayName?.trim()) {
    const trimmed = a.displayName.trim().toLowerCase();
    const isDuplicate = otherAgents.some(
      (other) => other.displayName?.trim().toLowerCase() === trimmed
    );
    if (isDuplicate) {
      errors.displayName = 'duplicate_name';
    }
  }

  return errors;
}

/** Returns true when validateAgent produces no errors. */
export function isAgentValid(a: IAgentConfig): boolean {
  return Object.keys(validateAgent(a)).length === 0;
}
