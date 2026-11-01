/**
 * Replace {0}, {1}, ... in a string with values (same pattern as SPFx/mailroom).
 */
export function formatString(template: string, ...values: string[]): string {
  return template.replace(/\{(\d+)\}/g, (_, i) => values[Number(i)] ?? '');
}
