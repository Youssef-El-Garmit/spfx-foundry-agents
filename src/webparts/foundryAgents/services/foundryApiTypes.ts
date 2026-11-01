/**
 * Shared types for Foundry Responses API (response.output / response.data items).
 */

/** One item in response.output or response.data. type or type_ from SDK. */
export interface IResponseOutputItemLike {
  type?: string;
  type_?: string;
  /** function_call: call_id */
  call_id?: string;
  /** function_call: name or function_name */
  name?: string;
  function_name?: string;
  /** function_call: arguments or function_arguments */
  arguments?: string;
  function_arguments?: string;
}
