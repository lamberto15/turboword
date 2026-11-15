// Non-fatal problems are reported as warnings through an optional callback instead of being
// thrown or silently ignored.

export interface Warning {
  /** Stable machine-readable code, e.g. `'DANGLING_OVERRIDE'`. */
  readonly code: string;
  readonly message: string;
  /** Part the warning is about, when there is one (no leading slash). */
  readonly part?: string;
}

export type WarningHandler = (warning: Warning) => void;

/** Builds a warning without `part: undefined` (exactOptionalPropertyTypes-friendly). */
export function warning(code: string, message: string, part?: string): Warning {
  return part === undefined ? { code, message } : { code, message, part };
}
