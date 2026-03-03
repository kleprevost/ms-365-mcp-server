import logger from './logger.js';

export interface FilterResult {
  data: unknown;
  filteredCount: number;
  wasBlocked: boolean;
}

/**
 * Reads the list of blocked sensitivity label names from the environment variable
 * MS365_MCP_BLOCKED_SENSITIVITY_LABELS (comma-separated).
 */
export function getBlockedSensitivityLabels(): string[] {
  const envValue = process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS;
  if (!envValue) return [];
  return envValue
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

/**
 * Returns true if a sensitivityLabel object's displayName matches any blocked label
 * (case-insensitive).
 */
function isBlockedLabel(sensitivityLabel: unknown, blockedLabels: string[]): boolean {
  if (!sensitivityLabel || typeof sensitivityLabel !== 'object') return false;
  const label = sensitivityLabel as Record<string, unknown>;
  const displayName = label.displayName;
  if (typeof displayName !== 'string') return false;
  return blockedLabels.some((blocked) => displayName.toLowerCase() === blocked.toLowerCase());
}

/**
 * Filters hits from a /search/query response.
 * Structure: { value: [{ hitsContainers: [{ hits: [{ resource: { sensitivityLabel } }] }] }] }
 */
function filterSearchQueryResults(
  data: Record<string, unknown>,
  blockedLabels: string[]
): { data: Record<string, unknown>; filteredCount: number } {
  let filteredCount = 0;

  if (!Array.isArray(data.value)) return { data, filteredCount };

  const newValue = data.value.map((searchResult: unknown) => {
    if (!searchResult || typeof searchResult !== 'object') return searchResult;
    const result = searchResult as Record<string, unknown>;

    if (!Array.isArray(result.hitsContainers)) return result;

    const newContainers = result.hitsContainers.map((container: unknown) => {
      if (!container || typeof container !== 'object') return container;
      const cont = container as Record<string, unknown>;

      if (!Array.isArray(cont.hits)) return cont;

      const originalCount = cont.hits.length;
      const filteredHits = cont.hits.filter((hit: unknown) => {
        if (!hit || typeof hit !== 'object') return true;
        const hitObj = hit as Record<string, unknown>;
        const resource = hitObj.resource as Record<string, unknown> | undefined;
        if (!resource) return true;

        if (isBlockedLabel(resource.sensitivityLabel, blockedLabels)) {
          const label = (resource.sensitivityLabel as Record<string, unknown>).displayName;
          logger.warn(
            `Sensitivity filter: blocked search hit "${resource.name || resource.id}" with label "${label}"`
          );
          return false;
        }
        return true;
      });

      filteredCount += originalCount - filteredHits.length;
      return { ...cont, hits: filteredHits, total: filteredHits.length };
    });

    return { ...result, hitsContainers: newContainers };
  });

  return { data: { ...data, value: newValue }, filteredCount };
}

/**
 * Filters items from a list response.
 * Structure: { value: [{ sensitivityLabel }] }
 */
function filterListResults(
  data: Record<string, unknown>,
  blockedLabels: string[]
): { data: Record<string, unknown>; filteredCount: number } {
  let filteredCount = 0;

  if (!Array.isArray(data.value)) return { data, filteredCount };

  const originalCount = data.value.length;
  const filteredValue = data.value.filter((item: unknown) => {
    if (!item || typeof item !== 'object') return true;
    const itemObj = item as Record<string, unknown>;

    if (isBlockedLabel(itemObj.sensitivityLabel, blockedLabels)) {
      const label = (itemObj.sensitivityLabel as Record<string, unknown>).displayName;
      logger.warn(
        `Sensitivity filter: blocked list item "${itemObj.name || itemObj.id}" with label "${label}"`
      );
      return false;
    }
    return true;
  });

  filteredCount = originalCount - filteredValue.length;
  return { data: { ...data, value: filteredValue }, filteredCount };
}

/**
 * Applies sensitivity label filtering to any Microsoft Graph API response.
 *
 * Handles three response shapes:
 *   1. /search/query  — value[].hitsContainers[].hits[].resource.sensitivityLabel
 *   2. List responses — value[].sensitivityLabel
 *   3. Single-item   — top-level sensitivityLabel
 *
 * Items without a sensitivityLabel field pass through (fail-open).
 */
export function filterSensitivityLabels(data: unknown, blockedLabels: string[]): FilterResult {
  if (!data || typeof data !== 'object' || blockedLabels.length === 0) {
    return { data, filteredCount: 0, wasBlocked: false };
  }

  const obj = data as Record<string, unknown>;

  // Single-item response: top-level sensitivityLabel is blocked
  if (isBlockedLabel(obj.sensitivityLabel, blockedLabels)) {
    const label = (obj.sensitivityLabel as Record<string, unknown>).displayName;
    logger.warn(
      `Sensitivity filter: blocked single item "${obj.name || obj.id}" with label "${label}"`
    );
    return {
      data: {
        error:
          'This item has a restricted sensitivity label and cannot be accessed through this interface.',
      },
      filteredCount: 1,
      wasBlocked: true,
    };
  }

  // /search/query response: value[].hitsContainers[].hits[]
  if (
    Array.isArray(obj.value) &&
    obj.value.length > 0 &&
    typeof obj.value[0] === 'object' &&
    obj.value[0] !== null &&
    'hitsContainers' in (obj.value[0] as object)
  ) {
    const { data: filtered, filteredCount } = filterSearchQueryResults(obj, blockedLabels);
    return { data: filtered, filteredCount, wasBlocked: false };
  }

  // List response: value[]
  if (Array.isArray(obj.value)) {
    const { data: filtered, filteredCount } = filterListResults(obj, blockedLabels);
    return { data: filtered, filteredCount, wasBlocked: false };
  }

  return { data, filteredCount: 0, wasBlocked: false };
}
