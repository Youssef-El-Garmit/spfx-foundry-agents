import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx, type SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-groups/web';
import type { IAudienceGroup } from '../models/IAgentConfig';

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Extract the AAD object ID (GUID) from a SharePoint claim-based LoginName.
 * SharePoint stores nested M365/AAD groups with a login like:
 *   c:0o.c|federateddirectoryclaimprovider|{guid}
 * Returns null if no GUID is found (safe fallback).
 */
// eslint-disable-next-line @rushstack/no-new-null
function _extractAadObjectId(loginName: string): string | null {
  const match = loginName?.match(/\|([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$/i);
  return match?.[1] ?? null;
}

// ─────────────────────────────────────────────────────────────────────────────
// Types used by the picker (search results)
// ─────────────────────────────────────────────────────────────────────────────

export interface IAudienceSearchResult {
  id: string;
  displayName: string;
  type: 'graph' | 'sharepoint' | 'user';
  mail?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// AudienceService
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Service for audience-related operations:
 *  - Searching Entra (Graph) groups and SharePoint site groups
 *  - Checking whether the current user belongs to a given set of audience groups
 */
export class AudienceService {
  private readonly _context: WebPartContext;
  private readonly _sp: SPFI;
  private _graphClient: MSGraphClientV3 | undefined;

  private readonly _graphMemberCache = new Map<string, Promise<boolean>>();
  private readonly _spMemberCache = new Map<string, Promise<boolean>>();

  constructor(context: WebPartContext) {
    this._context = context;
    this._sp = spfi().using(SPFx(context));
  }

  // ── Graph client ────────────────────────────────────────────────────────

  private async _getGraph(): Promise<MSGraphClientV3> {
    if (this._graphClient) return this._graphClient;
    const factory = (this._context as unknown as {
      msGraphClientFactory?: { getClient: (v: '3') => Promise<MSGraphClientV3> };
    }).msGraphClientFactory;
    if (!factory) throw new Error('MSGraphClientFactory not available in context');
    this._graphClient = await factory.getClient('3');
    return this._graphClient;
  }

  // ── Search ───────────────────────────────────────────────────────────────

  /**
   * Search Entra groups whose displayName or mail contains `query`.
   * Uses $search with ConsistencyLevel:eventual (same pattern as OnlineDirectoryService).
   */
  public async searchGraphGroups(query: string): Promise<IAudienceSearchResult[]> {
    if (!query.trim()) return [];
    try {
      const graph = await this._getGraph();
      const escaped = query.replace(/"/g, '\\"');
      const result = await graph
        .api('/groups')
        .header('ConsistencyLevel', 'eventual')
        .select('id,displayName,mail')
        .search(`"displayName:${escaped}" OR "mail:${escaped}"`)
        .top(15)
        .get() as { value?: Array<{ id: string; displayName: string; mail?: string }> };
      return (result?.value ?? []).map((g) => ({
        id: g.id,
        displayName: g.displayName,
        type: 'graph' as const,
        mail: g.mail
      }));
    } catch {
      return [];
    }
  }

  /**
   * Search SharePoint site groups whose title contains `query` using PnP.
   */
  public async searchSharePointGroups(query: string): Promise<IAudienceSearchResult[]> {
    if (!query.trim()) return [];
    try {
      const escaped = query.replace(/'/g, "''");
      const groups = await this._sp.web.siteGroups
        .filter(`substringof('${escaped}', Title)`)
        .select('Id', 'Title')
        .top(15)();
      return groups.map((g) => ({
        id: String(g.Id),
        displayName: g.Title,
        type: 'sharepoint' as const
      }));
    } catch {
      return [];
    }
  }

  /**
   * Search individual Entra users by displayName, mail, or UPN.
   */
  public async searchGraphUsers(query: string): Promise<IAudienceSearchResult[]> {
    if (!query.trim()) return [];
    try {
      const graph = await this._getGraph();
      const escaped = query.replace(/"/g, '\\"');
      const result = await graph
        .api('/users')
        .header('ConsistencyLevel', 'eventual')
        .select('id,displayName,mail,userPrincipalName')
        .search(`"displayName:${escaped}" OR "mail:${escaped}" OR "userPrincipalName:${escaped}"`)
        .top(10)
        .get() as { value?: Array<{ id: string; displayName: string; mail?: string; userPrincipalName?: string }> };
      return (result?.value ?? []).map((u) => ({
        id: u.id,
        displayName: u.displayName,
        type: 'user' as const,
        mail: u.mail ?? u.userPrincipalName
      }));
    } catch {
      return [];
    }
  }

  /**
   * Merge results from Graph groups, SharePoint groups, and individual users.
   */
  public async searchAll(query: string): Promise<IAudienceSearchResult[]> {
    const [graphGroups, spGroups, users] = await Promise.all([
      this.searchGraphGroups(query),
      this.searchSharePointGroups(query),
      this.searchGraphUsers(query)
    ]);
    const seen = new Set<string>();
    return [...graphGroups, ...users, ...spGroups].filter((r) => {
      const key = `${r.type}:${r.id}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  // ── Membership check ─────────────────────────────────────────────────────

  /**
   * Check whether the current user is a member of the given Entra group (by objectId).
   * Uses POST /me/checkMemberGroups which handles transitive membership and is the
   * most efficient single call for this purpose.
   */
  private async _isGraphGroupMember(groupId: string): Promise<boolean> {
    return this._cacheOrRun(this._graphMemberCache, `graph:${groupId}`, async () => {
      try {
        const graph = await this._getGraph();
        const result = await graph
          .api('/me/checkMemberGroups')
          .post({ groupIds: [groupId] }) as { value?: string[] };
        return (result?.value ?? []).includes(groupId);
      } catch {
        return false;
      }
    });
  }

  /**
   * Check whether the current user is a member of the given SharePoint site group.
   *
   * Two-pass strategy that handles both custom groups and default groups
   * (Owners / Members / Visitors) in modern Team sites:
   *
   *  Pass 1 – Direct match: load the SP group's members and compare by
   *            LoginName or Email (same User.equal logic). Covers custom groups
   *            and classic sites where users are added directly.
   *
   *  Pass 2 – Transitive match: default groups contain the M365 group itself
   *            as their only principal (not individual users). For those entries
   *            we extract the AAD objectId from the claim LoginName and reuse
   *            _isGraphGroupMember which calls Graph POST /me/checkMemberGroups
   *            (already works for Entra groups).
   */
  private _isSpGroupMember(groupId: string): Promise<boolean> {
    return this._cacheOrRun(this._spMemberCache, `sp:${groupId}`, async () => {
      try {
        const { loginName, email } = this._context.pageContext.user;

        const members = await this._sp.web.siteGroups
          .getById(parseInt(groupId, 10))
          .users
          .select('LoginName', 'Email', 'PrincipalType')();

        // Pass 1: direct SP member (custom groups / classic sites)
        const directMatch = members.some(
          (u) =>
            (loginName && u.LoginName?.toLowerCase() === loginName.toLowerCase()) ||
            (email && u.Email?.toLowerCase() === email.toLowerCase())
        );
        if (directMatch) return true;

        // Pass 2: nested AAD/M365 group — extract objectId from claim LoginName
        // e.g. "c:0o.c|federateddirectoryclaimprovider|{guid}"
        const aadGroupIds = members
          .filter((u) => u.PrincipalType !== 1) // skip regular SP users
          .map((u) => _extractAadObjectId(u.LoginName))
          .filter((id): id is string => id !== null);

        if (aadGroupIds.length === 0) return false;

        const checks = await Promise.all(aadGroupIds.map((id) => this._isGraphGroupMember(id)));
        return checks.some(Boolean);
      } catch {
        return false;
      }
    });
  }

  /**
   * Evaluate whether the current user can access an agent given its audience configuration.
   *
   * Rules:
   *  - If `audienceGroups` is empty/undefined → true (open to everyone).
   *  - Otherwise → true only if the user belongs to AT LEAST ONE of the listed groups.
   *
   * All group checks run in parallel for performance; we short-circuit on first positive result.
   */
  public async canAccessAgent(audienceGroups: IAudienceGroup[] | undefined): Promise<boolean> {
    if (!audienceGroups || audienceGroups.length === 0) return true;

    // Run all checks in parallel and resolve as soon as one is true
    return new Promise<boolean>((resolve) => {
      let resolved = false;
      let pending = audienceGroups.length;

      for (const entry of audienceGroups) {
        let check: Promise<boolean>;

        if (entry.type === 'user') {
          // Individual user: compare by email (stored in `mail`) against current user
          const { email } = this._context.pageContext.user;
          check = Promise.resolve(
            !!(email && entry.mail && email.toLowerCase() === entry.mail.toLowerCase())
          );
        } else if (entry.type === 'graph') {
          check = this._isGraphGroupMember(entry.id);
        } else {
          check = this._isSpGroupMember(entry.id);
        }

        check.then((isMember) => {
          if (resolved) return;
          if (isMember) { resolved = true; resolve(true); }
          else { pending--; if (pending === 0) resolve(false); }
        }).catch(() => {
          if (resolved) return;
          pending--;
          if (pending === 0) resolve(false);
        });
      }
    });
  }

  // ── Util ─────────────────────────────────────────────────────────────────

  private _cacheOrRun<T>(map: Map<string, Promise<T>>, key: string, fn: () => Promise<T>): Promise<T> {
    if (!map.has(key)) map.set(key, fn());
    return map.get(key)!;
  }
}
