/**
 * Sheets.gs
 * Sheet resolution that hinges on the stable sheet ID (gid), not the tab name.
 *
 * Why this exists:
 *   getSheetByName() breaks the moment anyone renames a tab. In live production a
 *   picker renamed the Processing tab to a stock number and the whole pick flow
 *   stopped — range protections don't help, because renaming/deleting a tab is a
 *   structural operation available to any editor regardless of cell protection.
 *
 *   Every Sheet has an immutable numeric id (getSheetId(), the gid in the URL)
 *   that survives renames. This module keeps a role→gid registry in Script
 *   Properties and resolves managed sheets by gid, transparently:
 *     - self-heals a drifted tab name back to its canonical CONFIG.SHEETS name
 *       (so the flow never breaks AND the label pickers see is restored),
 *     - falls back to a name lookup and (re)registers the gid when the gid is
 *       unknown (first run) or its sheet was deleted and recreated,
 *     - leaves the gid valid even if the name can't be restored (e.g. a name
 *       collision), because resolution never depends on the name again.
 *
 *   The registry self-populates from the current (correct) names on first
 *   access, so this is rename-proof without requiring a re-run of Setup — though
 *   Setup/Rebuild call syncRegistry() to register everything up front.
 *
 * Public API:
 *   Sheets.get(role[, ss])        → Sheet | null   (gid-resolved, self-healing)
 *   Sheets.require(role[, ss])    → Sheet          (throws if absent)
 *   Sheets.getByName(name[, ss])  → Sheet | null   (name→role→gid; passthrough if unmanaged)
 *   Sheets.repairAll([ss])        → void           (resolve+self-heal every managed sheet)
 *   Sheets.syncRegistry([ss])     → registry       (register gids for all managed sheets)
 */

const Sheets = (() => {

  const PROP_KEY  = 'SHEET_GID_REGISTRY';

  // Forms owns the response tab's name — resolve it by gid but never rename it.
  const NO_RENAME = { FORM_RESPONSES: true };

  function _load() {
    const raw = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
    if (!raw) return {};
    try { return JSON.parse(raw); } catch (_) { return {}; }
  }

  function _save(reg) {
    PropertiesService.getScriptProperties().setProperty(PROP_KEY, JSON.stringify(reg));
  }

  function _byGid(ss, gid) {
    const all = ss.getSheets();
    for (let i = 0; i < all.length; i++) {
      if (all[i].getSheetId() === gid) return all[i];
    }
    return null;
  }

  // Reverse-map a canonical name → role key. Lets name-passing call sites
  // (Ledger._getSheet, the reporting modules' _schema) resolve by gid unchanged.
  function _roleForName(name) {
    const keys = Object.keys(CONFIG.SHEETS);
    for (let i = 0; i < keys.length; i++) {
      if (CONFIG.SHEETS[keys[i]] === name) return keys[i];
    }
    return null;
  }

  // Core resolver. Mutates `reg` in place when it (re)registers a gid; the
  // boolean it returns reports whether `reg` changed, so callers batch one write.
  function _resolve(role, ss, reg) {
    const canonical = CONFIG.SHEETS[role];
    if (!canonical) throw new Error(`Sheets: unknown role '${role}'`);

    let changed = false;
    let sheet   = null;

    const gid = reg[role];
    if (gid !== undefined && gid !== null) sheet = _byGid(ss, gid);

    if (sheet) {
      // Drifted tab name → restore canonical. A collision (another tab already
      // holds the name) is swallowed: the gid is still the source of truth.
      if (!NO_RENAME[role] && sheet.getName() !== canonical) {
        try { sheet.setName(canonical); } catch (_) { /* gid still valid */ }
      }
      return { sheet, changed };
    }

    // gid unknown, or its sheet was deleted/recreated → adopt by name + register.
    sheet = ss.getSheetByName(canonical);
    if (sheet) {
      if (reg[role] !== sheet.getSheetId()) { reg[role] = sheet.getSheetId(); changed = true; }
      return { sheet, changed };
    }

    return { sheet: null, changed };  // truly absent — caller (Setup) creates it
  }

  function get(role, ss) {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    const reg = _load();
    const { sheet, changed } = _resolve(role, ss, reg);
    if (changed) _save(reg);
    return sheet;
  }

  function require(role, ss) {
    const s = get(role, ss);
    if (!s) {
      throw new Error(
        `Sheets.require: "${CONFIG.SHEETS[role]}" (role ${role}) not found. Run Setup.initialize().`
      );
    }
    return s;
  }

  function getByName(name, ss) {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    const role = _roleForName(name);
    if (role) return get(role, ss);
    return ss.getSheetByName(name);  // unmanaged sheet — plain passthrough
  }

  // Resolve + self-heal every managed sheet in a single pass (one registry write).
  // Cheap enough to call at the top of hot entry points and the health monitor.
  function repairAll(ss) {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    const reg = _load();
    let changed = false;
    Object.keys(CONFIG.SHEETS).forEach(role => {
      try { if (_resolve(role, ss, reg).changed) changed = true; } catch (_) { /* skip */ }
    });
    if (changed) _save(reg);
  }

  // Register/refresh gids for every managed sheet that currently exists by name.
  // Called by Setup/Rebuild after sheets are ensured so the registry is complete
  // before any rename can happen.
  function syncRegistry(ss) {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    const reg    = _load();
    const byName = {};
    ss.getSheets().forEach(s => { byName[s.getName()] = s.getSheetId(); });

    let changed = false;
    Object.entries(CONFIG.SHEETS).forEach(([role, name]) => {
      if (byName[name] !== undefined && reg[role] !== byName[name]) {
        reg[role] = byName[name];
        changed = true;
      }
    });
    if (changed) _save(reg);
    return reg;
  }

  return { get, require, getByName, repairAll, syncRegistry };

})();
