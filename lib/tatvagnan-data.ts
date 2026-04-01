import { query, DB_SCHEMA } from './db';

export interface TatvagnanRow {
  group: string;
  lang: string;
  old_qty: number;
  new_qty: number;
  rmv_qty: number;
  total_qty: number;
}

function langToCode(lang: string): string {
  const u = (lang || '').trim().toUpperCase();
  if (u === 'GUJARATI' || u === 'GUJRATI' || u === 'GUJ') return 'GJ';
  if (u === 'HINDI' || u === 'HIN') return 'HN';
  if (u === 'MARATHI' || u === 'MAR') return 'MH';
  if (u.length >= 2) return u.slice(0, 2);
  return u || '';
}

/**
 * Fetch Tatvagnan data for PDF - same logic as Excel:
 * 1) Try tatvagnan_simple_data (Excel source): OLD = previous pushp (ADD-REMOVE), NEW/RMV = current pushp
 * 2) Fallback to tatvagnan_data (old_qty, new_qty, rmv_qty columns)
 */
export async function fetchTatvagnanDataForPdf(pushpNo: string): Promise<TatvagnanRow[]> {
  const pushpInt = parseInt(pushpNo, 10);
  if (isNaN(pushpInt)) return [];

  // 1) Try tatvagnan_simple_data (Excel table)
  try {
    const oldSql = `
      WITH base AS (
        SELECT
          UPPER(TRIM(COALESCE("group", ''))) AS grp,
          UPPER(TRIM(COALESCE(village, ''))) AS village_name,
          UPPER(TRIM(COALESCE(language, ''))) AS lang,
          UPPER(TRIM(COALESCE(member, ''))) AS member_name,
          REGEXP_REPLACE(COALESCE(mobile_number, ''), '[^0-9]', '', 'g') AS mobile,
          action_type
        FROM ${DB_SCHEMA}.tatvagnan_simple_data
        WHERE pushp_no::integer < $1
      ),
      matched AS (
        SELECT
          grp,
          lang,
          village_name,
          member_name,
          mobile,
          SUM(CASE WHEN action_type = 'ADD' THEN 1 ELSE 0 END) AS add_cnt,
          SUM(CASE WHEN action_type = 'REMOVE' THEN 1 ELSE 0 END) AS rmv_cnt
        FROM base
        GROUP BY grp, lang, village_name, member_name, mobile
      ),
      net AS (
        SELECT
          grp,
          lang,
          GREATEST(add_cnt - rmv_cnt, 0) AS net_add,
          GREATEST(rmv_cnt - add_cnt, 0) AS net_rmv
        FROM matched
      )
      SELECT grp, lang, SUM(net_add) AS add_cnt, SUM(net_rmv) AS rmv_cnt
      FROM net
      GROUP BY grp, lang
    `;
    const newSql = `
      WITH base AS (
        SELECT
          UPPER(TRIM(COALESCE("group", ''))) AS grp,
          UPPER(TRIM(COALESCE(village, ''))) AS village_name,
          UPPER(TRIM(COALESCE(language, ''))) AS lang,
          UPPER(TRIM(COALESCE(member, ''))) AS member_name,
          REGEXP_REPLACE(COALESCE(mobile_number, ''), '[^0-9]', '', 'g') AS mobile,
          action_type
        FROM ${DB_SCHEMA}.tatvagnan_simple_data
        WHERE pushp_no = $1
      ),
      matched AS (
        SELECT
          grp,
          lang,
          village_name,
          member_name,
          mobile,
          SUM(CASE WHEN action_type = 'ADD' THEN 1 ELSE 0 END) AS add_cnt,
          SUM(CASE WHEN action_type = 'REMOVE' THEN 1 ELSE 0 END) AS rmv_cnt
        FROM base
        GROUP BY grp, lang, village_name, member_name, mobile
      ),
      net AS (
        SELECT
          grp,
          lang,
          GREATEST(add_cnt - rmv_cnt, 0) AS net_add,
          GREATEST(rmv_cnt - add_cnt, 0) AS net_rmv
        FROM matched
      )
      SELECT grp, lang, SUM(net_add) AS new_cnt, SUM(net_rmv) AS rmv_cnt
      FROM net
      GROUP BY grp, lang
    `;

    const [oldRes, newRes] = await Promise.all([
      query<{ grp: string; lang: string; add_cnt: string; rmv_cnt: string }>(oldSql, [pushpInt]),
      query<{ grp: string; lang: string; new_cnt: string; rmv_cnt: string }>(newSql, [pushpNo]),
    ]);

    if (newRes.rows.length === 0) {
      // No current pushp data in simple_data -> try tatvagnan_data
      return fetchFromTatvagnanData(pushpNo);
    }

    const key = (g: string, l: string) => `${g}|${l}`;
    const oldAddMap = new Map<string, number>();
    const oldRmvMap = new Map<string, number>();
    oldRes.rows.forEach((r) => {
      const code = langToCode(r.lang);
      const k = key(r.grp, code);
      oldAddMap.set(k, (oldAddMap.get(k) ?? 0) + (parseInt(r.add_cnt || '0', 10) || 0));
      oldRmvMap.set(k, (oldRmvMap.get(k) ?? 0) + (parseInt(r.rmv_cnt || '0', 10) || 0));
    });
    const oldMap = new Map<string, number>();
    oldAddMap.forEach((add, k) => {
      const rmv = oldRmvMap.get(k) ?? 0;
      oldMap.set(k, add - rmv);
    });
    const newMap = new Map<string, number>();
    const rmvMap = new Map<string, number>();
    newRes.rows.forEach((r) => {
      const code = langToCode(r.lang);
      const k = key(r.grp, code);
      newMap.set(k, (newMap.get(k) ?? 0) + (parseInt(r.new_cnt || '0', 10) || 0));
      rmvMap.set(k, (rmvMap.get(k) ?? 0) + (parseInt(r.rmv_cnt || '0', 10) || 0));
    });

    const allKeys = new Set<string>([...oldMap.keys(), ...newMap.keys()]);
    const out: TatvagnanRow[] = [];
    allKeys.forEach((k) => {
      const [group, lang] = k.split('|');
      const old_qty = oldMap.get(k) ?? 0;
      const new_qty = newMap.get(k) ?? 0;
      const rmv_qty = rmvMap.get(k) ?? 0;
      const total_qty = old_qty + new_qty - rmv_qty;
      out.push({
        group,
        lang,
        old_qty,
        new_qty,
        rmv_qty,
        total_qty,
      });
    });
    return out.filter((r) => r.old_qty !== 0 || r.new_qty !== 0 || r.rmv_qty !== 0 || r.total_qty !== 0);
  } catch {
    return fetchFromTatvagnanData(pushpNo);
  }
}

/**
 * Fetch data for group-wise PDF - same as VBA: key by (group, full language name), not code.
 * OLD = add - rmv for pushp_no < current; displayed as 0 if negative.
 */
export async function fetchTatvagnanGroupWiseData(pushpNo: string): Promise<TatvagnanRow[]> {
  const pushpInt = parseInt(pushpNo, 10);
  if (isNaN(pushpInt)) return [];

  try {
    const oldSql = `
      WITH base AS (
        SELECT
          UPPER(TRIM(COALESCE("group", ''))) AS grp,
          UPPER(TRIM(COALESCE(village, ''))) AS village_name,
          UPPER(TRIM(COALESCE(language, ''))) AS lang,
          UPPER(TRIM(COALESCE(member, ''))) AS member_name,
          REGEXP_REPLACE(COALESCE(mobile_number, ''), '[^0-9]', '', 'g') AS mobile,
          action_type
        FROM ${DB_SCHEMA}.tatvagnan_simple_data
        WHERE pushp_no::integer < $1
      ),
      matched AS (
        SELECT
          grp,
          lang,
          village_name,
          member_name,
          mobile,
          SUM(CASE WHEN action_type = 'ADD' THEN 1 ELSE 0 END) AS add_cnt,
          SUM(CASE WHEN action_type = 'REMOVE' THEN 1 ELSE 0 END) AS rmv_cnt
        FROM base
        GROUP BY grp, lang, village_name, member_name, mobile
      ),
      net AS (
        SELECT
          grp,
          lang,
          GREATEST(add_cnt - rmv_cnt, 0) AS net_add,
          GREATEST(rmv_cnt - add_cnt, 0) AS net_rmv
        FROM matched
      )
      SELECT grp, lang, SUM(net_add) AS add_cnt, SUM(net_rmv) AS rmv_cnt
      FROM net
      GROUP BY grp, lang
    `;
    const newSql = `
      WITH base AS (
        SELECT
          UPPER(TRIM(COALESCE("group", ''))) AS grp,
          UPPER(TRIM(COALESCE(village, ''))) AS village_name,
          UPPER(TRIM(COALESCE(language, ''))) AS lang,
          UPPER(TRIM(COALESCE(member, ''))) AS member_name,
          REGEXP_REPLACE(COALESCE(mobile_number, ''), '[^0-9]', '', 'g') AS mobile,
          action_type
        FROM ${DB_SCHEMA}.tatvagnan_simple_data
        WHERE pushp_no = $1
      ),
      matched AS (
        SELECT
          grp,
          lang,
          village_name,
          member_name,
          mobile,
          SUM(CASE WHEN action_type = 'ADD' THEN 1 ELSE 0 END) AS add_cnt,
          SUM(CASE WHEN action_type = 'REMOVE' THEN 1 ELSE 0 END) AS rmv_cnt
        FROM base
        GROUP BY grp, lang, village_name, member_name, mobile
      ),
      net AS (
        SELECT
          grp,
          lang,
          GREATEST(add_cnt - rmv_cnt, 0) AS net_add,
          GREATEST(rmv_cnt - add_cnt, 0) AS net_rmv
        FROM matched
      )
      SELECT grp, lang, SUM(net_add) AS new_cnt, SUM(net_rmv) AS rmv_cnt
      FROM net
      GROUP BY grp, lang
    `;

    const [oldRes, newRes] = await Promise.all([
      query<{ grp: string; lang: string; add_cnt: string; rmv_cnt: string }>(oldSql, [pushpInt]),
      query<{ grp: string; lang: string; new_cnt: string; rmv_cnt: string }>(newSql, [pushpNo]),
    ]);

    if (newRes.rows.length === 0) {
      return fetchFromTatvagnanData(pushpNo);
    }

    const dictData = new Map<string, number>();
    const grpLangSet = new Set<string>();

    oldRes.rows.forEach((r) => {
      const grp = (r.grp || '').trim();
      const lang = (r.lang || '').trim();
      grpLangSet.add(`${grp}|${lang}`);
      const val = (parseInt(r.add_cnt || '0', 10) || 0) - (parseInt(r.rmv_cnt || '0', 10) || 0);
      dictData.set(`${grp}|${lang}|OLD`, val);
    });
    newRes.rows.forEach((r) => {
      const grp = (r.grp || '').trim();
      const lang = (r.lang || '').trim();
      grpLangSet.add(`${grp}|${lang}`);
      dictData.set(`${grp}|${lang}|NEW`, parseInt(r.new_cnt || '0', 10) || 0);
      dictData.set(`${grp}|${lang}|RMV`, parseInt(r.rmv_cnt || '0', 10) || 0);
    });

    const out: TatvagnanRow[] = [];
    grpLangSet.forEach((k) => {
      const pipe = k.indexOf('|');
      const group = pipe >= 0 ? k.slice(0, pipe) : k;
      const lang = pipe >= 0 ? k.slice(pipe + 1) : '';
      let old_qty = dictData.get(`${group}|${lang}|OLD`) ?? 0;
      if (old_qty < 0) old_qty = 0;
      const new_qty = dictData.get(`${group}|${lang}|NEW`) ?? 0;
      const rmv_qty = dictData.get(`${group}|${lang}|RMV`) ?? 0;
      const total_qty = old_qty + new_qty - rmv_qty;
      const row: TatvagnanRow = {
        group,
        lang,
        old_qty,
        new_qty,
        rmv_qty,
        total_qty,
      };
      if (row.old_qty !== 0 || row.new_qty !== 0 || row.rmv_qty !== 0 || row.total_qty !== 0) {
        out.push(row);
      }
    });
    console.log('Tatvagnan GroupWise net totals rows:', out.length);
    return out;
  } catch {
    return fetchFromTatvagnanData(pushpNo);
  }
}

async function fetchFromTatvagnanData(pushpNo: string): Promise<TatvagnanRow[]> {
  const sql = `
    SELECT "group", lang AS lang, old_qty, new_qty, rmv_qty,
      (COALESCE(old_qty,0) + COALESCE(new_qty,0) - COALESCE(rmv_qty,0)) AS total_qty
    FROM ${DB_SCHEMA}.tatvagnan_data
    WHERE pushp_no = $1
    ORDER BY "group", lang
  `;
  const result = await query<TatvagnanRow>(sql, [parseInt(pushpNo, 10)]);
  return result.rows.map((r) => ({
    group: r.group,
    lang: r.lang,
    old_qty: Number(r.old_qty) || 0,
    new_qty: Number(r.new_qty) || 0,
    rmv_qty: Number(r.rmv_qty) || 0,
    total_qty: Number(r.total_qty) || 0,
  }));
}

export interface GroupVillageRow {
  group: string;
  village: string;
  language: string;
  action_type: string;
  name: string;
  member: string;
  mobile_number: string;
  address_line1: string;
  address_line2: string;
}

function normalizeText(v: string): string {
  return (v || '').trim().toUpperCase();
}

function normalizeMobile(v: string): string {
  return (v || '').replace(/\D/g, '');
}

function netGroupVillageRows(rows: GroupVillageRow[]): GroupVillageRow[] {
  type Bucket = { adds: GroupVillageRow[]; removes: GroupVillageRow[] };
  const buckets = new Map<string, Bucket>();

  const keyOf = (r: GroupVillageRow) =>
    [
      normalizeText(r.group),
      normalizeText(r.village),
      normalizeText(r.member),
      normalizeText(r.language),
      normalizeMobile(r.mobile_number),
    ].join('|');

  for (const row of rows) {
    const action = normalizeText(row.action_type);
    const key = keyOf(row);
    let bucket = buckets.get(key);
    if (!bucket) {
      bucket = { adds: [], removes: [] };
      buckets.set(key, bucket);
    }
    if (action === 'ADD') bucket.adds.push(row);
    else if (action === 'REMOVE') bucket.removes.push(row);
  }

  const out: GroupVillageRow[] = [];
  let inputAdds = 0;
  let inputRemoves = 0;
  let netAdds = 0;
  let netRemoves = 0;

  for (const bucket of buckets.values()) {
    inputAdds += bucket.adds.length;
    inputRemoves += bucket.removes.length;
    if (bucket.adds.length > bucket.removes.length) {
      const keep = bucket.adds.length - bucket.removes.length;
      netAdds += keep;
      out.push(...bucket.adds.slice(0, keep));
    } else if (bucket.removes.length > bucket.adds.length) {
      const keep = bucket.removes.length - bucket.adds.length;
      netRemoves += keep;
      out.push(...bucket.removes.slice(0, keep));
    }
  }

  console.log('Tatvagnan GroupVillage netting:', {
    inputRows: rows.length,
    inputAdds,
    inputRemoves,
    outputRows: out.length,
    outputAdds: netAdds,
    outputRemoves: netRemoves,
    fullyCancelledRows: rows.length - out.length,
    keys: buckets.size,
  });

  return out.sort((a, b) => {
    const g = a.group.localeCompare(b.group, 'en', { sensitivity: 'base' });
    if (g !== 0) return g;
    const v = a.village.localeCompare(b.village, 'en', { sensitivity: 'base' });
    if (v !== 0) return v;
    const aa = normalizeText(a.action_type) === 'ADD' ? 0 : 1;
    const bb = normalizeText(b.action_type) === 'ADD' ? 0 : 1;
    if (aa !== bb) return aa - bb;
    return a.name.localeCompare(b.name, 'en', { sensitivity: 'base' });
  });
}

/** Fetch from tatvagnan_simple_data for group+village PDF (name, member, language, mobile, address). */
export async function fetchGroupVillageForPdf(pushpNo: string): Promise<GroupVillageRow[]> {
  const mapRow = (r: any): GroupVillageRow => ({
    group: String(r.group ?? ''),
    village: String(r.village ?? ''),
    language: String(r.language ?? ''),
    action_type: String(r.action_type ?? ''),
    name: String(r.name ?? ''),
    member: String(r.member ?? ''),
    mobile_number: String(r.mobile_number ?? ''),
    address_line1: String(r.address_line1 ?? ''),
    address_line2: String(r.address_line2 ?? ''),
  });
  try {
    const sql = `
      SELECT COALESCE("group", '') AS group,
        COALESCE(village, language, '') AS village,
        COALESCE(language, '') AS language,
        action_type,
        COALESCE(name, '') AS name,
        COALESCE(member, '') AS member,
        COALESCE(mobile_number, '') AS mobile_number,
        COALESCE(address_line1, '') AS address_line1,
        COALESCE(address_line2, '') AS address_line2
      FROM ${DB_SCHEMA}.tatvagnan_simple_data
      WHERE pushp_no = $1
      ORDER BY "group", COALESCE(village, language), action_type, name
    `;
    const res = await query<GroupVillageRow>(sql, [pushpNo]);
    return netGroupVillageRows(res.rows.map(mapRow));
  } catch {
    try {
      const fallback = `
        SELECT COALESCE("group", '') AS group, COALESCE(language, '') AS village,
          COALESCE(language, '') AS language, action_type,
          COALESCE(name, '') AS name, '' AS member, '' AS mobile_number, '' AS address_line1, '' AS address_line2
        FROM ${DB_SCHEMA}.tatvagnan_simple_data WHERE pushp_no = $1
        ORDER BY "group", language, action_type
      `;
      const res = await query<GroupVillageRow>(fallback, [pushpNo]);
      return netGroupVillageRows(res.rows.map(mapRow));
    } catch {
      return [];
    }
  }
}
