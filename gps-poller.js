/*
 * BlackBuck GPS -> SharePoint poller  (GitHub Actions edition)
 * Runs as a loop: polls every 60s for LOOP_ITERS iterations, then exits.
 * The workflow schedules it every 5 min, so updates land ~every minute.
 *
 * Secrets/config come from env (set in the workflow + repo secrets):
 *   GRAPH_TENANT, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET   (app-only Graph auth)
 *   BLACKBUCK_FLEET_OWNER_ID, BLACKBUCK_TOKEN            (BlackBuck API)
 * Node 20 (global fetch). No npm dependencies.
 */

const SITE  = 'oregenesis.sharepoint.com:/sites/WCLTransportation:';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const BB_URL = 'https://api-fms.blackbuck.com/fmsiot/api/v2/gps/tracking/details'
             + '?fleet_owner_id=' + (process.env.BLACKBUCK_FLEET_OWNER_ID || '7854267')
             + '&status=All&truck_no=&map_view=true';

const HIST_MIN_MOVE_M = 25;
const HIST_MAX_GAP_MS = 5 * 60 * 1000;

const ZONES = {
  DHOPTALA:  { lat: 19.8122116, lng: 79.3370179, radiusM: 350 },
  SASTI_CHP: { lat: 19.8112580, lng: 79.3041404, radiusM: 350 },
  BALLARPUR: { lat: 19.8383790, lng: 79.3484910, radiusM: 350 },
};
const UNLOAD_ZONE = { 'Job 1': 'SASTI_CHP', 'Job 2': 'BALLARPUR' };

function haversineM(aLat, aLng, bLat, bLng) {
  const R = 6371000, toRad = d => d * Math.PI / 180;
  const dLat = toRad(bLat - aLat), dLng = toRad(bLng - aLng);
  const s = Math.sin(dLat/2)**2 + Math.cos(toRad(aLat))*Math.cos(toRad(bLat))*Math.sin(dLng/2)**2;
  return 2 * R * Math.asin(Math.sqrt(s));
}
function zoneOf(lat, lng) {
  for (const [name, z] of Object.entries(ZONES))
    if (haversineM(lat, lng, z.lat, z.lng) <= z.radiusM) return name;
  return 'NONE';
}
const istDate = d => new Date(d).toLocaleDateString('en-CA', { timeZone: 'Asia/Kolkata' });

async function getGraphToken() {
  const body = new URLSearchParams({
    client_id: process.env.GRAPH_CLIENT_ID,
    client_secret: process.env.GRAPH_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const r = await fetch('https://login.microsoftonline.com/' + process.env.GRAPH_TENANT + '/oauth2/v2.0/token', { method: 'POST', body });
  if (!r.ok) throw new Error('Graph token ' + r.status);
  return (await r.json()).access_token;
}
const gh = t => ({ Authorization: 'Bearer ' + t, 'Content-Type': 'application/json' });

async function listItems(gt, list) {
  let url = GRAPH + '/sites/' + SITE + '/lists/' + list + '/items?$expand=fields&$top=200';
  const out = [];
  while (url) {
    const r = await fetch(url, { headers: { Authorization: 'Bearer ' + gt } });
    if (!r.ok) throw new Error('read ' + list + ' ' + r.status);
    const d = await r.json();
    (d.value || []).forEach(x => out.push(x));
    url = d['@odata.nextLink'] || null;
  }
  return out;
}

async function pollOnce() {
  const bb = process.env.BLACKBUCK_TOKEN;
  const gt = await getGraphToken();

  const r = await fetch(BB_URL, { headers: { Authorization: 'Token ' + bb, 'x-aaa-enabled': 'true', Accept: 'application/json' } });
  if (!r.ok) { const _b = await r.text().catch(function(){return '';}); let _ip=''; try { _ip = await fetch('https://api.ipify.org').then(function(x){return x.text();}); } catch(e){} console.log('BlackBuck ' + r.status + ' runnerIP=' + _ip + ' body=' + _b.slice(0,300)); return; }
  const trucks = (await r.json()).list || [];

  const devMap = {};
  (await listItems(gt, 'DeviceMap')).forEach(it => {
    const f = it.fields || {};
    if (f.BBTruck) devMap[f.BBTruck] = { vehicle: f.Vehicle || '', job: f.Job || '', code: f.Title || '' };
  });
  const liveRows = {};
  (await listItems(gt, 'GpsLive')).forEach(it => { const f = it.fields || {}; if (f.DeviceCode) liveRows[f.DeviceCode] = { id: it.id, f }; });

  const today = istDate(Date.now());
  const nowIso = new Date().toISOString();
  const nowMs = Date.now();
  let updated = 0, appended = 0;

  for (const t of trucks) {
    const map = devMap[t.truck_no];
    if (!map) continue;
    const lat = +t.latitude, lng = +t.longitude;
    if (!isFinite(lat) || !isFinite(lng)) continue;
    const zone = zoneOf(lat, lng);

    const prev = liveRows[map.code];
    const prevZone = prev ? (prev.f.LastZone || 'NONE') : 'NONE';
    let trips = prev ? (+prev.f.TripsToday || 0) : 0;
    if (!prev || prev.f.TripDate !== today) trips = 0;

    const unloadZone = UNLOAD_ZONE[map.job];
    if (unloadZone && zone === unloadZone && prevZone !== unloadZone) trips += 1;

    const hLat = prev && prev.f.HLat != null ? +prev.f.HLat : null;
    const hLng = prev && prev.f.HLng != null ? +prev.f.HLng : null;
    const hTs  = prev && prev.f.HTs ? Date.parse(prev.f.HTs) : 0;
    const moved = (hLat == null) ? Infinity : haversineM(hLat, hLng, lat, lng);
    const gap   = nowMs - hTs;
    const doAppend = (moved >= HIST_MIN_MOVE_M) || (gap >= HIST_MAX_GAP_MS);

    if (doAppend) {
      const hist = {
        Title: map.code, DeviceCode: map.code, Vehicle: map.vehicle, Job: map.job,
        Latitude: lat, Longitude: lng, Speed: +t.current_speed || 0,
        Ignition: t.ignition_status || '', Zone: zone, Ts: nowIso,
      };
      try { await fetch(GRAPH + '/sites/' + SITE + '/lists/GpsHistory/items', { method: 'POST', headers: gh(gt), body: JSON.stringify({ fields: hist }) }); appended++; }
      catch (e) { console.log('GpsHistory append failed for ' + map.code); }
    }

    const fields = {
      Title: map.code, DeviceCode: map.code, Vehicle: map.vehicle, Job: map.job,
      Latitude: lat, Longitude: lng, Speed: +t.current_speed || 0,
      Ignition: t.ignition_status || '', TravelledToday: +t.travelled_today || 0,
      TripsToday: trips, LastZone: zone, TripDate: today,
      Address: (t.address || '').toString().slice(0, 255), LastUpdate: nowIso,
      HLat: doAppend ? lat : (hLat != null ? hLat : lat),
      HLng: doAppend ? lng : (hLng != null ? hLng : lng),
      HTs:  doAppend ? nowIso : (prev && prev.f.HTs ? prev.f.HTs : nowIso),
    };

    if (prev) await fetch(GRAPH + '/sites/' + SITE + '/lists/GpsLive/items/' + prev.id + '/fields', { method: 'PATCH', headers: gh(gt), body: JSON.stringify(fields) });
    else      await fetch(GRAPH + '/sites/' + SITE + '/lists/GpsLive/items', { method: 'POST', headers: gh(gt), body: JSON.stringify({ fields }) });
    updated++;
  }
  console.log('poll OK — ' + updated + ' live, ' + appended + ' history points @ ' + nowIso);
}

(async () => {
  const iters = parseInt(process.env.LOOP_ITERS || '5', 10);
  for (let i = 0; i < iters; i++) {
    try { await pollOnce(); } catch (e) { console.log('cycle error: ' + (e && e.message)); }
    if (i < iters - 1) await new Promise(r => setTimeout(r, 60000));
  }
})();
