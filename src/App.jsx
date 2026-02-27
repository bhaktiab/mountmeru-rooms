import { useState, useEffect, useCallback } from "react";
import { PublicClientApplication } from "@azure/msal-browser";

// ─── Azure / Microsoft Graph Config ──────────────────────────────────────────
const MSAL_CONFIG = {
  clientId: "0c2d3aa6-1e8d-4c4a-a290-9a8590b5597b",
  tenantId: "24067079-ff6a-4c4e-a5de-7c5ac7ddf4d8",
  redirectUri: "https://mountmeru-rooms.vercel.app",
};
const GRAPH_SCOPES = ["Calendars.ReadWrite", "User.Read"];
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const BOOKING_TAG = "MountmeruRoomBooking";
// ─────────────────────────────────────────────────────────────────────────────

// Mount Meru Brand Colors — Red #D1111C, Yellow #FFCC00, Cyan #00AEEF, Orange #F7941E, Green #6CBE45
const ROOMS = [
  { id: "serengeti", name: "Serengeti", capacity: 7, color: "#D1111C", accent: "#8B0000", light: "#FDECEA" },
  { id: "tarangire", name: "Tarangire", capacity: 3, color: "#00AEEF", accent: "#005B8E", light: "#E6F6FD" },
  { id: "ruaha",     name: "Ruaha",     capacity: 2, color: "#F7941E", accent: "#B35200", light: "#FFF3E0" },
];

const HOURS = Array.from({ length: 26 }, (_, i) => {
  const totalMins = (8 * 60) + (i * 30);
  const h = Math.floor(totalMins / 60);
  const m = totalMins % 60;
  const value = `${h.toString().padStart(2,"0")}:${m.toString().padStart(2,"0")}`;
  const label = m === 0
    ? (h < 12 ? `${h} AM` : h === 12 ? "12 PM" : `${h-12} PM`)
    : (h < 12 ? `${h}:30 AM` : h === 12 ? "12:30 PM" : `${h-12}:30 PM`);
  return { value, label };
});

const today = new Date().toISOString().split("T")[0];

function isPastSlot(date, hourValue) {
  const now = new Date();
  const slotTime = new Date(`${date}T${hourValue}:00`);
  return slotTime < now;
}

function initSlots() {
  const s = {};
  ROOMS.forEach(r => { s[r.id] = {}; HOURS.forEach(h => { s[r.id][h.value] = null; }); });
  return s;
}

function isValidEmail(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.trim()); }

// ─── Teams detection ──────────────────────────────────────────────────────────
function isInTeams() {
  try {
    return (
      window.self !== window.top ||
      navigator.userAgent.includes("Teams") ||
      window.name === "embedded-page-container" ||
      window.name === "extension-tab-frame"
    );
  } catch { return true; }
}

// ─── MSAL (browser only) ──────────────────────────────────────────────────────
let _msal = null;

async function getMsal() {
  if (_msal) return _msal;
  _msal = new PublicClientApplication({
    auth: {
      clientId: MSAL_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${MSAL_CONFIG.tenantId}`,
      redirectUri: MSAL_CONFIG.redirectUri,
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
    system: { allowNativeBroker: false },
  });
  await _msal.initialize();
  return _msal;
}

// Load Teams SDK (if not already loaded)
let _teamsReady = false;
async function initTeamsSDK() {
  if (_teamsReady) return;
  if (!window.microsoftTeams) {
    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = "https://res.cdn.office.net/teams-js/2.22.0/js/MicrosoftTeams.min.js";
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
  }
  await window.microsoftTeams.app.initialize();
  _teamsReady = true;
}

// Teams auth — opens a Teams-managed popup to auth-teams.html which does the OAuth flow
async function teamsAuthenticate() {
  await initTeamsSDK();
  return new Promise((resolve, reject) => {
    window.microsoftTeams.authentication.authenticate({
      url: `${MSAL_CONFIG.redirectUri}/auth-teams.html`,
      width: 600,
      height: 640,
      successCallback: (token) => {
        if (token) resolve(token);
        else reject(new Error("No token returned from auth popup"));
      },
      failureCallback: (reason) => reject(new Error(reason || "Auth popup failed")),
    });
  });
}

// In-memory token cache for Teams session
let _teamsToken = null;
let _teamsTokenExpiry = 0;

async function getToken() {
  if (isInTeams()) {
    // Return cached token if still valid (with 5min buffer)
    if (_teamsToken && Date.now() < _teamsTokenExpiry - 300000) return _teamsToken;
    // Open Teams-managed auth popup
    _teamsToken = await teamsAuthenticate();
    _teamsTokenExpiry = Date.now() + 3600000; // 1hr
    return _teamsToken;
  }

  // Browser: standard MSAL flow
  const msal = await getMsal();
  const accounts = msal.getAllAccounts();
  if (!accounts.length) throw new Error("NOT_SIGNED_IN");
  try {
    const r = await msal.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: accounts[0] });
    return r.accessToken;
  } catch {
    await msal.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account: accounts[0] });
    return null;
  }
}

async function gFetch(path, opts = {}) {
  const token = await getToken();
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Graph ${res.status}: ${err}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

// Search org directory for people matching query
async function searchPeople(query) {
  if (!query || query.length < 2) return [];
  try {
    // Use /users with $filter for reliable org-wide search
    const q = encodeURIComponent(query);
    const data = await gFetch(
      `/users?$filter=startswith(displayName,'${q}') or startswith(mail,'${q}') or startswith(givenName,'${q}') or startswith(surname,'${q}')&$select=displayName,mail,userPrincipalName&$top=8`
    );
    return (data?.value || [])
      .map(p => ({
        name: p.displayName || "",
        email: p.mail || p.userPrincipalName || "",
      }))
      .filter(p => p.email && p.email.includes("@"));
  } catch {
    // Fallback: try /me/people (personal contacts + org)
    try {
      const data = await gFetch(`/me/people?$search="${encodeURIComponent(query)}"&$select=displayName,scoredEmailAddresses&$top=8`);
      return (data?.value || [])
        .map(p => ({ name: p.displayName || "", email: p.scoredEmailAddresses?.[0]?.address || "" }))
        .filter(p => p.email && p.email.includes("@"));
    } catch { return []; }
  }
}

// Build timezone-aware ISO from local date + HH:MM
function toLocalISO(date, hhmm) {
  return `${date}T${hhmm}:00`;
}

// Get user's local timezone
function getTimezone() {
  return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
}

async function createOutlookEvent({ roomName, bookerName, bookerEmail, emailList, date, startHour, endHour, meetingTitle }) {
  const tz = getTimezone();
  const attendees = emailList
    .filter(isValidEmail)
    .map(e => ({ emailAddress: { address: e.trim() }, type: "required" }));

  // Always add the organizer if not already in list
  if (bookerEmail && !emailList.map(e => e.trim().toLowerCase()).includes(bookerEmail.toLowerCase())) {
    attendees.unshift({ emailAddress: { address: bookerEmail }, type: "required" });
  }

  const body = {
    subject: meetingTitle || `[${roomName}] ${bookerName}`,
    body: {
      contentType: "HTML",
      content: `<p>Room: <strong>${roomName}</strong></p><p>Booked by: ${bookerName}</p><p>Attendees: ${attendees.length}</p><p><em>Booked via Mountmeru Room Booking</em></p><p style="display:none">${BOOKING_TAG}</p>`,
    },
    start: { dateTime: toLocalISO(date, startHour), timeZone: tz },
    end:   { dateTime: toLocalISO(date, endHour),   timeZone: tz },
    location: { displayName: `${roomName} — Mountmeru` },
    attendees,
    isOrganizer: true,
    responseRequested: true,
  };

  return gFetch("/me/events", { method: "POST", body: JSON.stringify(body) });
}

async function deleteOutlookEvent(id) {
  return gFetch(`/me/events/${id}`, { method: "DELETE" });
}

async function fetchOutlookBookings(date) {
  const tz = getTimezone();
  const start = encodeURIComponent(`${date}T00:00:00`);
  const end   = encodeURIComponent(`${date}T23:59:59`);
  const data  = await gFetch(
    `/me/calendarView?startDateTime=${start}&endDateTime=${end}&$select=id,subject,start,end,location,body,organizer&$top=50&$orderby=start/dateTime`
  );
  return (data?.value || []).filter(e => (e.body?.content || "").includes(BOOKING_TAG));
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [dateBookings, setDateBookings] = useState({ [today]: initSlots() });
  const [activeDate, setActiveDate]     = useState(today);
  const [modal, setModal]               = useState(null); // { roomId, startHour }
  const [form, setForm]                 = useState({ name: "", email: "", title: "", endHour: "", emailInput: "", emails: [] });
  const [toast, setToast]               = useState(null);
  const [authState, setAuthState]       = useState("idle"); // idle | signing-in | signed-in
  const [userInfo, setUserInfo]         = useState(null);
  const [syncStatus, setSyncStatus]     = useState(""); // "" | syncing | synced | error
  const [isLoading, setIsLoading]       = useState(false);
  const [peopleSuggestions, setPeopleSuggestions] = useState([]);
  const [showSuggestions, setShowSuggestions]     = useState(false);
  const _searchTimeout = { current: null };

  const currentBookings = dateBookings[activeDate] || initSlots();

  const showToast = (msg, type = "success") => { setToast({ msg, type }); setTimeout(() => setToast(null), 4000); };

  // ── Handle auth on page load ──
  useEffect(() => {
    (async () => {
      try {
        if (isInTeams()) {
          // In Teams: don't auto-attempt auth on load (popup requires user gesture)
          // User will click Connect Outlook which triggers the Teams auth popup
          setAuthState("idle");
          return;
        } else {
          // Browser: check for existing MSAL session or redirect result
          const msal = await getMsal();
          const result = await msal.handleRedirectPromise();
          if (result?.account) {
            const user = await gFetch("/me?$select=displayName,mail,userPrincipalName");
            setUserInfo(user);
            setAuthState("signed-in");
            showToast(`Signed in as ${user.displayName}`);
          } else {
            const accounts = msal.getAllAccounts();
            if (accounts.length) {
              const user = await gFetch("/me?$select=displayName,mail,userPrincipalName");
              setUserInfo(user);
              setAuthState("signed-in");
            }
          }
        }
      } catch (e) {
        setAuthState("idle"); // SSO failed silently — user can click Connect Outlook
      }
    })();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Sync from Outlook ──
  const syncFromOutlook = useCallback(async (date) => {
    setSyncStatus("syncing");
    try {
      const events = await fetchOutlookBookings(date);
      const slots = initSlots();
      events.forEach(evt => {
        const loc = evt.location?.displayName || "";
        const room = ROOMS.find(r => loc.startsWith(r.name));
        if (!room) return;
        const startH = evt.start.dateTime.slice(11, 16);
        const endH   = evt.end.dateTime.slice(11, 16);
        const match  = evt.subject.match(/\] (.+)$/);
        const name   = match ? match[1] : evt.subject;
        slots[room.id][startH] = { name, endHour: endH, outlookEventId: evt.id, synced: true };
      });
      setDateBookings(prev => ({ ...prev, [date]: slots }));
      setSyncStatus("synced");
    } catch (e) {
      setSyncStatus("error");
      showToast("Outlook sync failed: " + e.message, "error");
    }
  }, []);

  useEffect(() => { if (authState === "signed-in") syncFromOutlook(activeDate); }, [activeDate, authState]);

  // ── Sign in ──
  const signIn = async () => {
    setAuthState("signing-in");
    try {
      // getToken() handles both Teams popup auth and browser redirect
      const token = await getToken();
      const res = await fetch(`${GRAPH_BASE}/me?$select=displayName,mail,userPrincipalName`, {
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }
      });
      const user = await res.json();
      setUserInfo(user);
      setAuthState("signed-in");
      showToast(`Signed in as ${user.displayName}`);
      await syncFromOutlook(activeDate);
    } catch (e) {
      if (!isInTeams()) {
        if (e?.message === "NOT_SIGNED_IN") {
          // No cached session — initiate login redirect
          const msal = await getMsal();
          await msal.loginRedirect({ scopes: GRAPH_SCOPES });
          return;
        }
        // Other redirect in progress — page navigates away, not an error
        return;
      }
      setAuthState("idle");
      const msg = e?.message || String(e);
      showToast("Sign-in failed: " + msg, "error");
    }
  };

  const signOut = async () => {
    if (!isInTeams()) {
      const msal = await getMsal();
      await msal.logoutRedirect();
    }
    setAuthState("idle"); setUserInfo(null); setSyncStatus("");
  };

  // ── Open booking modal ──
  const openModal = (roomId, startHour) => {
    if (currentBookings[roomId]?.[startHour]) return;
    if (isPastSlot(activeDate, startHour)) { showToast("Cannot book a time slot in the past", "error"); return; }
    const hIdx = HOURS.findIndex(h => h.value === startHour);
    const defaultEnd = HOURS[Math.min(hIdx + 1, HOURS.length - 1)].value;
    setForm({ name: userInfo?.displayName || "", email: userInfo?.mail || userInfo?.userPrincipalName || "", title: "", endHour: defaultEnd, emailInput: "", emails: [] });
    setModal({ roomId, startHour });
  };

  // ── Add email tag ──
  const addEmail = (emailOverride) => {
    const e = (emailOverride || form.emailInput).trim();
    if (!e) return;
    if (!isValidEmail(e)) { showToast("Invalid email address", "error"); return; }
    if (form.emails.includes(e)) { showToast("Already added", "error"); return; }
    setForm(f => ({ ...f, emails: [...f.emails, e], emailInput: "" }));
    setPeopleSuggestions([]);
    setShowSuggestions(false);
  };

  const selectSuggestion = (person) => {
    if (form.emails.includes(person.email)) { showToast("Already added", "error"); return; }
    setForm(f => ({ ...f, emails: [...f.emails, person.email], emailInput: "" }));
    setPeopleSuggestions([]);
    setShowSuggestions(false);
  };

  const handleAttendeeInput = (val) => {
    setForm(f => ({ ...f, emailInput: val }));
    clearTimeout(_searchTimeout.current);
    if (val.length < 2 || authState !== "signed-in") { setPeopleSuggestions([]); setShowSuggestions(false); return; }
    _searchTimeout.current = setTimeout(async () => {
      const results = await searchPeople(val);
      setPeopleSuggestions(results);
      setShowSuggestions(results.length > 0);
    }, 300);
  };

  const removeEmail = (e) => setForm(f => ({ ...f, emails: f.emails.filter(x => x !== e) }));

  // ── Confirm booking ──
  const handleBook = async () => {
    if (!form.name.trim()) { showToast("Please enter your name", "error"); return; }
    if (!form.endHour || form.endHour <= modal.startHour) { showToast("End time must be after start time", "error"); return; }
    const room = ROOMS.find(r => r.id === modal.roomId);

    setIsLoading(true);
    let outlookEventId = null;
    let outlookError = null;

    if (authState === "signed-in") {
      try {
        const evt = await createOutlookEvent({
          roomName: room.name,
          bookerName: form.name,
          bookerEmail: form.email,
          emailList: form.emails,
          date: activeDate,
          startHour: modal.startHour,
          endHour: form.endHour,
          meetingTitle: form.title || `[${room.name}] ${form.name}`,
        });
        outlookEventId = evt?.id;
      } catch (e) {
        outlookError = e.message;
      }
    }

    // Mark all slots covered by the booking
    const startIdx = HOURS.findIndex(h => h.value === modal.startHour);
    const endIdx   = HOURS.findIndex(h => h.value === form.endHour);
    const newSlots = { ...currentBookings[modal.roomId] };
    for (let i = startIdx; i < endIdx; i++) {
      newSlots[HOURS[i].value] = {
        name: form.name,
        endHour: form.endHour,
        emails: form.emails,
        outlookEventId,
        isSpan: i > startIdx,
      };
    }

    setDateBookings(prev => ({ ...prev, [activeDate]: { ...currentBookings, [modal.roomId]: newSlots } }));
    setModal(null);
    setIsLoading(false);

    if (outlookError) showToast(`Booked locally. Outlook error: ${outlookError}`, "error");
    else showToast(`${room.name} booked!${outlookEventId ? " ✓ Outlook invite sent" : ""}`);
  };

  // ── Cancel booking ──
  const handleCancel = async (roomId, hour) => {
    const booking = currentBookings[roomId]?.[hour];
    if (!booking || booking.isSpan) return;
    if (booking.outlookEventId && authState === "signed-in") {
      try { await deleteOutlookEvent(booking.outlookEventId); }
      catch (e) { showToast("Couldn't remove from Outlook: " + e.message, "error"); }
    }
    // Clear all slots of this booking
    const newSlots = { ...currentBookings[roomId] };
    Object.entries(newSlots).forEach(([h, b]) => { if (b?.outlookEventId === booking.outlookEventId || h === hour) newSlots[h] = null; });
    setDateBookings(prev => ({ ...prev, [activeDate]: { ...currentBookings, [roomId]: newSlots } }));
    showToast("Booking cancelled");
  };

  const formatDate = d => new Date(d + "T12:00:00").toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" });

  const endHourOptions = modal ? HOURS.filter(h => h.value > modal.startHour) : [];

  return (
    <div style={{ fontFamily: '"Avenir","Century Gothic","Helvetica Neue",Helvetica,Arial,sans-serif', minHeight: "100vh", background: "#F7F7F7", color: "#1A1A1A" }}>
      <style>{`
        *{box-sizing:border-box;}
        .slot{cursor:pointer;border-radius:7px;transition:all .14s;min-height:48px;}
        .slot-free{background:#F0F0F0;border:1.5px dashed #CCCCCC;display:flex;align-items:center;justify-content:center;color:#AAAAAA;font-family:"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif;font-size:11px;letter-spacing:.8px;}
        .slot-free:hover{background:#FDECEA;border-color:#D1111C;color:#D1111C;}
        .slot-span{background:repeating-linear-gradient(45deg,transparent,transparent 4px,rgba(0,0,0,.03) 4px,rgba(0,0,0,.03) 8px);border-radius:0;min-height:48px;}
        .slot-past{background:#F5F5F5;border:1.5px dashed #E0E0E0;opacity:0.45;cursor:default;min-height:48px;}
        .cancel-btn{background:none;border:none;cursor:pointer;opacity:.4;font-size:13px;padding:3px 5px;border-radius:4px;transition:all .14s;color:#1A1A1A;}
        .cancel-btn:hover{opacity:1;background:rgba(209,17,28,.12);color:#D1111C;}
        .modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center;z-index:100;backdrop-filter:blur(3px);}
        .modal{background:#FFFFFF;border-radius:16px;padding:32px;width:440px;max-width:95vw;box-shadow:0 28px 80px rgba(0,0,0,.18);animation:popIn .2s ease;max-height:90vh;overflow-y:auto;}
        @keyframes popIn{from{opacity:0;transform:scale(.95) translateY(10px)}to{opacity:1;transform:scale(1) translateY(0)}}
        .field-label{font-family:"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif;font-size:11px;letter-spacing:1.5px;text-transform:uppercase;color:#888;display:block;margin-bottom:5px;}
        .field-input{width:100%;padding:10px 13px;border:1.5px solid #E0E0E0;border-radius:8px;font-family:"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif;font-size:13px;background:#FFF;outline:none;transition:border .15s,box-shadow .15s;color:#1A1A1A;}
        .field-input:focus{border-color:#D1111C;box-shadow:0 0 0 3px rgba(209,17,28,.12);}
        .btn{padding:10px 20px;border-radius:8px;font-family:"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif;font-weight:700;font-size:13px;cursor:pointer;letter-spacing:.5px;transition:all .14s;border:none;}
        .btn-primary{background:#D1111C;color:#FFFFFF;}
        .btn-primary:hover:not(:disabled){background:#A80E17;}
        .btn-primary:disabled{opacity:.5;cursor:not-allowed;}
        .btn-ghost{background:transparent;color:#888;border:1.5px solid #DDDDDD;}
        .btn-ghost:hover{border-color:#D1111C;color:#D1111C;}
        .email-tag{display:inline-flex;align-items:center;gap:5px;background:#FFF0F0;border-radius:20px;padding:3px 10px;font-family:"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif;font-size:12px;color:#8B0000;margin:3px;}
        .email-tag button{background:none;border:none;cursor:pointer;color:#999;font-size:13px;padding:0;line-height:1;}
        .email-tag button:hover{color:#D1111C;}
        .toast{position:fixed;bottom:26px;right:26px;padding:13px 18px;border-radius:10px;font-family:"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif;font-size:13px;font-weight:600;z-index:200;animation:slideUp .25s ease;max-width:320px;line-height:1.5;}
        @keyframes slideUp{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}
        .toast-success{background:#D1111C;color:#FFFFFF;}
        .toast-error{background:#8B0000;color:#fff;}
        .spinner{display:inline-block;width:13px;height:13px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite;vertical-align:middle;margin-right:6px;}
        @keyframes spin{to{transform:rotate(360deg)}}
        input[type=date]::-webkit-calendar-picker-indicator{filter:invert(1);}
        ::-webkit-scrollbar{width:5px;height:5px;}
        ::-webkit-scrollbar-thumb{background:#CCCCCC;border-radius:3px;}
        .row-alt{background:rgba(0,0,0,.022);}
      `}</style>

      {/* ── Header ── */}
      <div style={{ background:"#D1111C", padding:"16px 32px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:12 }}>
        <div style={{ display:"flex", alignItems:"center", gap:14 }}>
          {/* MMG Logo on white badge */}
          <div style={{ background:"#FFFFFF", borderRadius:8, padding:"5px 8px", lineHeight:0, flexShrink:0 }}>
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 80" width="42" height="34">
              {/* Red m — two arches */}
              <path d="M3,72 L3,38 C3,12 50,12 50,38 C50,12 97,12 97,38 L97,72 L85,72 L85,38 C85,24 63,24 63,38 L63,72 L37,72 L37,38 C37,24 15,24 15,38 L15,72 Z" fill="#D1111C"/>
              {/* Yellow oil drop */}
              <path d="M22,38 Q9,50 9,59 A13,13 0 0,0 35,59 Q35,50 22,38 Z" fill="#FFCC00"/>
            </svg>
          </div>
          <div>
            <div style={{ fontFamily:'"Avenir","Century Gothic","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:20, fontWeight:700, color:"#FFFFFF", letterSpacing:.3 }}>Mount Meru</div>
            <div style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:10, color:"#FFCC00", letterSpacing:3, textTransform:"uppercase", marginTop:2 }}>Room Booking</div>
          </div>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
          {authState === "signed-in" ? (
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <span style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:12, color: syncStatus==="synced"?"#FFCC00": syncStatus==="syncing"?"rgba(255,255,255,.7)":"rgba(255,255,255,.5)" }}>
                {syncStatus==="syncing"?"⟳ Syncing…": syncStatus==="synced"?"✓ Synced": syncStatus==="error"?"⚠ Sync error":""}
              </span>
              <span style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:12, color:"#FFFFFF" }}>{userInfo?.displayName}</span>
              <button className="btn btn-ghost" style={{ padding:"6px 13px", fontSize:11, border:"1px solid rgba(255,255,255,.4)", color:"#FFCC00", background:"transparent" }} onClick={signOut}>Sign out</button>
              <button className="btn" style={{ padding:"6px 13px", fontSize:11, background:"#A80E17", color:"#FFCC00", border:"none" }} onClick={() => syncFromOutlook(activeDate)}>↻ Sync</button>
            </div>
          ) : (
            <button onClick={signIn} disabled={authState==="signing-in"}
              style={{ background:"rgba(255,255,255,.15)", color:"#fff", border:"1.5px solid rgba(255,255,255,.5)", padding:"9px 16px", borderRadius:8, fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontWeight:700, fontSize:13, cursor:"pointer", display:"flex", alignItems:"center", gap:8, opacity:authState==="signing-in"?.7:1 }}>
              {authState==="signing-in"
                ? <><span className="spinner"/>Signing in…</>
                : <><svg width="14" height="14" viewBox="0 0 21 21" fill="none"><rect width="10" height="10" fill="#F25022"/><rect x="11" width="10" height="10" fill="#7FBA00"/><rect y="11" width="10" height="10" fill="#00A4EF"/><rect x="11" y="11" width="10" height="10" fill="#FFB900"/></svg>Connect Outlook</>}
            </button>
          )}
          <input type="date" value={activeDate}
            onChange={e => { const d=e.target.value; setActiveDate(d); if(!dateBookings[d]) setDateBookings(p=>({...p,[d]:initSlots()})); }}
            style={{ background:"#A80E17", border:"none", color:"#FFFFFF", padding:"10px 14px", borderRadius:8, fontSize:13, cursor:"pointer", fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif' }}
          />
        </div>
      </div>

      <div style={{ padding:"24px 32px" }}>
        <div style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:15, fontWeight:600, color:"#D1111C", marginBottom:20 }}>{formatDate(activeDate)}</div>

        {/* Room summary cards */}
        <div style={{ display:"flex", gap:12, marginBottom:24, flexWrap:"wrap" }}>
          {ROOMS.map(room => {
            const booked = Object.values(currentBookings[room.id]||{}).filter(b=>b&&!b.isSpan).length;
            return (
              <div key={room.id} style={{ background:room.light, borderRadius:12, padding:"14px 20px", flex:"1 1 140px", borderLeft:`5px solid ${room.color}`, boxShadow:"0 2px 8px rgba(0,0,0,.06)" }}>
                <div style={{ fontFamily:'"Avenir","Century Gothic","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:16, fontWeight:700, color:"#1A1A1A" }}>{room.name}</div>
                <div style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:11, color:"#999", letterSpacing:1.2, textTransform:"uppercase", marginTop:2 }}>{room.capacity} pax max</div>
                <div style={{ marginTop:8, display:"flex", gap:6 }}>
                  <span style={{ background:room.color+"22", color:room.accent, padding:"2px 9px", borderRadius:20, fontSize:10, fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontWeight:700, letterSpacing:.5, textTransform:"uppercase" }}>{booked} booked</span>
                  <span style={{ background:"#EEEEEE", color:"#999", padding:"2px 9px", borderRadius:20, fontSize:10, fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontWeight:700, letterSpacing:.5, textTransform:"uppercase" }}>{HOURS.length-booked} free</span>
                </div>
              </div>
            );
          })}
        </div>

        {/* Grid */}
        <div style={{ overflowX:"auto" }}>
          <div style={{ minWidth:540 }}>
            <div style={{ display:"flex", paddingLeft:64, marginBottom:6 }}>
              {ROOMS.map(r => (
                <div key={r.id} style={{ flex:1, fontFamily:'"Avenir","Century Gothic","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:14, fontWeight:700, color:"#1A1A1A", padding:"0 5px 7px", borderBottom:`3px solid ${r.color}` }}>{r.name}</div>
              ))}
            </div>
            {HOURS.map(({ value, label }, idx) => (
              <div key={value} className={idx%2===1?"row-alt":""} style={{ display:"flex", alignItems:"stretch", marginBottom:3, borderRadius:6 }}>
                <div style={{ width:64, minWidth:64, fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:11, color:"#D1111C", textAlign:"right", paddingRight:10, display:"flex", alignItems:"center", justifyContent:"flex-end" }}>{label}</div>
                {ROOMS.map(room => {
                  const booking = currentBookings[room.id]?.[value];
                  if (booking?.isSpan) {
                    return <div key={room.id} style={{ flex:1, padding:"3px 5px" }}><div className="slot slot-span" style={{ border:`1px solid ${room.color}40`, background:room.color+"18" }} /></div>;
                  }
                  return (
                    <div key={room.id} style={{ flex:1, padding:"3px 5px" }}>
                      {booking ? (
                        <div className="slot" style={{ background:room.color+"22", border:`1.5px solid ${room.color}80`, padding:"8px 11px", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                          <div>
                            <div style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontWeight:700, fontSize:12, color:room.accent }}>{booking.name}</div>
                            <div style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:10, color:"#999", marginTop:1 }}>
                              until {HOURS.find(h=>h.value===booking.endHour)?.label || booking.endHour}
                              {booking.outlookEventId && <span style={{ color:"#0078D4", marginLeft:5 }}>📅</span>}
                              {booking.emails?.length>0 && <span style={{ marginLeft:5 }}>👥 {booking.emails.length}</span>}
                            </div>
                          </div>
                          <button className="cancel-btn" onClick={()=>handleCancel(room.id,value)} title="Cancel">✕</button>
                        </div>
                      ) : (() => {
                          const past = isPastSlot(activeDate, value);
                          return <div className={"slot " + (past ? "slot-past" : "slot-free")} onClick={()=>!past && openModal(room.id,value)} style={past ? {cursor:"default"} : {}}>{past ? "" : "+ Book"}</div>;
                        })()}
                    </div>
                  );
                })}
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* ── Booking Modal ── */}
      {modal && (() => {
        const room = ROOMS.find(r=>r.id===modal.roomId);
        return (
          <div className="modal-overlay" onClick={()=>setModal(null)}>
            <div className="modal" onClick={e=>e.stopPropagation()}>
              {/* Title */}
              <div style={{ display:"flex", alignItems:"center", gap:9, marginBottom:4 }}>
                <div style={{ width:12, height:12, borderRadius:"50%", background:room.color, flexShrink:0 }}/>
                <div style={{ fontFamily:'"Avenir","Century Gothic","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:20, fontWeight:700, color:"#1A1A1A" }}>Book {room.name}</div>
              </div>
              <div style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:12, color:"#AAA", marginBottom:22 }}>
                {HOURS.find(h=>h.value===modal.startHour)?.label} · {activeDate} · max {room.capacity} pax
              </div>

              {/* Meeting title */}
              <div style={{ marginBottom:14 }}>
                <label className="field-label">Meeting Title</label>
                <input className="field-input" placeholder={`[${room.name}] Team Standup`} value={form.title}
                  onChange={e=>setForm(f=>({...f,title:e.target.value}))} />
              </div>

              {/* Time row */}
              <div style={{ display:"flex", gap:10, marginBottom:14 }}>
                <div style={{ flex:1 }}>
                  <label className="field-label">Start Time</label>
                  <input className="field-input" value={HOURS.find(h=>h.value===modal.startHour)?.label} disabled style={{ background:"#F5F0E8", color:"#AAA" }} />
                </div>
                <div style={{ flex:1 }}>
                  <label className="field-label">End Time</label>
                  <select className="field-input" value={form.endHour} onChange={e=>setForm(f=>({...f,endHour:e.target.value}))}>
                    <option value="">Select end time</option>
                    {endHourOptions.map(h=><option key={h.value} value={h.value}>{h.label}</option>)}
                  </select>
                </div>
              </div>

              {/* Organizer */}
              <div style={{ display:"flex", gap:10, marginBottom:14 }}>
                <div style={{ flex:1 }}>
                  <label className="field-label">Your Name *</label>
                  <input className="field-input" placeholder="Alex Kimani" value={form.name}
                    onChange={e=>setForm(f=>({...f,name:e.target.value}))} />
                </div>
                <div style={{ flex:1 }}>
                  <label className="field-label">Your Email</label>
                  <input className="field-input" type="email" placeholder="alex@company.com" value={form.email}
                    onChange={e=>setForm(f=>({...f,email:e.target.value}))} />
                </div>
              </div>

              {/* Invite attendees */}
              <div style={{ marginBottom:18 }}>
                <label className="field-label">Invite Attendees (optional)</label>
                <div style={{ position:"relative" }}>
                  <div style={{ display:"flex", gap:8 }}>
                    <input className="field-input" type="text" placeholder="Type a name or email…" value={form.emailInput}
                      onChange={e=>handleAttendeeInput(e.target.value)}
                      onKeyDown={e=>{
                        if (e.key==="Enter"||e.key===",") { e.preventDefault(); addEmail(); }
                        if (e.key==="Escape") { setShowSuggestions(false); }
                        if (e.key==="ArrowDown" && showSuggestions && peopleSuggestions.length > 0) {
                          e.preventDefault();
                          document.querySelector(".people-suggestion")?.focus();
                        }
                      }}
                      onBlur={()=>setTimeout(()=>setShowSuggestions(false), 150)}
                      onFocus={()=>peopleSuggestions.length>0 && setShowSuggestions(true)}
                      style={{ flex:1 }} />
                    <button className="btn btn-ghost" style={{ padding:"10px 14px", whiteSpace:"nowrap" }} onClick={()=>addEmail()}>+ Add</button>
                  </div>
                  {showSuggestions && peopleSuggestions.length > 0 && (
                    <div style={{ position:"absolute", top:"100%", left:0, right:48, background:"white",
                      border:"1.5px solid #E8E0D4", borderRadius:8, boxShadow:"0 4px 16px rgba(0,0,0,.10)",
                      zIndex:100, overflow:"hidden", marginTop:3 }}>
                      {peopleSuggestions.map((p, i) => (
                        <div key={p.email} className="people-suggestion" tabIndex={0}
                          onMouseDown={()=>selectSuggestion(p)}
                          onKeyDown={e=>{ if(e.key==="Enter") selectSuggestion(p); }}
                          style={{ padding:"9px 14px", cursor:"pointer", display:"flex", alignItems:"center", gap:10,
                            background: i%2===0 ? "white" : "#FAFAFA",
                            borderBottom: i < peopleSuggestions.length-1 ? "1px solid #F0EDE6" : "none" }}
                          onMouseEnter={e=>e.currentTarget.style.background="#FDECEA"}
                          onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"white":"#FAFAFA"}>
                          <div style={{ width:28, height:28, borderRadius:"50%", background:"#FDECEA",
                            display:"flex", alignItems:"center", justifyContent:"center",
                            fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontWeight:700, fontSize:11, color:"#D1111C", flexShrink:0 }}>
                            {p.name.charAt(0).toUpperCase()}
                          </div>
                          <div>
                            <div style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:13, fontWeight:600, color:"#1A1A1A" }}>{p.name}</div>
                            <div style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:11, color:"#AAA" }}>{p.email}</div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
                {form.emails.length>0 && (
                  <div style={{ marginTop:8, display:"flex", flexWrap:"wrap" }}>
                    {form.emails.map(e=>(
                      <span key={e} className="email-tag">
                        {e}<button onClick={()=>removeEmail(e)}>×</button>
                      </span>
                    ))}
                  </div>
                )}
                <div style={{ fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:11, color:"#BBB", marginTop:6 }}>Type a name to search, or enter an email directly</div>
              </div>

              {/* Outlook notice */}
              {authState==="signed-in" ? (
                <div style={{ background:"#EBF5FB", borderRadius:8, padding:"9px 14px", marginBottom:20, fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:12, color:"#0078D4", display:"flex", alignItems:"center", gap:7 }}>
                  <svg width="13" height="13" viewBox="0 0 21 21" fill="none"><rect width="10" height="10" fill="#F25022"/><rect x="11" width="10" height="10" fill="#7FBA00"/><rect y="11" width="10" height="10" fill="#00A4EF"/><rect x="11" y="11" width="10" height="10" fill="#FFB900"/></svg>
                  Outlook calendar invite will be sent to all attendees
                </div>
              ) : (
                <div style={{ background:"#FFF8E6", borderRadius:8, padding:"9px 14px", marginBottom:20, fontFamily:'"Avenir","Helvetica Neue",Helvetica,Arial,sans-serif', fontSize:12, color:"#9A6F00" }}>
                  ⚠ Connect Outlook (top right) to send calendar invites
                </div>
              )}

              <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
                <button className="btn btn-ghost" onClick={()=>setModal(null)}>Cancel</button>
                <button className="btn btn-primary" onClick={handleBook} disabled={isLoading||!form.name.trim()||!form.endHour}>
                  {isLoading?<><span className="spinner"/>Saving…</>:"Confirm Booking"}
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {toast && <div className={`toast toast-${toast.type}`}>{toast.msg}</div>}
    </div>
  );
}
