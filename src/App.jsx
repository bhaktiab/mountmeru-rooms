import { useState, useEffect, useCallback, useRef } from "react";
import { PublicClientApplication } from "@azure/msal-browser";

// ─── Config ───────────────────────────────────────────────────────────────────
const CLIENT_ID   = "0c2d3aa6-1e8d-4c4a-a290-9a8590b5597b";
const TENANT_ID   = "24067079-ff6a-4c4e-a5de-7c5ac7ddf4d8";
const REDIRECT_URI = "https://mountmeru-rooms.vercel.app";
const GRAPH_SCOPES = ["Calendars.ReadWrite", "User.Read", "People.Read"];
const GRAPH_BASE   = "https://graph.microsoft.com/v1.0";
const BOOKING_TAG  = "MountmeruRoomBooking";
const AUTO_REFRESH_MS = 60_000; // 60 seconds
const SETTINGS_KEY    = "mm_room_settings_v3";

// ─── Room Definitions ─────────────────────────────────────────────────────────
const ROOMS = [
  { id: "serengeti", name: "Serengeti", capacity: 7,  color: "#C8A96E", accent: "#7A5C10", light: "#FDF3E0" },
  { id: "tarangire", name: "Tarangire", capacity: 3,  color: "#6BADA0", accent: "#1E6657", light: "#E6F5F2" },
  { id: "ruaha",     name: "Ruaha",     capacity: 2,  color: "#D47E6A", accent: "#8B3020", light: "#FDEEE9" },
];

// ─── Time Slots (8 AM – 5:30 PM, 30-min intervals) ───────────────────────────
const HOURS = Array.from({ length: 26 }, (_, i) => {
  const totalMins = 8 * 60 + i * 30;
  const h = Math.floor(totalMins / 60);
  const m = totalMins % 60;
  return {
    value: `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`,
    label: m === 0
      ? (h < 12 ? `${h} AM` : h === 12 ? "12 PM" : `${h - 12} PM`)
      : (h < 12 ? `${h}:30 AM` : h === 12 ? "12:30 PM" : `${h - 12}:30 PM`),
  };
});

// ─── Utilities ────────────────────────────────────────────────────────────────
const todayStr    = () => new Date().toISOString().split("T")[0];
const isValidEmail = e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.trim());

function isPastSlot(date, hourValue) {
  return new Date(`${date}T${hourValue}:00`) < new Date();
}

function initSlots() {
  return Object.fromEntries(
    ROOMS.map(r => [r.id, Object.fromEntries(HOURS.map(h => [h.value, null]))])
  );
}

function getTimezone() {
  return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
}

function formatDate(d) {
  return new Date(d + "T12:00:00").toLocaleDateString("en-US", {
    weekday: "long", month: "long", day: "numeric", year: "numeric",
  });
}

function formatShortDate(d) {
  return new Date(d + "T12:00:00").toLocaleDateString("en-US", {
    weekday: "short", month: "short", day: "numeric",
  });
}

function addDays(dateStr, n) {
  const d = new Date(dateStr + "T12:00:00");
  d.setDate(d.getDate() + n);
  return d.toISOString().split("T")[0];
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) return JSON.parse(raw);
  } catch {}
  return { roomEmails: { serengeti: "", tarangire: "", ruaha: "" } };
}

function persistSettings(s) {
  try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(s)); } catch {}
}

// ─── Teams Detection ──────────────────────────────────────────────────────────
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

// ─── Teams SDK ────────────────────────────────────────────────────────────────
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

let _teamsToken = null;
let _teamsTokenExpiry = 0;

async function teamsAuthenticate() {
  await initTeamsSDK();
  return new Promise((resolve, reject) => {
    window.microsoftTeams.authentication.authenticate({
      url: `${REDIRECT_URI}/auth-teams.html`,
      width: 600, height: 640,
      successCallback: t => t ? resolve(t) : reject(new Error("No token returned")),
      failureCallback: r => reject(new Error(r || "Teams auth failed")),
    });
  });
}

// ─── MSAL (browser) ───────────────────────────────────────────────────────────
let _msal = null;
async function getMsal() {
  if (_msal) return _msal;
  _msal = new PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
      redirectUri: REDIRECT_URI,
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
    system: { allowNativeBroker: false },
  });
  await _msal.initialize();
  return _msal;
}

async function getToken() {
  if (isInTeams()) {
    if (_teamsToken && Date.now() < _teamsTokenExpiry - 300_000) return _teamsToken;
    _teamsToken = await teamsAuthenticate();
    _teamsTokenExpiry = Date.now() + 3_600_000;
    return _teamsToken;
  }
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
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(opts.headers || {}),
    },
  });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  if (res.status === 204) return null;
  return res.json();
}

// ─── People Search ────────────────────────────────────────────────────────────
async function searchPeople(query) {
  if (!query || query.length < 2) return [];
  try {
    const q = encodeURIComponent(query);
    const data = await gFetch(
      `/users?$filter=startswith(displayName,'${q}') or startswith(mail,'${q}') or startswith(givenName,'${q}') or startswith(surname,'${q}')&$select=displayName,mail,userPrincipalName&$top=8`
    );
    return (data?.value || [])
      .map(p => ({ name: p.displayName || "", email: p.mail || p.userPrincipalName || "" }))
      .filter(p => p.email?.includes("@"));
  } catch {
    try {
      const data = await gFetch(`/me/people?$search="${encodeURIComponent(query)}"&$select=displayName,scoredEmailAddresses&$top=8`);
      return (data?.value || [])
        .map(p => ({ name: p.displayName || "", email: p.scoredEmailAddresses?.[0]?.address || "" }))
        .filter(p => p.email?.includes("@"));
    } catch { return []; }
  }
}

// ─── Calendar: Create Event ───────────────────────────────────────────────────
async function createOutlookEvent({ roomName, roomEmail, bookerName, bookerEmail, emailList, date, startHour, endHour, meetingTitle }) {
  const tz = getTimezone();
  const attendees = emailList.filter(isValidEmail).map(e => ({
    emailAddress: { address: e.trim() }, type: "required",
  }));
  // Always include organizer
  if (bookerEmail && !emailList.map(e => e.toLowerCase()).includes(bookerEmail.toLowerCase())) {
    attendees.unshift({ emailAddress: { address: bookerEmail }, type: "required" });
  }
  // Add room resource as attendee → event appears in room's shared calendar
  if (roomEmail && isValidEmail(roomEmail)) {
    attendees.push({ emailAddress: { address: roomEmail }, type: "required" });
  }
  return gFetch("/me/events", {
    method: "POST",
    body: JSON.stringify({
      subject: meetingTitle || `[${roomName}] ${bookerName}`,
      body: {
        contentType: "HTML",
        content: `<p>Room: <strong>${roomName}</strong></p><p>Booked by: ${bookerName}</p><p>Attendees: ${attendees.length}</p><p style="display:none">${BOOKING_TAG}</p>`,
      },
      start: { dateTime: `${date}T${startHour}:00`, timeZone: tz },
      end:   { dateTime: `${date}T${endHour}:00`,   timeZone: tz },
      location: { displayName: `${roomName} — Mountmeru` },
      attendees,
      responseRequested: true,
    }),
  });
}

async function deleteOutlookEvent(id) {
  return gFetch(`/me/events/${id}`, { method: "DELETE" });
}

// ─── Calendar: Fetch ──────────────────────────────────────────────────────────
async function fetchRoomCalendar(roomEmail, date) {
  const s = encodeURIComponent(`${date}T00:00:00`);
  const e = encodeURIComponent(`${date}T23:59:59`);
  const data = await gFetch(
    `/users/${encodeURIComponent(roomEmail)}/calendarView?startDateTime=${s}&endDateTime=${e}&$select=id,subject,start,end,organizer,body,attendees&$top=50&$orderby=start/dateTime`
  );
  return data?.value || [];
}

async function fetchOwnCalendar(date) {
  const s = encodeURIComponent(`${date}T00:00:00`);
  const e = encodeURIComponent(`${date}T23:59:59`);
  const data = await gFetch(
    `/me/calendarView?startDateTime=${s}&endDateTime=${e}&$select=id,subject,start,end,organizer,body,location,attendees&$top=50&$orderby=start/dateTime`
  );
  return (data?.value || []).filter(ev => (ev.body?.content || "").includes(BOOKING_TAG));
}

// ─── Slot Builders ────────────────────────────────────────────────────────────
function eventToSlotInfo(evt) {
  const startH = (evt.start?.dateTime || "").slice(11, 16);
  const endH   = (evt.end?.dateTime   || "").slice(11, 16);
  if (!startH || !endH) return null;
  const subjectMatch = evt.subject?.match(/\] (.+)$/);
  return {
    name: subjectMatch ? subjectMatch[1] : (evt.organizer?.emailAddress?.name || evt.subject || "Reserved"),
    title: evt.subject || "",
    organizer: evt.organizer?.emailAddress?.name || "",
    organizerEmail: evt.organizer?.emailAddress?.address || "",
    startHour: startH,
    endHour: endH,
    outlookEventId: evt.id,
    attendeeCount: (evt.attendees || []).filter(a =>
      a.emailAddress?.address && !Object.values(ROOMS).some(r => false) // include all
    ).length,
    synced: true,
  };
}

function buildRoomSlots(events) {
  const roomSlots = Object.fromEntries(HOURS.map(h => [h.value, null]));
  events.forEach(evt => {
    const info = eventToSlotInfo(evt);
    if (!info) return;
    const startIdx = HOURS.findIndex(h => h.value === info.startHour);
    const endIdx   = HOURS.findIndex(h => h.value === info.endHour);
    if (startIdx === -1) return;
    const endSafe = endIdx === -1 ? HOURS.length : endIdx;
    for (let i = startIdx; i < endSafe; i++) {
      roomSlots[HOURS[i].value] = { ...info, isSpan: i > startIdx };
    }
  });
  return roomSlots;
}

function buildSlotsFromOwnCalendar(events) {
  const slots = initSlots();
  events.forEach(evt => {
    const loc = evt.location?.displayName || "";
    let room = ROOMS.find(r => loc.includes(r.name));
    if (!room) room = ROOMS.find(r => evt.subject?.includes(`[${r.name}]`));
    if (!room) return;
    const info = eventToSlotInfo(evt);
    if (!info) return;
    const startIdx = HOURS.findIndex(h => h.value === info.startHour);
    const endIdx   = HOURS.findIndex(h => h.value === info.endHour);
    if (startIdx === -1) return;
    const endSafe = endIdx === -1 ? HOURS.length : endIdx;
    for (let i = startIdx; i < endSafe; i++) {
      slots[room.id][HOURS[i].value] = { ...info, isSpan: i > startIdx };
    }
  });
  return slots;
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  // URL params (Teams single-room tab)
  const roomFilter = new URLSearchParams(window.location.search).get("room");
  const visibleRooms = roomFilter ? ROOMS.filter(r => r.id === roomFilter) : ROOMS;

  // ── State ──
  const [settings, setSettings]       = useState(loadSettings);
  const [settingsForm, setSettingsForm] = useState(() => loadSettings());
  const [activeDate, setActiveDate]   = useState(todayStr);
  const [dateBookings, setDateBookings] = useState(() => ({ [todayStr()]: initSlots() }));
  const [authState, setAuthState]     = useState("idle"); // idle | signing-in | signed-in
  const [userInfo, setUserInfo]       = useState(null);
  const [syncStatus, setSyncStatus]   = useState(""); // "" | syncing | synced | error
  const [lastSynced, setLastSynced]   = useState(null);
  const [roomCalStatus, setRoomCalStatus] = useState({}); // roomId → "ok"|"error"
  const [modal, setModal]             = useState(null); // booking modal
  const [viewModal, setViewModal]     = useState(null); // view-booking modal
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [confirmCancel, setConfirmCancel] = useState(null);
  const [form, setForm]               = useState({ name: "", email: "", title: "", endHour: "", emailInput: "", emails: [] });
  const [isLoading, setIsLoading]     = useState(false);
  const [toast, setToast]             = useState(null);
  const [peopleSugg, setPeopleSugg]   = useState([]);
  const [showSugg, setShowSugg]       = useState(false);
  const [currentTime, setCurrentTime] = useState(() => new Date());

  const searchTimer  = useRef(null);
  const refreshTimer = useRef(null);
  const today        = todayStr();
  const currentBookings = dateBookings[activeDate] || initSlots();
  const hasRoomEmails   = Object.values(settings.roomEmails || {}).some(e => e && e.includes("@"));

  // ── Toast ──
  const showToast = useCallback((msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 4500);
  }, []);

  // ── Current time clock ──
  useEffect(() => {
    const t = setInterval(() => setCurrentTime(new Date()), 30_000);
    return () => clearInterval(t);
  }, []);

  // ── Sync ──
  const doSync = useCallback(async (date, settingsOverride) => {
    const s = settingsOverride ?? settings;
    const roomEmails = s.roomEmails || {};
    setSyncStatus("syncing");
    try {
      const slots   = initSlots();
      const statMap = {};

      // 1. Try room resource calendars (shared visibility for entire org)
      for (const room of ROOMS) {
        const email = roomEmails[room.id];
        if (!email || !email.includes("@")) continue;
        try {
          const evts = await fetchRoomCalendar(email, date);
          slots[room.id] = buildRoomSlots(evts);
          statMap[room.id] = "ok";
        } catch {
          statMap[room.id] = "error";
        }
      }

      // 2. For rooms without room-calendar data, fall back to own calendar
      const needsOwn = ROOMS.some(r => !roomEmails[r.id] || statMap[r.id] === "error");
      if (needsOwn) {
        try {
          const ownEvts  = await fetchOwnCalendar(date);
          const ownSlots = buildSlotsFromOwnCalendar(ownEvts);
          ROOMS.forEach(r => {
            if (!roomEmails[r.id] || statMap[r.id] === "error") {
              slots[r.id] = ownSlots[r.id];
            }
          });
        } catch { /* ignore */ }
      }

      setRoomCalStatus(statMap);
      setDateBookings(prev => ({ ...prev, [date]: slots }));
      setSyncStatus("synced");
      setLastSynced(new Date());
    } catch (e) {
      setSyncStatus("error");
      showToast("Sync failed: " + e.message, "error");
    }
  }, [settings, showToast]);

  // ── Auth init (browser) ──
  useEffect(() => {
    (async () => {
      try {
        if (isInTeams()) { setAuthState("idle"); return; }
        const msal   = await getMsal();
        const result = await msal.handleRedirectPromise();
        if (result?.account) {
          const user = await gFetch("/me?$select=displayName,mail,userPrincipalName");
          setUserInfo(user); setAuthState("signed-in");
          showToast(`Welcome, ${user.displayName}`);
        } else {
          const accounts = msal.getAllAccounts();
          if (accounts.length) {
            const user = await gFetch("/me?$select=displayName,mail,userPrincipalName");
            setUserInfo(user); setAuthState("signed-in");
          }
        }
      } catch { setAuthState("idle"); }
    })();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Sync on auth/date change ──
  useEffect(() => {
    if (authState === "signed-in") {
      if (!dateBookings[activeDate]) setDateBookings(p => ({ ...p, [activeDate]: initSlots() }));
      doSync(activeDate);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeDate, authState]);

  // ── Auto-refresh ──
  useEffect(() => {
    if (authState !== "signed-in") return;
    refreshTimer.current = setInterval(() => doSync(activeDate), AUTO_REFRESH_MS);
    return () => clearInterval(refreshTimer.current);
  }, [activeDate, authState, doSync]);

  // ── Sync on tab focus ──
  useEffect(() => {
    const handler = () => {
      if (!document.hidden && authState === "signed-in") doSync(activeDate);
    };
    document.addEventListener("visibilitychange", handler);
    return () => document.removeEventListener("visibilitychange", handler);
  }, [activeDate, authState, doSync]);

  // ── Sign in ──
  const signIn = async () => {
    setAuthState("signing-in");
    try {
      const token = await getToken();
      const res   = await fetch(`${GRAPH_BASE}/me?$select=displayName,mail,userPrincipalName`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const user = await res.json();
      setUserInfo(user); setAuthState("signed-in");
      showToast(`Welcome, ${user.displayName}`);
      await doSync(activeDate);
    } catch (e) {
      if (!isInTeams()) return; // redirect flow, page navigates away
      setAuthState("idle");
      showToast("Sign-in failed: " + (e.message || String(e)), "error");
    }
  };

  const signOut = async () => {
    _teamsToken = null; _teamsTokenExpiry = 0;
    if (!isInTeams()) { (await getMsal()).logoutRedirect(); }
    setAuthState("idle"); setUserInfo(null); setSyncStatus(""); setLastSynced(null);
    setDateBookings({ [activeDate]: initSlots() });
  };

  // ── Navigate date ──
  const navigateDate = (delta) => {
    setActiveDate(d => {
      const next = addDays(d, delta);
      if (!dateBookings[next]) setDateBookings(p => ({ ...p, [next]: initSlots() }));
      return next;
    });
  };

  // ── Open booking modal ──
  const openModal = (roomId, startHour) => {
    if (currentBookings[roomId]?.[startHour]) return;
    if (isPastSlot(activeDate, startHour)) { showToast("Cannot book past time slots", "error"); return; }
    const hIdx = HOURS.findIndex(h => h.value === startHour);
    const defaultEnd = HOURS[Math.min(hIdx + 2, HOURS.length - 1)].value; // 1 hr default
    setForm({
      name: userInfo?.displayName || "",
      email: userInfo?.mail || userInfo?.userPrincipalName || "",
      title: "", endHour: defaultEnd, emailInput: "", emails: [],
    });
    setModal({ roomId, startHour });
  };

  // ── Attendee helpers ──
  const addEmail = (override) => {
    const e = (override ?? form.emailInput).trim();
    if (!e) return;
    if (!isValidEmail(e)) { showToast("Invalid email address", "error"); return; }
    if (form.emails.includes(e)) { showToast("Already added", "error"); return; }
    setForm(f => ({ ...f, emails: [...f.emails, e], emailInput: "" }));
    setPeopleSugg([]); setShowSugg(false);
  };

  const selectSugg = (person) => {
    if (form.emails.includes(person.email)) { showToast("Already added", "error"); return; }
    setForm(f => ({ ...f, emails: [...f.emails, person.email], emailInput: "" }));
    setPeopleSugg([]); setShowSugg(false);
  };

  const handleAttInput = (val) => {
    setForm(f => ({ ...f, emailInput: val }));
    clearTimeout(searchTimer.current);
    if (val.length < 2 || authState !== "signed-in") { setPeopleSugg([]); setShowSugg(false); return; }
    searchTimer.current = setTimeout(async () => {
      const results = await searchPeople(val);
      setPeopleSugg(results); setShowSugg(results.length > 0);
    }, 300);
  };

  const removeEmail = e => setForm(f => ({ ...f, emails: f.emails.filter(x => x !== e) }));

  // ── Apply quick duration ──
  const applyDuration = (minutes) => {
    if (!modal) return;
    const startIdx = HOURS.findIndex(h => h.value === modal.startHour);
    const endIdx   = Math.min(startIdx + Math.ceil(minutes / 30), HOURS.length - 1);
    setForm(f => ({ ...f, endHour: HOURS[endIdx].value }));
  };

  // ── Confirm booking ──
  const handleBook = async () => {
    if (!form.name.trim()) { showToast("Your name is required", "error"); return; }
    if (!form.endHour || form.endHour <= modal.startHour) { showToast("End time must be after start time", "error"); return; }

    const startIdx = HOURS.findIndex(h => h.value === modal.startHour);
    const endIdx   = HOURS.findIndex(h => h.value === form.endHour);

    // Conflict check (in case someone else booked while modal was open)
    for (let i = startIdx; i < endIdx; i++) {
      if (currentBookings[modal.roomId]?.[HOURS[i].value]) {
        showToast("This slot was just booked by someone else. Please refresh.", "error");
        setModal(null);
        await doSync(activeDate);
        return;
      }
    }

    const room = ROOMS.find(r => r.id === modal.roomId);
    setIsLoading(true);
    let outlookEventId = null;
    let outlookError   = null;

    if (authState === "signed-in") {
      try {
        const evt = await createOutlookEvent({
          roomName:    room.name,
          roomEmail:   settings.roomEmails?.[room.id] || "",
          bookerName:  form.name,
          bookerEmail: form.email,
          emailList:   form.emails,
          date:        activeDate,
          startHour:   modal.startHour,
          endHour:     form.endHour,
          meetingTitle: form.title || `[${room.name}] ${form.name}`,
        });
        outlookEventId = evt?.id;
      } catch (e) { outlookError = e.message; }
    }

    // Optimistically update UI
    const newSlots = { ...currentBookings[modal.roomId] };
    for (let i = startIdx; i < endIdx; i++) {
      newSlots[HOURS[i].value] = {
        name:          form.name,
        title:         form.title || `[${room.name}] ${form.name}`,
        organizer:     form.name,
        organizerEmail: form.email,
        startHour:     modal.startHour,
        endHour:       form.endHour,
        emails:        form.emails,
        outlookEventId,
        isSpan:        i > startIdx,
        attendeeCount: form.emails.length,
        synced:        !!outlookEventId,
      };
    }
    setDateBookings(prev => ({ ...prev, [activeDate]: { ...currentBookings, [modal.roomId]: newSlots } }));
    setModal(null);
    setIsLoading(false);

    if (outlookError) showToast(`Booked. Outlook error: ${outlookError}`, "error");
    else showToast(`${room.name} booked!${outlookEventId ? " · Invite sent" : ""}`);
  };

  // ── Cancel booking ──
  const doCancel = async (roomId, hour) => {
    const booking = currentBookings[roomId]?.[hour];
    if (!booking || booking.isSpan) return;
    if (booking.outlookEventId && authState === "signed-in") {
      try { await deleteOutlookEvent(booking.outlookEventId); }
      catch (e) { showToast("Couldn't remove from Outlook: " + e.message, "error"); }
    }
    const newSlots = { ...currentBookings[roomId] };
    Object.entries(newSlots).forEach(([h, b]) => {
      if (b?.outlookEventId === booking.outlookEventId || h === hour) newSlots[h] = null;
    });
    setDateBookings(prev => ({ ...prev, [activeDate]: { ...currentBookings, [roomId]: newSlots } }));
    setViewModal(null); setConfirmCancel(null);
    showToast("Booking cancelled");
  };

  // ── Save settings ──
  const handleSaveSettings = () => {
    const next = { ...settings, roomEmails: settingsForm.roomEmails };
    setSettings(next);
    persistSettings(next);
    setSettingsOpen(false);
    showToast("Settings saved");
    if (authState === "signed-in") doSync(activeDate, next);
  };

  // ── Helpers: UI ──
  const getTimelinePos = () => {
    if (activeDate !== today) return null;
    const now = currentTime;
    const totalMins = now.getHours() * 60 + now.getMinutes();
    const startMins = 8 * 60; const endMins = 17 * 60 + 30;
    if (totalMins < startMins || totalMins > endMins) return null;
    return (totalMins - startMins) / (endMins - startMins);
  };

  const isRoomBusyNow = (roomId) => {
    if (activeDate !== today) return null;
    const now = currentTime;
    const hh  = `${String(now.getHours()).padStart(2,"0")}:${now.getMinutes() < 30 ? "00" : "30"}`;
    const b   = currentBookings[roomId]?.[hh];
    return b && !b.isSpan ? b : null;
  };

  const canCancel = (booking) => {
    if (!booking || !userInfo) return true; // allow if no auth info
    const me = (userInfo.mail || userInfo.userPrincipalName || "").toLowerCase();
    return !booking.organizerEmail || booking.organizerEmail.toLowerCase() === me;
  };

  const endHourOptions  = modal ? HOURS.filter(h => h.value > modal.startHour) : [];
  const timelinePos     = getTimelinePos();
  const lastSyncedLabel = lastSynced
    ? lastSynced.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    : "";

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div className="app">
      <style>{CSS}</style>

      {/* ══ HEADER ══ */}
      <header className="header">
        <div className="logo">
          <div className="logo-mark">M</div>
          <div className="logo-text">
            <span className="logo-title">Mountmeru</span>
            <span className="logo-sub">Room Booking</span>
          </div>
        </div>

        <div className="header-nav">
          <button className="nav-arrow" onClick={() => navigateDate(-1)} title="Previous day">‹</button>
          <div className="header-date-wrap">
            <button
              className={"today-btn" + (activeDate === today ? " today-active" : "")}
              onClick={() => setActiveDate(today)}
            >Today</button>
            <input
              type="date"
              className="date-picker"
              value={activeDate}
              onChange={e => { if (e.target.value) setActiveDate(e.target.value); }}
            />
          </div>
          <button className="nav-arrow" onClick={() => navigateDate(1)} title="Next day">›</button>
        </div>

        <div className="header-actions">
          {authState === "signed-in" ? (
            <>
              <span className={"sync-badge " + syncStatus}>
                {syncStatus === "syncing" && "⟳ Syncing…"}
                {syncStatus === "synced"  && `✓ ${lastSyncedLabel}`}
                {syncStatus === "error"   && "⚠ Sync error"}
              </span>
              <button className="icon-btn" onClick={() => doSync(activeDate)} title="Refresh now">↻</button>
              <div className="user-chip">
                <div className="user-avatar">{(userInfo?.displayName || "?").charAt(0).toUpperCase()}</div>
                <span className="user-name">{(userInfo?.displayName || "").split(" ")[0]}</span>
              </div>
              <button className="btn btn-ghost btn-sm" onClick={signOut}>Sign out</button>
            </>
          ) : (
            <button
              className="btn btn-connect"
              onClick={signIn}
              disabled={authState === "signing-in"}
            >
              {authState === "signing-in"
                ? <><span className="spin-sm" />Connecting…</>
                : <><MsLogo />Connect Outlook</>}
            </button>
          )}
          <button className="icon-btn" title="Settings"
            onClick={() => { setSettingsForm({ ...settings }); setSettingsOpen(true); }}>⚙</button>
        </div>
      </header>

      {/* ══ SUB-HEADER ══ */}
      <div className="sub-header">
        <span className="sub-date">{formatDate(activeDate)}</span>
        {activeDate === today && <span className="chip chip-today">Today</span>}
        {hasRoomEmails && <span className="chip chip-shared">🔗 Org-wide view</span>}
        {authState !== "signed-in" && <span className="chip chip-warn">Connect Outlook to see &amp; create bookings</span>}
      </div>

      {/* ══ CONNECT BANNER ══ */}
      {authState !== "signed-in" && (
        <div className="connect-banner">
          <MsLogo />
          <span>Sign in with your Microsoft account to view bookings and reserve rooms across the organization.</span>
          <button className="btn btn-connect btn-sm" onClick={signIn} disabled={authState === "signing-in"}>
            {authState === "signing-in" ? "Connecting…" : "Connect Outlook"}
          </button>
        </div>
      )}

      {/* ══ ROOM CARDS ══ */}
      <div className="room-cards">
        {visibleRooms.map(room => {
          const booked      = Object.values(currentBookings[room.id] || {}).filter(b => b && !b.isSpan).length;
          const free        = HOURS.length - booked;
          const nowBooking  = isRoomBusyNow(room.id);
          return (
            <div key={room.id} className="room-card" style={{ "--room-color": room.color, "--room-light": room.light, "--room-accent": room.accent }}>
              <div className="room-card-top">
                <div className="room-dot" style={{ background: room.color }} />
                <span className="room-card-name">{room.name}</span>
                <span className={"room-status-badge " + (nowBooking ? "busy" : "free")}>
                  {nowBooking ? "In Use" : "Available"}
                </span>
              </div>
              <div className="room-card-cap">{room.capacity} person max</div>
              {nowBooking && (
                <div className="room-now-info">
                  <span style={{ color: room.accent, fontWeight: 700 }}>{nowBooking.name}</span>
                  {" "}until {HOURS.find(h => h.value === nowBooking.endHour)?.label || nowBooking.endHour}
                </div>
              )}
              <div className="room-card-chips">
                <span className="chip" style={{ background: room.color + "22", color: room.accent }}>{booked} booked</span>
                <span className="chip chip-free">{free} free</span>
              </div>
            </div>
          );
        })}
      </div>

      {/* ══ BOOKING GRID ══ */}
      <div className="grid-wrap">
        <div className="grid">
          {/* Column headers */}
          <div className="grid-header">
            <div className="time-gutter" />
            {visibleRooms.map(room => (
              <div key={room.id} className="col-header" style={{ borderBottomColor: room.color }}>
                <span className="col-name">{room.name}</span>
                <span className="col-cap">{room.capacity}p max</span>
              </div>
            ))}
          </div>

          {/* Rows */}
          <div className="grid-body" style={{ position: "relative" }}>
            {HOURS.map(({ value, label }, idx) => (
              <div key={value} className={"grid-row" + (idx % 2 === 1 ? " row-alt" : "")}>
                <div className="time-label">{label}</div>
                {visibleRooms.map(room => {
                  const booking = currentBookings[room.id]?.[value];
                  const past    = isPastSlot(activeDate, value);

                  if (booking?.isSpan) {
                    return (
                      <div key={room.id} className="slot-cell">
                        <div className="slot slot-span" style={{ background: room.color + "18", borderColor: room.color + "50" }} />
                      </div>
                    );
                  }

                  if (booking) {
                    const isOwn = userInfo &&
                      booking.organizerEmail?.toLowerCase() === (userInfo.mail || userInfo.userPrincipalName || "").toLowerCase();
                    return (
                      <div key={room.id} className="slot-cell">
                        <div
                          className="slot slot-booked"
                          style={{ background: room.color + "26", borderColor: room.color + "80" }}
                          onClick={() => setViewModal({ booking, roomId: room.id, hour: value, room })}
                          title={`${booking.name} · until ${HOURS.find(h => h.value === booking.endHour)?.label || booking.endHour}`}
                        >
                          <div className="booking-name" style={{ color: room.accent }}>{booking.name}</div>
                          {booking.title && booking.title !== booking.name && (
                            <div className="booking-title">{booking.title.replace(`[${room.name}] `, "")}</div>
                          )}
                          <div className="booking-meta">
                            <span>until {HOURS.find(h => h.value === booking.endHour)?.label || booking.endHour}</span>
                            {booking.outlookEventId && <span title="Synced with Outlook"> 📅</span>}
                            {booking.attendeeCount > 0 && <span title={`${booking.attendeeCount} attendees`}> 👥{booking.attendeeCount}</span>}
                            {isOwn && <span className="own-tag">you</span>}
                          </div>
                        </div>
                      </div>
                    );
                  }

                  return (
                    <div key={room.id} className="slot-cell">
                      {past
                        ? <div className="slot slot-past" />
                        : (
                          <div
                            className="slot slot-free"
                            style={{ "--room-color": room.color }}
                            onClick={() => {
                              if (authState !== "signed-in")
                                showToast("Connect Outlook first to book a room", "error");
                              else openModal(room.id, value);
                            }}
                          >
                            <span className="slot-plus">＋</span>
                            <span className="slot-book-text">Book</span>
                          </div>
                        )
                      }
                    </div>
                  );
                })}
              </div>
            ))}

            {/* Current time indicator */}
            {timelinePos !== null && (
              <div className="time-line" style={{ top: `calc(${timelinePos * 100}% - 1px)` }}>
                <div className="time-line-dot" />
                <div className="time-line-bar" />
              </div>
            )}
          </div>
        </div>
      </div>

      {/* ══ BOOKING MODAL ══ */}
      {modal && (() => {
        const room = ROOMS.find(r => r.id === modal.roomId);
        return (
          <div className="overlay" onClick={() => setModal(null)}>
            <div className="modal" onClick={e => e.stopPropagation()}>
              <div className="modal-head">
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div className="room-dot" style={{ background: room.color, width: 12, height: 12 }} />
                  <div>
                    <div className="modal-title">Book {room.name}</div>
                    <div className="modal-sub">
                      {HOURS.find(h => h.value === modal.startHour)?.label} · {formatShortDate(activeDate)} · max {room.capacity} pax
                    </div>
                  </div>
                </div>
                <button className="close-btn" onClick={() => setModal(null)}>✕</button>
              </div>

              {/* Meeting title */}
              <div className="field">
                <label className="field-lbl">Meeting Title</label>
                <input className="field-inp" placeholder="e.g. Weekly standup, Project review…" value={form.title}
                  onChange={e => setForm(f => ({ ...f, title: e.target.value }))} />
              </div>

              {/* Time */}
              <div className="field-row">
                <div className="field">
                  <label className="field-lbl">Start Time</label>
                  <input className="field-inp field-disabled"
                    value={HOURS.find(h => h.value === modal.startHour)?.label} disabled />
                </div>
                <div className="field">
                  <label className="field-lbl">End Time</label>
                  <select className="field-inp" value={form.endHour}
                    onChange={e => setForm(f => ({ ...f, endHour: e.target.value }))}>
                    <option value="">Select end time…</option>
                    {endHourOptions.map(h => <option key={h.value} value={h.value}>{h.label}</option>)}
                  </select>
                </div>
              </div>

              {/* Quick duration */}
              <div className="dur-row">
                {[30, 60, 90, 120].map(m => (
                  <button key={m} className="dur-btn" onClick={() => applyDuration(m)}>
                    {m < 60 ? `${m}m` : `${m / 60}h`}
                  </button>
                ))}
              </div>

              {/* Organizer */}
              <div className="field-row">
                <div className="field">
                  <label className="field-lbl">Your Name *</label>
                  <input className="field-inp" placeholder="Full name" value={form.name}
                    onChange={e => setForm(f => ({ ...f, name: e.target.value }))} />
                </div>
                <div className="field">
                  <label className="field-lbl">Your Email</label>
                  <input className="field-inp" type="email" placeholder="you@company.com" value={form.email}
                    onChange={e => setForm(f => ({ ...f, email: e.target.value }))} />
                </div>
              </div>

              {/* Attendees */}
              <div className="field">
                <label className="field-lbl">Invite Attendees (optional)</label>
                <div style={{ position: "relative" }}>
                  <div className="att-row">
                    <input className="field-inp" type="text" placeholder="Type a name or email address…"
                      value={form.emailInput}
                      onChange={e => handleAttInput(e.target.value)}
                      onKeyDown={e => {
                        if (e.key === "Enter" || e.key === ",") { e.preventDefault(); addEmail(); }
                        if (e.key === "Escape") setShowSugg(false);
                      }}
                      onBlur={() => setTimeout(() => setShowSugg(false), 180)}
                      onFocus={() => peopleSugg.length > 0 && setShowSugg(true)}
                      style={{ flex: 1 }}
                    />
                    <button className="btn btn-ghost btn-sm" onClick={() => addEmail()}>+ Add</button>
                  </div>
                  {showSugg && peopleSugg.length > 0 && (
                    <div className="sugg-list">
                      {peopleSugg.map(p => (
                        <div key={p.email} className="sugg-item" tabIndex={0}
                          onMouseDown={() => selectSugg(p)}
                          onKeyDown={e => e.key === "Enter" && selectSugg(p)}>
                          <div className="sugg-av">{p.name.charAt(0).toUpperCase()}</div>
                          <div>
                            <div className="sugg-name">{p.name}</div>
                            <div className="sugg-email">{p.email}</div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
                {form.emails.length > 0 && (
                  <div className="tag-list">
                    {form.emails.map(e => (
                      <span key={e} className="tag">
                        {e}<button onClick={() => removeEmail(e)}>×</button>
                      </span>
                    ))}
                  </div>
                )}
                <div className="field-hint">Type a name to search the org directory, or enter an email and press Enter</div>
              </div>

              {/* Outlook notice */}
              {authState === "signed-in" ? (
                <div className="notice notice-info">
                  <MsLogo /> Outlook calendar invite will be sent to all attendees
                </div>
              ) : (
                <div className="notice notice-warn">
                  ⚠ Connect Outlook (button at top) to send calendar invites and sync bookings
                </div>
              )}

              <div className="modal-foot">
                <button className="btn btn-ghost" onClick={() => setModal(null)}>Cancel</button>
                <button className="btn btn-primary" onClick={handleBook}
                  disabled={isLoading || !form.name.trim() || !form.endHour}>
                  {isLoading ? <><span className="spin-sm" />Saving…</> : "Confirm Booking"}
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ══ VIEW BOOKING MODAL ══ */}
      {viewModal && (
        <div className="overlay" onClick={() => setViewModal(null)}>
          <div className="modal modal-view" onClick={e => e.stopPropagation()}>
            <div className="modal-head">
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div className="room-dot" style={{ background: viewModal.room.color, width: 12, height: 12 }} />
                <div>
                  <div className="modal-title">{viewModal.booking.title || viewModal.booking.name}</div>
                  <div className="modal-sub">
                    {viewModal.room.name} ·{" "}
                    {HOURS.find(h => h.value === viewModal.booking.startHour || h.value === viewModal.hour)?.label}
                    {" – "}
                    {HOURS.find(h => h.value === viewModal.booking.endHour)?.label || viewModal.booking.endHour}
                  </div>
                </div>
              </div>
              <button className="close-btn" onClick={() => setViewModal(null)}>✕</button>
            </div>
            <div className="view-rows">
              <div className="view-row">
                <span className="view-lbl">Room</span>
                <span className="view-val">{viewModal.room.name} <span style={{ color: "#AAA" }}>({viewModal.room.capacity} pax max)</span></span>
              </div>
              <div className="view-row">
                <span className="view-lbl">Booked by</span>
                <span className="view-val">{viewModal.booking.organizer || viewModal.booking.name}</span>
              </div>
              {viewModal.booking.attendeeCount > 0 && (
                <div className="view-row">
                  <span className="view-lbl">Attendees</span>
                  <span className="view-val">{viewModal.booking.attendeeCount} people invited</span>
                </div>
              )}
              {viewModal.booking.outlookEventId && (
                <div className="view-row">
                  <span className="view-lbl">Outlook</span>
                  <span className="view-val" style={{ color: "#0078D4" }}>📅 Synced with calendar</span>
                </div>
              )}
            </div>
            <div className="modal-foot">
              <button className="btn btn-ghost" onClick={() => setViewModal(null)}>Close</button>
              {canCancel(viewModal.booking) && (
                <button className="btn btn-danger" onClick={() => {
                  setConfirmCancel({ roomId: viewModal.roomId, hour: viewModal.hour });
                  setViewModal(null);
                }}>Cancel Booking</button>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ══ CONFIRM CANCEL ══ */}
      {confirmCancel && (
        <div className="overlay" onClick={() => setConfirmCancel(null)}>
          <div className="modal modal-sm" onClick={e => e.stopPropagation()}>
            <div className="modal-title" style={{ marginBottom: 8 }}>Cancel this booking?</div>
            <p className="modal-body-text">
              This will remove the booking and delete the Outlook calendar event for all attendees.
            </p>
            <div className="modal-foot">
              <button className="btn btn-ghost" onClick={() => setConfirmCancel(null)}>Keep it</button>
              <button className="btn btn-danger" onClick={() => doCancel(confirmCancel.roomId, confirmCancel.hour)}>
                Yes, cancel booking
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ══ SETTINGS MODAL ══ */}
      {settingsOpen && (
        <div className="overlay" onClick={() => setSettingsOpen(false)}>
          <div className="modal" onClick={e => e.stopPropagation()}>
            <div className="modal-head">
              <div className="modal-title">Settings</div>
              <button className="close-btn" onClick={() => setSettingsOpen(false)}>✕</button>
            </div>

            <div className="settings-section">
              <div className="settings-section-title">Room Resource Mailboxes</div>
              <p className="settings-desc">
                Enter the Microsoft 365 resource mailbox email for each meeting room.
                When configured, <strong>all users in your organization</strong> will see the same live room
                availability — bookings made by anyone will be visible to everyone. Ask your IT admin to
                create Exchange Room Mailboxes if not already set up.
              </p>
              {ROOMS.map(room => (
                <div key={room.id} className="field" style={{ marginBottom: 14 }}>
                  <label className="field-lbl" style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span className="room-dot" style={{ background: room.color, width: 8, height: 8, flexShrink: 0 }} />
                    {room.name} mailbox email
                  </label>
                  <input
                    className="field-inp"
                    type="email"
                    placeholder={`e.g. ${room.name.toLowerCase()}@yourorg.com`}
                    value={settingsForm.roomEmails?.[room.id] || ""}
                    onChange={e => setSettingsForm(s => ({
                      ...s, roomEmails: { ...s.roomEmails, [room.id]: e.target.value },
                    }))}
                  />
                  {roomCalStatus[room.id] === "ok" && (
                    <div className="field-hint" style={{ color: "#2E7D32" }}>✓ Connected — reading shared calendar</div>
                  )}
                  {roomCalStatus[room.id] === "error" && (
                    <div className="field-hint" style={{ color: "#C0392B" }}>⚠ Cannot access — check permissions with IT admin</div>
                  )}
                </div>
              ))}
              <div className="settings-tip">
                💡 Without room mailboxes configured, you will only see bookings from your own Outlook calendar.
                Other users' bookings will not be visible.
              </div>
            </div>

            <div className="modal-foot">
              <button className="btn btn-ghost" onClick={() => setSettingsOpen(false)}>Cancel</button>
              <button className="btn btn-primary" onClick={handleSaveSettings}>Save Settings</button>
            </div>
          </div>
        </div>
      )}

      {/* ══ TOAST ══ */}
      {toast && <div className={"toast toast-" + toast.type}>{toast.msg}</div>}
    </div>
  );
}

// ─── Microsoft Logo ───────────────────────────────────────────────────────────
function MsLogo() {
  return (
    <svg width="14" height="14" viewBox="0 0 21 21" fill="none" style={{ flexShrink: 0 }}>
      <rect width="10" height="10" fill="#F25022" /><rect x="11" width="10" height="10" fill="#7FBA00" />
      <rect y="11" width="10" height="10" fill="#00A4EF" /><rect x="11" y="11" width="10" height="10" fill="#FFB900" />
    </svg>
  );
}

// ─── Styles ───────────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=Lato:wght@300;400;600;700&display=swap');
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

/* ── Layout ── */
.app { font-family: 'Lato', 'Segoe UI', sans-serif; min-height: 100vh; background: #FAF7F2; color: #2C2416; }

/* ── Header ── */
.header {
  background: #2C2416; padding: 14px 24px;
  display: flex; align-items: center; justify-content: space-between;
  gap: 12px; flex-wrap: wrap; position: sticky; top: 0; z-index: 50;
  box-shadow: 0 2px 12px rgba(0,0,0,.25);
}
.logo { display: flex; align-items: center; gap: 10px; }
.logo-mark {
  width: 36px; height: 36px; background: #C8A96E; border-radius: 9px;
  display: flex; align-items: center; justify-content: center;
  font-family: 'Playfair Display', serif; font-size: 18px; font-weight: 700; color: #2C2416;
}
.logo-text { display: flex; flex-direction: column; }
.logo-title { font-family: 'Playfair Display', serif; font-size: 18px; color: #FAF7F2; letter-spacing: .3px; line-height: 1.2; }
.logo-sub   { font-size: 9px; color: #C8A96E; letter-spacing: 2.5px; text-transform: uppercase; }

.header-nav { display: flex; align-items: center; gap: 6px; }
.nav-arrow {
  background: rgba(255,255,255,.1); border: none; color: #FAF7F2;
  width: 32px; height: 32px; border-radius: 8px; cursor: pointer;
  font-size: 20px; line-height: 1; transition: background .15s; display: flex; align-items: center; justify-content: center;
}
.nav-arrow:hover { background: rgba(255,255,255,.2); }
.header-date-wrap { display: flex; align-items: center; gap: 6px; }
.today-btn {
  padding: 6px 12px; border-radius: 7px; border: 1.5px solid rgba(255,255,255,.25);
  background: transparent; color: #FAF7F2; font-size: 12px; font-family: 'Lato', sans-serif;
  font-weight: 600; cursor: pointer; transition: all .15s; white-space: nowrap;
}
.today-btn:hover  { background: rgba(255,255,255,.15); border-color: rgba(255,255,255,.5); }
.today-active     { background: #C8A96E !important; border-color: #C8A96E !important; color: #2C2416 !important; }
.date-picker {
  background: rgba(255,255,255,.1); border: 1.5px solid rgba(255,255,255,.2);
  color: #FAF7F2; padding: 7px 10px; border-radius: 7px; font-size: 12px;
  font-family: 'Lato', sans-serif; cursor: pointer; outline: none;
}
.date-picker::-webkit-calendar-picker-indicator { filter: invert(1) opacity(.7); }

.header-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.sync-badge { font-size: 11px; font-weight: 600; padding: 4px 9px; border-radius: 20px; white-space: nowrap; }
.sync-badge.syncing { color: #C8A96E; }
.sync-badge.synced  { color: #6BCB8B; }
.sync-badge.error   { color: #E57373; }
.icon-btn {
  width: 34px; height: 34px; border-radius: 8px; border: 1.5px solid rgba(255,255,255,.2);
  background: transparent; color: #C8A96E; font-size: 16px; cursor: pointer;
  transition: all .15s; display: flex; align-items: center; justify-content: center;
}
.icon-btn:hover { background: rgba(255,255,255,.12); border-color: rgba(255,255,255,.4); }
.user-chip {
  display: flex; align-items: center; gap: 7px;
  background: rgba(255,255,255,.1); border-radius: 20px;
  padding: 4px 12px 4px 4px;
}
.user-avatar {
  width: 26px; height: 26px; border-radius: 50%;
  background: #C8A96E; color: #2C2416; font-size: 12px; font-weight: 700;
  display: flex; align-items: center; justify-content: center;
}
.user-name { color: #FAF7F2; font-size: 12px; font-weight: 600; }
.btn-connect {
  display: flex; align-items: center; gap: 8px;
  background: #0078D4; color: #fff; border: none;
  padding: 9px 16px; border-radius: 8px; font-family: 'Lato', sans-serif;
  font-weight: 700; font-size: 13px; cursor: pointer; transition: all .15s;
  white-space: nowrap;
}
.btn-connect:hover:not(:disabled) { background: #106EBE; }
.btn-connect:disabled { opacity: .65; cursor: default; }

/* ── Sub-header ── */
.sub-header {
  padding: 10px 24px; display: flex; align-items: center; gap: 10px;
  flex-wrap: wrap; border-bottom: 1px solid #EDE8E0;
}
.sub-date { font-family: 'Playfair Display', serif; font-size: 15px; color: #9A7A4A; }
.chip {
  padding: 3px 10px; border-radius: 20px; font-size: 11px;
  font-weight: 600; letter-spacing: .4px;
}
.chip-today  { background: #C8A96E22; color: #7A5C10; }
.chip-shared { background: #E3F2FD; color: #1565C0; }
.chip-warn   { background: #FFF8E1; color: #9A6F00; }
.chip-free   { background: #EEE8DF; color: #AAA; }

/* ── Connect Banner ── */
.connect-banner {
  display: flex; align-items: center; gap: 12px; padding: 14px 24px;
  background: #EBF5FB; border-bottom: 1px solid #B3D7F5;
  font-size: 13px; color: #1565C0; flex-wrap: wrap;
}
.connect-banner > span { flex: 1; min-width: 200px; }

/* ── Room Cards ── */
.room-cards {
  display: flex; gap: 12px; padding: 16px 24px 8px;
  flex-wrap: wrap;
}
.room-card {
  flex: 1 1 180px; border-radius: 12px; padding: 14px 16px;
  background: var(--room-light); border-left: 5px solid var(--room-color);
  box-shadow: 0 2px 8px rgba(0,0,0,.05);
}
.room-card-top { display: flex; align-items: center; gap: 8px; margin-bottom: 4px; }
.room-dot { border-radius: 50%; width: 10px; height: 10px; flex-shrink: 0; }
.room-card-name { font-family: 'Playfair Display', serif; font-size: 15px; flex: 1; }
.room-status-badge {
  font-size: 10px; font-weight: 700; letter-spacing: .8px; text-transform: uppercase;
  padding: 2px 8px; border-radius: 20px;
}
.room-status-badge.busy { background: #FDECEA; color: #C62828; }
.room-status-badge.free { background: #E8F5E9; color: #2E7D32; }
.room-card-cap { font-size: 11px; color: #AAA; letter-spacing: .8px; text-transform: uppercase; margin-bottom: 6px; }
.room-now-info { font-size: 12px; color: #666; margin-bottom: 6px; }
.room-card-chips { display: flex; gap: 6px; flex-wrap: wrap; }

/* ── Grid ── */
.grid-wrap { padding: 8px 24px 40px; overflow-x: auto; }
.grid { min-width: 480px; }
.grid-header {
  display: flex; padding-left: 0; margin-bottom: 4px; position: sticky; top: 64px;
  background: #FAF7F2; z-index: 10; padding-bottom: 2px;
}
.time-gutter { width: 66px; min-width: 66px; }
.col-header {
  flex: 1; padding: 8px 6px 10px; border-bottom: 3px solid transparent;
  display: flex; flex-direction: column; align-items: center;
}
.col-name { font-family: 'Playfair Display', serif; font-size: 15px; }
.col-cap  { font-size: 10px; color: #B09060; letter-spacing: .8px; text-transform: uppercase; margin-top: 2px; }

.grid-body { position: relative; }
.grid-row  { display: flex; align-items: stretch; margin-bottom: 3px; border-radius: 5px; }
.row-alt   { background: rgba(0,0,0,.018); border-radius: 5px; }
.time-label {
  width: 66px; min-width: 66px; font-size: 11px; color: #B09060;
  text-align: right; padding-right: 10px; display: flex; align-items: center; justify-content: flex-end;
  font-weight: 600;
}
.slot-cell { flex: 1; padding: 2px 4px; }

/* ── Slots ── */
.slot {
  min-height: 46px; border-radius: 7px; transition: all .14s;
  display: flex; flex-direction: column; justify-content: center;
}
.slot-free {
  cursor: pointer; background: #F0EDE6; border: 1.5px dashed #C8BFA8;
  align-items: center; justify-content: center; gap: 3px;
  color: #C0B5A5; flex-direction: row;
}
.slot-free:hover {
  background: color-mix(in srgb, var(--room-color, #C8A96E) 15%, white);
  border-color: var(--room-color, #C8A96E);
  color: var(--room-color, #C8A96E);
}
.slot-plus { font-size: 14px; }
.slot-book-text { font-size: 11px; font-weight: 700; letter-spacing: .6px; text-transform: uppercase; }
.slot-past  { background: #F7F6F4; border: 1.5px dashed #E8E3DB; opacity: .4; cursor: default; }
.slot-span  { border: 1px solid; border-top: none; min-height: 46px; }
.slot-booked {
  border: 1.5px solid; padding: 8px 10px; cursor: pointer;
  display: flex; flex-direction: column; justify-content: center; gap: 2px;
}
.slot-booked:hover { filter: brightness(.97); }

.booking-name  { font-size: 12px; font-weight: 700; line-height: 1.3; }
.booking-title { font-size: 11px; color: #888; margin-top: 1px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.booking-meta  { display: flex; align-items: center; gap: 4px; font-size: 10px; color: #AAA; margin-top: 2px; flex-wrap: wrap; }
.own-tag       { background: #0078D422; color: #0078D4; border-radius: 10px; padding: 0 6px; font-size: 9px; font-weight: 700; text-transform: uppercase; }

/* ── Current time line ── */
.time-line { position: absolute; left: 0; right: 0; pointer-events: none; display: flex; align-items: center; z-index: 5; }
.time-line-dot { width: 10px; height: 10px; border-radius: 50%; background: #E53935; flex-shrink: 0; margin-left: 60px; }
.time-line-bar { flex: 1; height: 2px; background: #E53935; opacity: .7; }

/* ── Modals ── */
.overlay {
  position: fixed; inset: 0; background: rgba(28,18,8,.55);
  display: flex; align-items: center; justify-content: center; z-index: 100;
  backdrop-filter: blur(4px); padding: 16px;
}
.modal {
  background: #FFF9F2; border-radius: 16px; padding: 28px; width: 100%; max-width: 460px;
  box-shadow: 0 28px 80px rgba(0,0,0,.28); animation: popIn .2s ease;
  max-height: 90vh; overflow-y: auto;
}
.modal-sm  { max-width: 360px; }
.modal-view { max-width: 400px; }
@keyframes popIn { from { opacity:0; transform:scale(.94) translateY(10px) } to { opacity:1; transform:scale(1) translateY(0) } }

.modal-head {
  display: flex; align-items: flex-start; justify-content: space-between;
  margin-bottom: 20px; gap: 10px;
}
.modal-title  { font-family: 'Playfair Display', serif; font-size: 20px; line-height: 1.2; }
.modal-sub    { font-size: 12px; color: #AAA; margin-top: 3px; }
.modal-body-text { font-size: 13px; color: #888; line-height: 1.6; margin-bottom: 20px; }
.close-btn {
  background: none; border: none; font-size: 18px; color: #BBB;
  cursor: pointer; padding: 2px 6px; border-radius: 6px; flex-shrink: 0; line-height: 1;
}
.close-btn:hover { background: #F0EDE6; color: #2C2416; }

/* ── Form fields ── */
.field        { margin-bottom: 14px; }
.field-row    { display: flex; gap: 10px; margin-bottom: 14px; }
.field-row .field { flex: 1; margin-bottom: 0; }
.field-lbl {
  display: block; font-size: 11px; font-weight: 700; letter-spacing: 1.2px;
  text-transform: uppercase; color: #999; margin-bottom: 5px;
}
.field-inp {
  width: 100%; padding: 10px 13px; border: 1.5px solid #DDD5C4; border-radius: 8px;
  font-family: 'Lato', sans-serif; font-size: 13px; background: #FFF;
  outline: none; transition: border .15s, box-shadow .15s; color: #2C2416;
}
.field-inp:focus { border-color: #C8A96E; box-shadow: 0 0 0 3px rgba(200,169,110,.18); }
.field-disabled { background: #F5F0E8 !important; color: #AAA !important; }
.field-hint   { font-size: 11px; color: #C0B5A5; margin-top: 5px; }

/* ── Duration quick buttons ── */
.dur-row { display: flex; gap: 6px; margin-bottom: 14px; }
.dur-btn {
  flex: 1; padding: 6px 0; border: 1.5px solid #DDD5C4; border-radius: 7px;
  background: transparent; font-family: 'Lato', sans-serif; font-size: 12px; font-weight: 700;
  color: #888; cursor: pointer; transition: all .14s;
}
.dur-btn:hover { border-color: #C8A96E; color: #7A5C10; background: #FDF3E0; }

/* ── Attendees ── */
.att-row { display: flex; gap: 8px; }
.sugg-list {
  position: absolute; top: calc(100% + 4px); left: 0; right: 56px;
  background: white; border: 1.5px solid #E8E0D4; border-radius: 9px;
  box-shadow: 0 6px 20px rgba(0,0,0,.12); z-index: 200; overflow: hidden;
}
.sugg-item {
  display: flex; align-items: center; gap: 10px;
  padding: 10px 14px; cursor: pointer; transition: background .1s;
  border-bottom: 1px solid #F5F0E8;
}
.sugg-item:last-child { border-bottom: none; }
.sugg-item:hover, .sugg-item:focus { background: #FDF3E0; outline: none; }
.sugg-av {
  width: 30px; height: 30px; border-radius: 50%; background: #EEE8DF;
  display: flex; align-items: center; justify-content: center;
  font-weight: 700; font-size: 12px; color: #B09060; flex-shrink: 0;
}
.sugg-name  { font-size: 13px; font-weight: 600; color: #2C2416; }
.sugg-email { font-size: 11px; color: #AAA; }
.tag-list { display: flex; flex-wrap: wrap; gap: 5px; margin-top: 8px; }
.tag {
  display: inline-flex; align-items: center; gap: 5px;
  background: #EEE8DF; border-radius: 20px; padding: 3px 10px;
  font-size: 12px; color: #5A4A30;
}
.tag button { background: none; border: none; cursor: pointer; color: #AAA; font-size: 14px; padding: 0; line-height: 1; }
.tag button:hover { color: #C0392B; }

/* ── Notices ── */
.notice {
  display: flex; align-items: center; gap: 8px; border-radius: 8px;
  padding: 10px 14px; margin-bottom: 20px; font-size: 12px; font-weight: 600;
}
.notice-info { background: #EBF5FB; color: #0078D4; }
.notice-warn { background: #FFF8E1; color: #9A6F00; }

/* ── Modal footer ── */
.modal-foot {
  display: flex; gap: 10px; justify-content: flex-end; margin-top: 4px;
  padding-top: 16px; border-top: 1px solid #F0EDE6;
}

/* ── View rows ── */
.view-rows { display: flex; flex-direction: column; gap: 0; margin-bottom: 8px; }
.view-row  { display: flex; padding: 10px 0; border-bottom: 1px solid #F0EDE6; gap: 16px; }
.view-row:last-child { border-bottom: none; }
.view-lbl  { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #AAA; width: 80px; flex-shrink: 0; padding-top: 1px; }
.view-val  { font-size: 13px; color: #2C2416; flex: 1; }

/* ── Settings ── */
.settings-section       { margin-bottom: 8px; }
.settings-section-title { font-family: 'Playfair Display', serif; font-size: 16px; margin-bottom: 8px; }
.settings-desc { font-size: 13px; color: #888; line-height: 1.6; margin-bottom: 16px; }
.settings-tip  { background: #FFF8E1; border-radius: 8px; padding: 10px 14px; font-size: 12px; color: #9A6F00; margin-top: 4px; }

/* ── Buttons ── */
.btn {
  padding: 10px 20px; border-radius: 8px; font-family: 'Lato', sans-serif;
  font-weight: 700; font-size: 13px; cursor: pointer; letter-spacing: .4px;
  transition: all .14s; border: none; display: inline-flex; align-items: center; gap: 7px;
  white-space: nowrap;
}
.btn-primary { background: #2C2416; color: #FAF7F2; }
.btn-primary:hover:not(:disabled) { background: #4A3D28; }
.btn-primary:disabled { opacity: .5; cursor: default; }
.btn-ghost  { background: transparent; color: #888; border: 1.5px solid #D4C8B0; }
.btn-ghost:hover { border-color: #AAA; color: #2C2416; }
.btn-danger { background: #B03A2E; color: #fff; }
.btn-danger:hover { background: #8B2D24; }
.btn-sm { padding: 7px 14px; font-size: 12px; }

/* ── Spinner ── */
.spin-sm {
  display: inline-block; width: 12px; height: 12px;
  border: 2px solid rgba(255,255,255,.35); border-top-color: #fff;
  border-radius: 50%; animation: spin .7s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }

/* ── Toast ── */
.toast {
  position: fixed; bottom: 24px; right: 24px;
  padding: 13px 18px; border-radius: 10px;
  font-size: 13px; font-weight: 600; z-index: 300;
  animation: slideUp .25s ease; max-width: 340px; line-height: 1.5;
  box-shadow: 0 8px 24px rgba(0,0,0,.22);
}
.toast-success { background: #2C2416; color: #FAF7F2; }
.toast-error   { background: #B03A2E; color: #fff; }
@keyframes slideUp { from { opacity:0; transform:translateY(12px) } to { opacity:1; transform:translateY(0) } }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #D4C8B0; border-radius: 3px; }

/* ── Responsive ── */
@media (max-width: 600px) {
  .header { padding: 10px 14px; }
  .grid-wrap { padding: 6px 10px 30px; }
  .room-cards { padding: 10px 14px 4px; gap: 8px; }
  .sub-header { padding: 8px 14px; }
  .connect-banner { padding: 10px 14px; }
  .logo-title { font-size: 15px; }
  .header-nav { order: 3; width: 100%; justify-content: center; }
  .time-label { font-size: 10px; width: 52px; min-width: 52px; }
  .time-gutter { width: 52px; min-width: 52px; }
  .time-line-dot { margin-left: 44px; }
}
`;
