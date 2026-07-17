import express, { Express, Request, Response } from 'express';
import {
    connect, getStats, getSignedIn, getUserFromPin, healthPing, punch,
    getStudentPortal, getReport, getRoster, getSeasons, getChecklistOverview, signOutAll, signOutOne,
} from './sheet';

const app: Express = express();
const port = process.env.PORT || 3000;

// Coach-app config (were hardcoded in the Apps Script; env-driven here).
const COACH_PASSWORD = process.env.COACH_PASSWORD || '';
const SIGNOUT_ALL_PIN = process.env.SIGNOUT_ALL_PIN || '9999';

/** Guard a coach-only route by the shared password (?password=...). */
function requireCoach(req: Request, res: Response): boolean {
    const password = String(req.query.password || '');
    if (!COACH_PASSWORD || password !== COACH_PASSWORD) {
        res.status(403).json({ success: false, message: 'Incorrect password.' });
        return false;
    }
    return true;
}

/** Guard a mentor-only route by their personal PIN (?mentorPin=...). View-only:
 *  proves "you're a mentor," same as /stats already does — write actions
 *  (signing people out) still go through requireCoach separately. */
async function requireMentor(req: Request, res: Response): Promise<boolean> {
    const mentorPin = String(req.query.mentorPin || '');
    if (!mentorPin) {
        res.status(401).json({ success: false, message: 'Mentor PIN required.' });
        return false;
    }
    const user = await getUserFromPin(mentorPin);
    if (!user || (user.get('type') || '').toUpperCase() !== 'MENTOR') {
        res.status(403).json({ success: false, message: 'Not a mentor PIN.' });
        return false;
    }
    return true;
}

// ─────────────────────────────────────────────────────────────
// Crash safety
// A single unhandled rejection used to be able to take the whole process into a
// bad state (several sheet saves were fire-and-forget). We now await everything,
// but keep these as a backstop: log loudly, and on a truly uncaught exception
// exit so the container restarts clean rather than limping.
// ─────────────────────────────────────────────────────────────
process.on('unhandledRejection', (reason) => {
    console.error('[unhandledRejection]', reason);
});
process.on('uncaughtException', (err) => {
    console.error('[uncaughtException]', err);
    process.exit(1);
});

// Kick off the first connection (non-fatal if it fails; requests will retry).
connect().catch((e) => console.error('[startup connect failed]', e?.message || e));

app.use(express.static('./public'));

// ─────────────────────────────────────────────────────────────
// Punch in/out.  /punch is the canonical name; /login kept as an alias so the
// old beartime frontend keeps working.
//   ?pin=1234&sessionType=Outreach&eventName=Food%20Drive
// ─────────────────────────────────────────────────────────────
async function handlePunch(req: Request, res: Response) {
    try {
        const pin = String(req.query.pin || '').trim();
        if (!pin) return res.status(400).json({ success: false, message: 'PIN required' });

        // Kiosk end-of-night: the sign-out-all PIN signs everyone out at once.
        if (pin === SIGNOUT_ALL_PIN) {
            const { count } = await signOutAll();
            const message = count === 0
                ? 'No students were currently signed in.'
                : `Signed out ${count} ${count === 1 ? 'person' : 'people'}. Good night!`;
            return res.json({ success: true, event: 'SIGNOUT_ALL', count, message, msg: message });
        }

        const sessionType = req.query.sessionType ? String(req.query.sessionType) : undefined;
        const eventName = req.query.eventName ? String(req.query.eventName) : undefined;

        const result = await punch(pin, sessionType, eventName);
        if (!result.success) return res.status(404).json(result);

        // Back-compat field for the old frontend, which read `msg`.
        return res.json({ ...result, msg: result.message });
    } catch (e: any) {
        console.error('[punch error]', e?.message || e);
        return res.status(500).json({ success: false, message: 'Server error. Try again.' });
    }
}
app.get('/punch', handlePunch);
app.get('/login', handlePunch);

// Who's currently signed in (kiosk polls this).
app.get('/signedin', async (_req: Request, res: Response) => {
    try {
        const data = await getSignedIn();
        return res.json({ success: true, ...data });
    } catch (e: any) {
        console.error('[signedin error]', e?.message || e);
        return res.status(500).json({ success: false, mentors: [], students: [] });
    }
});

// Mentor-only stats.
app.get('/stats', async (req: Request, res: Response) => {
    try {
        const pin = req.query.pin as string | undefined;
        if (!pin) return res.status(401).json({ error: 'PIN required' });

        const user = await getUserFromPin(pin);
        if (!user) return res.status(403).json({ error: 'Invalid PIN' });
        if ((user.get('type') || '').toUpperCase() !== 'MENTOR') {
            return res.status(403).json({ error: 'Access denied' });
        }

        return res.json(await getStats());
    } catch (e: any) {
        console.error('[stats error]', e?.message || e);
        return res.status(500).json({ error: 'Server error' });
    }
});

// ─────────────────────────────────────────────────────────────
// Argonauts coach-app parity endpoints (ported from the Apps Script).
// ─────────────────────────────────────────────────────────────

// Student self-service portal (PIN-gated, no coach password).
app.get('/portal', async (req: Request, res: Response) => {
    try {
        const pin = String(req.query.pin || '').trim();
        if (!pin) return res.status(400).json({ success: false, message: 'PIN required' });
        return res.json(await getStudentPortal(pin));
    } catch (e: any) {
        console.error('[portal error]', e?.message || e);
        return res.status(500).json({ success: false, message: 'Server error. Try again.' });
    }
});

// Coach date-range report.
app.get('/report', async (req: Request, res: Response) => {
    if (!requireCoach(req, res)) return;
    try {
        const startDate = String(req.query.startDate || '');
        const endDate = String(req.query.endDate || '');
        return res.json(await getReport(startDate, endDate));
    } catch (e: any) {
        console.error('[report error]', e?.message || e);
        return res.status(500).json({ success: false, message: 'Server error.' });
    }
});

// Coach roster with live in/out status.
app.get('/roster', async (req: Request, res: Response) => {
    if (!requireCoach(req, res)) return;
    try {
        return res.json(await getRoster());
    } catch (e: any) {
        console.error('[roster error]', e?.message || e);
        return res.status(500).json({ success: false, message: 'Server error.' });
    }
});

// Coach seasons list (for the report date presets).
app.get('/seasons', async (req: Request, res: Response) => {
    if (!requireCoach(req, res)) return;
    try {
        return res.json(await getSeasons());
    } catch (e: any) {
        console.error('[seasons error]', e?.message || e);
        return res.status(500).json({ success: false, message: 'Server error.' });
    }
});

// Coach checklist overview grid.
app.get('/checklist', async (req: Request, res: Response) => {
    if (!requireCoach(req, res)) return;
    try {
        return res.json(await getChecklistOverview());
    } catch (e: any) {
        console.error('[checklist error]', e?.message || e);
        return res.status(500).json({ success: false, message: 'Server error.' });
    }
});

// Coach sign-out-all (password-gated; the kiosk uses the PIN path in /punch).
app.get('/signoutall', async (req: Request, res: Response) => {
    if (!requireCoach(req, res)) return;
    try {
        const { count } = await signOutAll();
        return res.json({ success: true, count });
    } catch (e: any) {
        console.error('[signoutall error]', e?.message || e);
        return res.status(500).json({ success: false, message: 'Server error.' });
    }
});

// Coach sign-out for one specific person (mentor portal's per-row button).
// Password-gated like /signoutall — a mentor's own PIN only unlocks viewing.
app.get('/signout', async (req: Request, res: Response) => {
    if (!requireCoach(req, res)) return;
    try {
        const pin = String(req.query.pin || '').trim();
        if (!pin) return res.status(400).json({ success: false, message: 'PIN required.' });
        return res.json(await signOutOne(pin));
    } catch (e: any) {
        console.error('[signout error]', e?.message || e);
        return res.status(500).json({ success: false, message: 'Server error.' });
    }
});

// Mentor-portal roster view (?mentorPin=... proves "you're a mentor"; no
// write actions here — those still require the shared coach password above).
app.get('/mentor/roster', async (req: Request, res: Response) => {
    if (!(await requireMentor(req, res))) return;
    try {
        return res.json(await getRoster());
    } catch (e: any) {
        console.error('[mentor/roster error]', e?.message || e);
        return res.status(500).json({ success: false, message: 'Server error.' });
    }
});

// Which commit is actually running. Baked in at build time (Dockerfile ARG
// GIT_SHA). Lets a client poll after a push to see when their change is live.
const GIT_SHA = process.env.GIT_SHA || 'unknown';
app.get('/version', (_req: Request, res: Response) => {
    res.json({ commit: GIT_SHA, uptimeSeconds: Math.round(process.uptime()) });
});

// ─────────────────────────────────────────────────────────────
// Health check — the proper replacement for the old "restart every 24h" cron.
// Coolify hits this; if the sheet connection is wedged it returns 503 and the
// container is restarted. The watchdog below is a second, self-contained net.
// ─────────────────────────────────────────────────────────────
let lastHealthy = Date.now();
app.get('/health', async (_req: Request, res: Response) => {
    try {
        await healthPing();
        lastHealthy = Date.now();
        return res.json({ ok: true });
    } catch (e: any) {
        console.error('[health check failed]', e?.message || e);
        return res.status(503).json({ ok: false, error: e?.message || 'unhealthy' });
    }
});

// Self-watchdog: probe the sheet on an interval; if it's been unreachable for
// too long, exit so Docker's restart policy brings up a fresh process (which
// re-runs connect()). This is what actually cures the "stops working after a
// week" wedge without a blunt scheduled restart.
const WATCHDOG_INTERVAL_MS = 5 * 60 * 1000; // probe every 5 min
const WATCHDOG_MAX_UNHEALTHY_MS = 15 * 60 * 1000; // give up after 15 min unhealthy
setInterval(async () => {
    try {
        await healthPing();
        lastHealthy = Date.now();
    } catch (e: any) {
        const down = Date.now() - lastHealthy;
        console.error(`[watchdog] unhealthy for ${Math.round(down / 1000)}s:`, e?.message || e);
        if (down > WATCHDOG_MAX_UNHEALTHY_MS) {
            console.error('[watchdog] exceeded max unhealthy window, exiting for restart');
            process.exit(1);
        }
    }
}, WATCHDOG_INTERVAL_MS).unref();

app.listen(port, () => {
    console.log(`argotrack running at http://localhost:${port}`);
});
