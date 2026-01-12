
import 'dotenv/config';
import express from 'express';
import sql from 'mssql';
import cors from 'cors';
import { Client } from 'ldapts';
import jwt from 'jsonwebtoken';

process.on('unhandledRejection', (reason, p) => {
  console.error('Unhandled Rejection at:', p, 'reason:', reason);
});

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(cors());

const dbPort = process.env.DB_PORT ? parseInt(process.env.DB_PORT) : undefined;
// Wenn ein Port explizit gesetzt ist, darf instanceName nicht verwendet werden (direkte TCP Verbindung)
const dbInstance = dbPort ? undefined : (process.env.DB_INSTANCE || undefined);

const config = {
    user: process.env.DB_USER || 'pxm',
    password: process.env.DB_PASSWORD || 'test1234',
    server: process.env.DB_SERVER || 'localhost',
    database: process.env.DB_DATABASE || 'pxm',
    port: dbPort,
    options: {
        encrypt: false,
        trustServerCertificate: true,
        instanceName: dbInstance,
        connectTimeout: 8000,
        useUTC: false
    }
};

console.log("--- SERVER CONFIGURATION ---");
console.log("DB Config:", { ...config, password: '***' });
console.log("AD Config:", { url: process.env.AD_URL, baseDN: process.env.AD_BASE_DN, domain: process.env.AD_DOMAIN });
console.log("AD Groups:", { AD_REQUIRED_GROUP: process.env.AD_REQUIRED_GROUP, AD_READONLY_GROUP: process.env.AD_READONLY_GROUP, AD_MAINTENANCE_GROUP: process.env.AD_MAINTENANCE_GROUP });
console.log("----------------------------");

// Active Directory Configuration
const adConfig = {
    url: process.env.AD_URL || 'ldap://192.168.178.10', // e.g. ldap://192.168.1.5
    baseDN: process.env.AD_BASE_DN || 'dc=lab,dc=local', // Replace with your domain components
    // Optional: If you need a specific user to bind to AD to perform searches
    // username: 'reader@example.com', 
    // password: 'password' 
};

// Service Account for AD Search (User Picker)
const AD_SERVICE_USER = process.env.AD_SERVICE_USER || 'svc_pxm_ad_read'; // TODO: Configure Service Account
const AD_SERVICE_PASS = process.env.AD_SERVICE_PASS || 'testuser1234!'; 

const JWT_SECRET = process.env.JWT_SECRET || 'CHANGE_THIS_TO_A_SECURE_SECRET_KEY';
const AD_DOMAIN = process.env.AD_DOMAIN || 'lab.local';
const AD_REQUIRED_GROUP = process.env.AD_REQUIRED_GROUP || 'CN=g_appl_pxm,OU=groups,DC=lab,DC=local';
const AD_READONLY_GROUP = process.env.AD_READONLY_GROUP || 'CN=g_appl_pxm_portal_readonly,OU=groups,DC=lab,DC=local';
const AD_MAINTENANCE_GROUP = process.env.AD_MAINTENANCE_GROUP || 'CN=g_Appl_pxm_portal_pflege,OU=groups,DC=lab,DC=local';

// Middleware to verify JWT token
const authenticateToken = (req, res, next) => {
    const authHeader = req.headers['authorization'];
    const token = authHeader && authHeader.split(' ')[1];

    if (!token) return res.status(401).json({ error: 'Zugriff verweigert' });

    jwt.verify(token, JWT_SECRET, (err, user) => {
        if (err) return res.status(403).json({ error: 'Token ungültig' });
        req.user = user;
        next();
    });
};

app.post('/api/login', async (req, res) => {
    console.log(`[LOGIN] Attempt for user: ${req.body.username}`);
    const { username, password } = req.body;
    
    // Format username for AD (often needs domain\user or user@domain)
    // Adjust 'local.lab' to match your actual AD domain name
    const userPrincipalName = (username.includes('@') || username.includes('\\')) 
        ? username 
        : `${username}@${AD_DOMAIN}`;

    const client = new Client({
        url: adConfig.url,
        tlsOptions: { rejectUnauthorized: false }
    });

    try {
        console.log(`[LOGIN] Binding to AD at ${adConfig.url} as ${userPrincipalName}...`);
        await client.bind(userPrincipalName, password);
        
        // Check group membership
        let sAMAccountName = username;
        if (username.includes('\\')) sAMAccountName = username.split('\\')[1];
        else if (username.includes('@')) sAMAccountName = username.split('@')[0];

        console.log(`[LOGIN] Searching for user groups (sAMAccountName=${sAMAccountName})...`);
        const { searchEntries } = await client.search(adConfig.baseDN, {
            scope: 'sub',
            filter: `(&(objectClass=user)(sAMAccountName=${sAMAccountName}))`,
            attributes: ['memberOf']
        });
        
        console.log(`[LOGIN] Search result count: ${searchEntries.length}`);

        await client.unbind();
        
        const groups = (searchEntries[0] && searchEntries[0].memberOf) || [];
        const requiredGroup = AD_REQUIRED_GROUP;
        const readonlyGroup = AD_READONLY_GROUP;
        const maintenanceGroup = AD_MAINTENANCE_GROUP;
        
        console.log(`[LOGIN] User groups found:`, groups);

        const isMember = Array.isArray(groups) 
            ? groups.some(g => g.toUpperCase() === requiredGroup.toUpperCase())
            : (typeof groups === 'string' && groups.toUpperCase() === requiredGroup.toUpperCase());

        const isReadonly = Array.isArray(groups) 
            ? groups.some(g => g.toUpperCase() === readonlyGroup.toUpperCase())
            : (typeof groups === 'string' && groups.toUpperCase() === readonlyGroup.toUpperCase());

        const isMaintenance = Array.isArray(groups) 
            ? groups.some(g => g.toUpperCase().includes(maintenanceGroup.toUpperCase()))
            : (typeof groups === 'string' && groups.toUpperCase().includes(maintenanceGroup.toUpperCase()));

        console.log(`[LOGIN] Permissions: Admin=${isMember}, ReadOnly=${isReadonly}, Maintenance=${isMaintenance}`);

        if (!isMember && !isReadonly && !isMaintenance) {
            return res.status(403).json({ message: 'Zugriff verweigert: Sie sind nicht Mitglied der berechtigten Gruppen.' });
        }

        let role = 'readonly';
        if (isMember) role = 'admin';
        else if (isMaintenance) role = 'maintenance';

        const token = jwt.sign({ username: username, role: role }, JWT_SECRET, { expiresIn: '8h' });
        console.log(`[LOGIN] Success. Role: ${role}`);
        return res.json({ token, user: username, role: role });
    } catch (err) {
        console.error('[LOGIN] AD Auth Error:', err);
        if (err.message && (err.message.includes('integrity checking') || err.message.includes('00002028'))) {
            console.error('!!! WICHTIG: Der AD-Server verlangt LDAP-Signing. Bitte ändern Sie die AD_URL in der .env Datei auf "ldaps://..." (Port 636). !!!');
        }
        try { await client.unbind(); } catch (e) {}
        return res.status(401).json({ message: 'Authentifizierung fehlgeschlagen' });
    }
});

const TABLE_NAME = 'PAM_Governance';
const HISTORY_TABLE_NAME = 'PAM_History';
const ONBOARDING_TABLE_NAME = 'PAM_Onboarding';
const TECHNICAL_TABLE_NAME = 'PAM_TechnicalStructure';
const SECRETS_TABLE_NAME = 'PAM_SecretsManagement';
const SECRETS_ONBOARDING_TABLE_NAME = 'PAM_SecretsOnboarding';
const COLUMNS = [
  "ICTO", "Name", "Objektstatus", "Stereotyp", "tAV", "Kritikalität", 
  "Anmerkung zur Variante", "Anbindungsvariante", "PAM-Relevanz",
  "Protokollierung privilegierter Rechte auf Anwendungsebene",
  "Protokollierung priviligierter Rechte für DB auf Server",
  "Protokollierung priviligierter Rechte für Betriebssystem auf Server",
  "Schnittstellendokument",
  "Workorder Abnahme", "Entzug privilegierte Berechtigungen",
  "Kurzname", "Stellvertreter tAV", "fAV", "Betriebsverantwortlicher", "Objektpflege",
  "AbnahmePAMOnboarding"
];

let pool;

async function getPool() {
    if (!pool) {
        try {
            console.log("[DB] Connecting to SQL Server...");
            pool = await sql.connect(config);
            console.log("[DB] Connected. Checking tables...");
            await ensureTableExists(pool);
        } catch (err) {
            console.error("[DB] Connection Failed:", err);
            throw err;
        }
    }
    return pool;
}

async function ensureTableExists(p) {
    try {
        console.log("[DB] Running CREATE TABLE scripts if needed...");
        const checkTable = await p.request().query(`
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='${TABLE_NAME}' AND xtype='U')
            BEGIN
                CREATE TABLE [${TABLE_NAME}] (
                    id INT IDENTITY(1,1) PRIMARY KEY,
                    ${COLUMNS.map(col => `[${col}] NVARCHAR(MAX)`).join(',\n                    ')}
                )
            END

            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='${HISTORY_TABLE_NAME}' AND xtype='U')
            BEGIN
                CREATE TABLE [${HISTORY_TABLE_NAME}] (
                    id INT IDENTITY(1,1) PRIMARY KEY,
                    record_id INT,
                    username NVARCHAR(255),
                    action NVARCHAR(50),
                    timestamp DATETIME DEFAULT GETDATE(),
                    details NVARCHAR(MAX)
                )
            END

            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='${ONBOARDING_TABLE_NAME}' AND xtype='U')
            BEGIN
                CREATE TABLE [${ONBOARDING_TABLE_NAME}] (
                    id INT IDENTITY(1,1) PRIMARY KEY,
                    governance_id INT,
                    data NVARCHAR(MAX),
                    updated_at DATETIME DEFAULT GETDATE(),
                    updated_by NVARCHAR(255)
                )
            END

            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='${TECHNICAL_TABLE_NAME}' AND xtype='U')
            BEGIN
                CREATE TABLE [${TECHNICAL_TABLE_NAME}] (
                    id INT IDENTITY(1,1) PRIMARY KEY,
                    governance_id INT,
                    data NVARCHAR(MAX),
                    updated_at DATETIME DEFAULT GETDATE(),
                    updated_by NVARCHAR(255)
                )
            END

            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='${SECRETS_TABLE_NAME}' AND xtype='U')
            BEGIN
                CREATE TABLE [${SECRETS_TABLE_NAME}] (
                    id INT IDENTITY(1,1) PRIMARY KEY,
                    governance_id INT,
                    data NVARCHAR(MAX),
                    updated_at DATETIME DEFAULT GETDATE(),
                    updated_by NVARCHAR(255)
                )
            END

            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='${SECRETS_ONBOARDING_TABLE_NAME}' AND xtype='U')
            BEGIN
                CREATE TABLE [${SECRETS_ONBOARDING_TABLE_NAME}] (
                    id INT IDENTITY(1,1) PRIMARY KEY,
                    governance_id INT,
                    data NVARCHAR(MAX),
                    updated_at DATETIME DEFAULT GETDATE(),
                    updated_by NVARCHAR(255)
                )
            END
        `);

        console.log("[DB] Checking for missing columns...");
        // Spalten nachträglich hinzufügen, falls sie fehlen (Migration für PAM_Governance)
        for (const col of COLUMNS) {
            // console.log(`[DB] Checking column: ${col}`);
            await p.request().query(`
                IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'${col}' AND Object_ID = Object_ID(N'${TABLE_NAME}'))
                BEGIN
                    ALTER TABLE [${TABLE_NAME}] ADD [${col}] NVARCHAR(MAX)
                END
            `);
        }

        console.log(`[DB] Tables and Columns are ready.`);
    } catch (err) {
        console.error("[DB] Error creating/updating tables:", err);
    }
}

app.get('/api/health', async (req, res) => {
    try {
        const p = await getPool();
        await p.request().query('SELECT 1');
        res.json({ status: 'ok', database: 'connected' });
    } catch (err) {
        console.error("[HEALTH] Check failed:", err.message);
        res.status(200).json({ status: 'api_ok', database: 'error', message: err.message });
    }
});

let adSearchClient = null;

async function getAdClient() {
    if (adSearchClient) return adSearchClient;

    const client = new Client({
        url: adConfig.url,
        tlsOptions: { rejectUnauthorized: false },
        connectTimeout: 5000
    });

    const serviceUserUpn = (AD_SERVICE_USER.includes('@') || AD_SERVICE_USER.includes('\\'))
        ? AD_SERVICE_USER
        : `${AD_SERVICE_USER}@${AD_DOMAIN}`;

    console.log(`[AD] Establishing persistent connection...`);
    await client.bind(serviceUserUpn, AD_SERVICE_PASS);
    
    adSearchClient = client;
    return adSearchClient;
}

app.get('/api/directory/search', authenticateToken, async (req, res) => {
    const { q } = req.query;
    if (!q || String(q).length < 3) return res.json([]);

    try {
        let client;
        try {
            client = await getAdClient();
        } catch (e) {
            console.error("[AD] Connection failed:", e);
            return res.json([]);
        }
        
        // Optimierung: Prefix-Suche (Nutzt AD-Indizes) statt Wildcard am Anfang
        const filter = `(&(objectClass=user)(|(sAMAccountName=${q}*)(displayName=${q}*)(mail=${q}*)))`;
        
        const { searchEntries } = await client.search(adConfig.baseDN, {
            scope: 'sub',
            filter,
            attributes: ['sAMAccountName', 'displayName', 'mail', 'userPrincipalName'],
            sizeLimit: 20 // Begrenzung der Ergebnisse für Performance
        });
        
        const users = searchEntries.map(u => ({
            username: u.sAMAccountName,
            displayName: u.displayName,
            email: u.mail,
            upn: u.userPrincipalName
        }));

        res.json(users);
    } catch (err) {
        console.error('[AD SEARCH] Error:', err.message);
        // Force reconnect next time
        adSearchClient = null;
        res.json([]); 
    }
});

app.get('/api/onboarding/:id', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const { id } = req.params;
        const result = await p.request().input('id', sql.Int, id).query(`SELECT * FROM [${ONBOARDING_TABLE_NAME}] WHERE governance_id = @id`);
        res.json(result.recordset[0] || {});
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/onboarding/:id', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const { id } = req.params;
        const data = req.body;
        const jsonData = JSON.stringify(data);
        const user = req.user.username;

        // Check if exists
        const check = await p.request().input('id', sql.Int, id).query(`SELECT id, data FROM [${ONBOARDING_TABLE_NAME}] WHERE governance_id = @id`);
        
        if (check.recordset.length > 0) {
            // Calculate Diff
            const oldData = check.recordset[0].data ? JSON.parse(check.recordset[0].data) : {};
            const changes = [];
            const allKeys = new Set([...Object.keys(oldData), ...Object.keys(data)]);
            allKeys.forEach(k => {
                const oldVal = oldData[k] !== undefined && oldData[k] !== null ? String(oldData[k]) : '';
                const newVal = data[k] !== undefined && data[k] !== null ? String(data[k]) : '';
                if (oldVal !== newVal) {
                    changes.push({ field: k, old: oldVal, new: newVal });
                }
            });

            await p.request()
                .input('id', sql.Int, id)
                .input('data', sql.NVarChar(sql.MAX), jsonData)
                .input('user', sql.NVarChar, user)
                .query(`UPDATE [${ONBOARDING_TABLE_NAME}] SET data = @data, updated_at = GETDATE(), updated_by = @user WHERE governance_id = @id`);

            if (changes.length > 0) {
                const historyReq = p.request();
                historyReq.input('rid', sql.Int, id).input('user', sql.NVarChar, user).input('action', sql.NVarChar, 'ONBOARDING_UPDATE').input('details', sql.NVarChar, JSON.stringify(changes));
                await historyReq.query(`INSERT INTO [${HISTORY_TABLE_NAME}] (record_id, username, action, details) VALUES (@rid, @user, @action, @details)`);
            }
        } else {
            await p.request()
                .input('id', sql.Int, id)
                .input('data', sql.NVarChar(sql.MAX), jsonData)
                .input('user', sql.NVarChar, user)
                .query(`INSERT INTO [${ONBOARDING_TABLE_NAME}] (governance_id, data, updated_by) VALUES (@id, @data, @user)`);
            
            // Log creation
            const historyReq = p.request();
            historyReq.input('rid', sql.Int, id).input('user', sql.NVarChar, user).input('action', sql.NVarChar, 'ONBOARDING_ERSTELLT').input('details', sql.NVarChar, '[]');
            await historyReq.query(`INSERT INTO [${HISTORY_TABLE_NAME}] (record_id, username, action, details) VALUES (@rid, @user, @action, @details)`);
        }

        res.json({ message: 'Onboarding gespeichert' });
    } catch (err) {
        console.error("Onboarding Error:", err.message);
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/technical/:id', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const { id } = req.params;
        const result = await p.request().input('id', sql.Int, id).query(`SELECT * FROM [${TECHNICAL_TABLE_NAME}] WHERE governance_id = @id`);
        res.json(result.recordset[0] || {});
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/technical/:id', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const { id } = req.params;
        const data = req.body;
        const jsonData = JSON.stringify(data);
        const user = req.user.username;

        // Check if exists
        const check = await p.request().input('id', sql.Int, id).query(`SELECT id, data FROM [${TECHNICAL_TABLE_NAME}] WHERE governance_id = @id`);
        
        if (check.recordset.length > 0) {
            // Calculate Diff (Section level)
            const oldData = check.recordset[0].data ? JSON.parse(check.recordset[0].data) : {};
            const changes = [];
            const sections = ['servers', 'databases', 'ports', 'safes', 'safeMembers', 'sharedAccounts', 'permissions', 'mapping'];
            
            sections.forEach(k => {
                if (JSON.stringify(oldData[k]) !== JSON.stringify(data[k])) {
                    const oldCount = Array.isArray(oldData[k]) ? oldData[k].length : 0;
                    const newCount = Array.isArray(data[k]) ? data[k].length : 0;
                    changes.push({ field: k, old: `${oldCount} Einträge`, new: `${newCount} Einträge (Geändert)` });
                }
            });

            await p.request()
                .input('id', sql.Int, id)
                .input('data', sql.NVarChar(sql.MAX), jsonData)
                .input('user', sql.NVarChar, user)
                .query(`UPDATE [${TECHNICAL_TABLE_NAME}] SET data = @data, updated_at = GETDATE(), updated_by = @user WHERE governance_id = @id`);
            
            if (changes.length > 0) {
                const historyReq = p.request();
                historyReq.input('rid', sql.Int, id).input('user', sql.NVarChar, user).input('action', sql.NVarChar, 'TECHNICAL_UPDATE').input('details', sql.NVarChar, JSON.stringify(changes));
                await historyReq.query(`INSERT INTO [${HISTORY_TABLE_NAME}] (record_id, username, action, details) VALUES (@rid, @user, @action, @details)`);
            }
        } else {
            await p.request()
                .input('id', sql.Int, id)
                .input('data', sql.NVarChar(sql.MAX), jsonData)
                .input('user', sql.NVarChar, user)
                .query(`INSERT INTO [${TECHNICAL_TABLE_NAME}] (governance_id, data, updated_by) VALUES (@id, @data, @user)`);

            // Log creation
            const historyReq = p.request();
            historyReq.input('rid', sql.Int, id).input('user', sql.NVarChar, user).input('action', sql.NVarChar, 'TECHNICAL_ERSTELLT').input('details', sql.NVarChar, '[]');
            await historyReq.query(`INSERT INTO [${HISTORY_TABLE_NAME}] (record_id, username, action, details) VALUES (@rid, @user, @action, @details)`);
        }

        res.json({ message: 'Technische Struktur gespeichert' });
    } catch (err) {
        console.error("Technical Structure Error:", err.message);
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/secrets/:id', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const { id } = req.params;
        const result = await p.request().input('id', sql.Int, id).query(`SELECT * FROM [${SECRETS_TABLE_NAME}] WHERE governance_id = @id`);
        res.json(result.recordset[0] || {});
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/secrets/:id', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const { id } = req.params;
        const data = req.body;
        const jsonData = JSON.stringify(data);
        const user = req.user.username;

        // Check if exists
        const check = await p.request().input('id', sql.Int, id).query(`SELECT id, data FROM [${SECRETS_TABLE_NAME}] WHERE governance_id = @id`);
        
        if (check.recordset.length > 0) {
            // Calculate Diff (Section level)
            const oldData = check.recordset[0].data ? JSON.parse(check.recordset[0].data) : {};
            const changes = [];
            const sections = ['inventory', 'safes', 'members', 'mapping'];
            
            sections.forEach(k => {
                if (JSON.stringify(oldData[k]) !== JSON.stringify(data[k])) {
                    const oldCount = Array.isArray(oldData[k]) ? oldData[k].length : 0;
                    const newCount = Array.isArray(data[k]) ? data[k].length : 0;
                    changes.push({ field: k, old: `${oldCount} Einträge`, new: `${newCount} Einträge (Geändert)` });
                }
            });

            await p.request()
                .input('id', sql.Int, id)
                .input('data', sql.NVarChar(sql.MAX), jsonData)
                .input('user', sql.NVarChar, user)
                .query(`UPDATE [${SECRETS_TABLE_NAME}] SET data = @data, updated_at = GETDATE(), updated_by = @user WHERE governance_id = @id`);
            
            if (changes.length > 0) {
                const historyReq = p.request();
                historyReq.input('rid', sql.Int, id).input('user', sql.NVarChar, user).input('action', sql.NVarChar, 'SECRETS_UPDATE').input('details', sql.NVarChar, JSON.stringify(changes));
                await historyReq.query(`INSERT INTO [${HISTORY_TABLE_NAME}] (record_id, username, action, details) VALUES (@rid, @user, @action, @details)`);
            }
        } else {
            await p.request()
                .input('id', sql.Int, id)
                .input('data', sql.NVarChar(sql.MAX), jsonData)
                .input('user', sql.NVarChar, user)
                .query(`INSERT INTO [${SECRETS_TABLE_NAME}] (governance_id, data, updated_by) VALUES (@id, @data, @user)`);

            // Log creation
            const historyReq = p.request();
            historyReq.input('rid', sql.Int, id).input('user', sql.NVarChar, user).input('action', sql.NVarChar, 'SECRETS_ERSTELLT').input('details', sql.NVarChar, '[]');
            await historyReq.query(`INSERT INTO [${HISTORY_TABLE_NAME}] (record_id, username, action, details) VALUES (@rid, @user, @action, @details)`);
        }

        res.json({ message: 'Secrets Management gespeichert' });
    } catch (err) {
        console.error("Secrets Management Error:", err.message);
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/secrets-onboarding/:id', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const { id } = req.params;
        const result = await p.request().input('id', sql.Int, id).query(`SELECT * FROM [${SECRETS_ONBOARDING_TABLE_NAME}] WHERE governance_id = @id`);
        res.json(result.recordset[0] || {});
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/secrets-onboarding/:id', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const { id } = req.params;
        const data = req.body;
        const jsonData = JSON.stringify(data);
        const user = req.user.username;

        // Check if exists
        const check = await p.request().input('id', sql.Int, id).query(`SELECT id, data FROM [${SECRETS_ONBOARDING_TABLE_NAME}] WHERE governance_id = @id`);
        
        if (check.recordset.length > 0) {
            // Calculate Diff
            const oldData = check.recordset[0].data ? JSON.parse(check.recordset[0].data) : {};
            const changes = [];
            const allKeys = new Set([...Object.keys(oldData), ...Object.keys(data)]);
            allKeys.forEach(k => {
                const oldVal = oldData[k] !== undefined && oldData[k] !== null ? String(oldData[k]) : '';
                const newVal = data[k] !== undefined && data[k] !== null ? String(data[k]) : '';
                if (oldVal !== newVal) {
                    changes.push({ field: k, old: oldVal, new: newVal });
                }
            });

            await p.request().input('id', sql.Int, id).input('data', sql.NVarChar(sql.MAX), jsonData).input('user', sql.NVarChar, user).query(`UPDATE [${SECRETS_ONBOARDING_TABLE_NAME}] SET data = @data, updated_at = GETDATE(), updated_by = @user WHERE governance_id = @id`);
            
            if (changes.length > 0) {
                const historyReq = p.request();
                historyReq.input('rid', sql.Int, id).input('user', sql.NVarChar, user).input('action', sql.NVarChar, 'SECRETS_ONB_UPDATE').input('details', sql.NVarChar, JSON.stringify(changes));
                await historyReq.query(`INSERT INTO [${HISTORY_TABLE_NAME}] (record_id, username, action, details) VALUES (@rid, @user, @action, @details)`);
            }
        } else {
            await p.request().input('id', sql.Int, id).input('data', sql.NVarChar(sql.MAX), jsonData).input('user', sql.NVarChar, user).query(`INSERT INTO [${SECRETS_ONBOARDING_TABLE_NAME}] (governance_id, data, updated_by) VALUES (@id, @data, @user)`);
            const historyReq = p.request();
            historyReq.input('rid', sql.Int, id).input('user', sql.NVarChar, user).input('action', sql.NVarChar, 'SECRETS_ONB_ERSTELLT').input('details', sql.NVarChar, '[]');
            await historyReq.query(`INSERT INTO [${HISTORY_TABLE_NAME}] (record_id, username, action, details) VALUES (@rid, @user, @action, @details)`);
        }
        res.json({ message: 'Secrets Onboarding gespeichert' });
    } catch (err) {
        console.error("Secrets Onboarding Error:", err.message);
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/data', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        let query = `SELECT * FROM [${TABLE_NAME}]`;
        
        if (req.user.role === 'maintenance') {
            query += ` WHERE [Objektpflege] LIKE '%(${req.user.username})%'`;
        }
        
        query += ` ORDER BY id DESC`;
        let result = await p.request().query(query);
        res.json(result.recordset);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/history/:id', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const { id } = req.params;
        const result = await p.request().input('id', sql.Int, id).query(`SELECT * FROM [${HISTORY_TABLE_NAME}] WHERE record_id = @id ORDER BY timestamp DESC`);
        res.json(result.recordset);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/data', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const data = req.body;
        // Filtere ungültige Keys und die ID
        const validKeys = Object.keys(data).filter(k => COLUMNS.includes(k));
        
        const columns = validKeys.map(k => `[${k}]`).join(', ');
        const values = validKeys.map(k => `@${k.replace(/[\s-]/g, '_')}`).join(', ');

        const request = p.request();
        validKeys.forEach(k => {
            request.input(k.replace(/[\s-]/g, '_'), sql.NVarChar, data[k] || '');
        });

        // Insert and get ID
        const query = `INSERT INTO [${TABLE_NAME}] (${columns}) VALUES (${values}); SELECT SCOPE_IDENTITY() AS id`;
        const result = await request.query(query);
        const newId = result.recordset[0].id;

        // Log History
        const historyReq = p.request();
        historyReq.input('rid', sql.Int, newId).input('user', sql.NVarChar, req.user.username).input('action', sql.NVarChar, 'ERSTELLT').input('details', sql.NVarChar, JSON.stringify(data));
        await historyReq.query(`INSERT INTO [${HISTORY_TABLE_NAME}] (record_id, username, action, details) VALUES (@rid, @user, @action, @details)`);

        res.status(201).send({ message: "Erfolgreich erstellt" });
    } catch (err) {
        console.error("POST Error:", err.message);
        res.status(500).json({ error: err.message });
    }
});

app.put('/api/data/:id', authenticateToken, async (req, res) => {
    try {
        const p = await getPool();
        const { id } = req.params;
        const data = req.body;
        
        // Fetch current record for diff
        const currentResult = await p.request()
            .input('id_lookup', sql.Int, id)
            .query(`SELECT * FROM [${TABLE_NAME}] WHERE id = @id_lookup`);
        const currentRecord = currentResult.recordset[0];

        const validKeys = Object.keys(data).filter(k => COLUMNS.includes(k));

        // Calculate changes
        const changes = [];
        if (currentRecord) {
            validKeys.forEach(k => {
                const oldVal = currentRecord[k] != null ? String(currentRecord[k]) : '';
                const newVal = data[k] != null ? String(data[k]) : '';
                if (oldVal !== newVal) {
                    changes.push({ field: k, old: oldVal, new: newVal });
                }
            });
        }

        const setClause = validKeys.map(k => `[${k}] = @${k.replace(/[\s-]/g, '_')}`).join(', ');

        const request = p.request();
        request.input('id', sql.Int, id);
        validKeys.forEach(k => {
            request.input(k.replace(/[\s-]/g, '_'), sql.NVarChar, data[k] || '');
        });

        const query = `UPDATE [${TABLE_NAME}] SET ${setClause} WHERE id = @id`;
        await request.query(query);

        // Log History
        if (changes.length > 0) {
            const historyReq = p.request();
            historyReq.input('rid', sql.Int, id).input('user', sql.NVarChar, req.user.username).input('action', sql.NVarChar, 'BEARBEITET').input('details', sql.NVarChar, JSON.stringify(changes));
            await historyReq.query(`INSERT INTO [${HISTORY_TABLE_NAME}] (record_id, username, action, details) VALUES (@rid, @user, @action, @details)`);
        }

        res.send({ message: "Erfolgreich aktualisiert" });
    } catch (err) {
        console.error("PUT Error:", err.message);
        res.status(500).json({ error: err.message });
    }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Backend läuft auf http://127.0.0.1:${PORT}`);
});
