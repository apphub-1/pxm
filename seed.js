import 'dotenv/config';
import sql from 'mssql';

const dbPort = process.env.DB_PORT ? parseInt(process.env.DB_PORT) : undefined;
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
        connectTimeout: 60000, // Höheres Timeout für Massenoperationen
        useUTC: false
    }
};

const GOVERNANCE_TABLE = 'PAM_Governance';
const ONBOARDING_TABLE = 'PAM_Onboarding';
const TECHNICAL_TABLE = 'PAM_TechnicalStructure';
const SECRETS_TABLE = 'PAM_SecretsManagement';
const SECRETS_ONBOARDING_TABLE = 'PAM_SecretsOnboarding';

// Hilfsfunktionen für Zufallsdaten
const pick = (arr) => arr[Math.floor(Math.random() * arr.length)];
const bool = () => Math.random() > 0.5;
const int = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;

const generateGovernance = (i) => ({
    ICTO: `ICTO-${10000 + i}`,
    Name: `Applikation ${i} - ${pick(['Core', 'Support', 'Legacy', 'Innovation'])} System`,
    Kurzname: `APP-${i}`,
    Objektstatus: pick(["Plan", "Aktiv", "Stillgelegt"]),
    Stereotyp: pick(["ExternalSystem_ASPSAAS", "InternalSystem", "CloudService"]),
    tAV: `Max Mustermann ${int(1, 20)} (u123${i})`,
    "Stellvertreter tAV": `Erika Musterfrau ${int(1, 20)} (u234${i})`,
    fAV: `Fachlicher User ${int(1, 20)}`,
    Betriebsverantwortlicher: `Ops Manager ${int(1, 20)}`,
    Kritikalität: pick(["1-niedrig", "2-mittel", "3-hoch", "4-kritisch"]),
    "Anmerkung zur Variante": "Automatisch generiert durch Seed-Skript",
    Anbindungsvariante: pick(["Variante 1", "Variante 2", "Variante 3", "Sonderlösung"]),
    "PAM-Relevanz": pick(["Ja", "Nein"]),
    "Protokollierung privilegierter Rechte auf Anwendungsebene": pick([
        "Keine PAM-Anbindung - keine privilegierten Rechte",
        "Protokollierung privilegierter Rechte über CyberArk"
    ]),
    "Protokollierung priviligierter Rechte für DB auf Server": pick([
        "keine PAM-Anbindung für den eingesetzten Datenbanktyp",
        "Protokollierung privilegierter Rechte über CyberArk"
    ]),
    "Protokollierung priviligierter Rechte für Betriebssystem auf Server": pick([
        "keine PAM-Anbindung für den eingesetzten Betriebssystemtyp",
        "Protokollierung privilegierter Rechte über CyberArk"
    ]),
    Schnittstellendokument: pick(["Vorhanden", "Nicht vorhanden", "In Arbeit"]),
    "Workorder Abnahme": `WO-${int(100000, 999999)}`,
    "Entzug privilegierte Berechtigungen": pick(["erfolgt", "in Arbeit", "offen"])
});

const generateOnboarding = (gov) => ({
    appName: gov.Name,
    shortName: gov.Kurzname,
    ictoId: gov.ICTO,
    criticality: gov.Kritikalität,
    tav: gov.tAV,
    archDesc: "Standard 3-Tier Architektur mit Web, App und DB Layer. Loadbalancer vorgeschaltet.",
    osList: "RHEL 8.6, Windows Server 2019",
    dbList: "Oracle 19c, MSSQL 2019",
    serverReplace: bool(),
    serverReplaceDate: "2025-12-31",
    hasLoadBalancer: bool(),
    techType: pick(["Fat Client", "Webbasiert", "Java", "Mischform", "Cloud-Service"]),
    needsCerts: bool(),
    fourEyes: bool(),
    licenseType: pick(["Accountgebunden", "Identitätsgebunden", "Serverbasiert"]),
    urlProd: `https://app${int(1,300)}.prod.company.local`,
    urlTest: `https://app${int(1,300)}.test.company.local`,
    loginMethod: pick(["Single Sign-On", "Anmeldemaske", "Multifaktor-Authentifizierung"]),
    accountType: pick(["Lokale Accounts", "Domain Accounts"]),
    multiSession: bool(),
    usernameNaming: "u_app_xyz",
    pwChangeDesc: "Manuell durch Admin im Wartungsfenster",
    autoRotation: bool(),
    whoChangesPw: "Betriebsteam",
    pwInterval: "90 Tage",
    hasTestUsers: bool(),
    testUsersList: "test_user_1, test_user_2",
    omadaRights: bool(),
    hasEmergencyAccounts: bool(),
    emergencyAccountsList: "breakglass_admin",
    matrixLogin: pick(["Automatisch", "Manuell"]),
    matrixPwChange: pick(["Automatisch", "Manuell"]),
    selectedVariant: gov.Anbindungsvariante
});

const generateTechnical = () => ({
    servers: Array.from({ length: int(1, 5) }, (_, i) => ({
        serverName: `srv-app-${int(100, 999)}-0${i+1}`,
        ip: `10.20.${int(1, 255)}.${int(1, 255)}`,
        fqdn: `srv-app-${int(100, 999)}-0${i+1}.company.local`,
        stage: pick(['Prod', 'Test', 'Dev']),
        dmz: bool(),
        desc: "App Server",
        os: "RHEL 8",
        port: "22",
        expiry: "2026-01-01"
    })),
    databases: Array.from({ length: int(1, 3) }, (_, i) => ({
        serverName: `db-app-${int(100, 999)}-0${i+1}`,
        ip: `10.30.${int(1, 255)}.${int(1, 255)}`,
        fqdn: `db-app-${int(100, 999)}-0${i+1}.company.local`,
        stage: 'Prod',
        dbType: "Oracle",
        instance: "ORCL",
        product: "Oracle 19c",
        port: "1521"
    })),
    ports: [],
    safes: Array.from({ length: int(1, 2) }, (_, i) => ({
        userGroup: "PXM-Admins",
        safeName: `Safe ${i+1}`,
        safeDesc: "Hauptsafe für Applikation",
        techSafeName: `PXM_SAFE_${i+1}_PROD`,
        adGroup: "g_pxm_safe_access",
        adGroupDesc: "Zugriffsgruppe AD",
        approver: "Manager",
        sod: "Keine"
    })),
    safeMembers: [],
    sharedAccounts: Array.from({ length: int(2, 5) }, (_, i) => ({
        bizName: `TechUser ${i+1}`,
        techName: `svc_app_${i+1}`,
        sam: `svc_app_${i+1}`,
        login: `svc_app_${i+1}`,
        desc: "Service Account für Backend",
        isAd: true,
        owner: "Team Lead",
        ownerId: "u12345"
    })),
    permissions: [],
    mapping: []
});

const generateSecrets = () => ({
    inventory: Array.from({ length: int(3, 8) }, (_, i) => ({
        category: pick(['Passwort', 'SSH-Key', 'API-Key', 'Zertifikat']),
        name: `Secret_${i+1}`,
        owner: "App Owner",
        holder: "DevOps Team",
        layer: pick(['Anwendung', 'Betriebssystem', 'Datenbank']),
        localOrAd: pick(['Lokal', 'AD']),
        stage: pick(['Prod', 'Test', 'Dev', 'Int']),
        complexity: "Hoch (20 Zeichen, Sonderzeichen)",
        autoRotation: pick(['Ja', 'Nein']),
        rotationMech: "CPM",
        frequency: "Monatlich",
        timeWindow: "Wartungsfenster Sonntag"
    })),
    safes: [],
    members: [],
    mapping: []
});

const generateSecretsOnboarding = () => ({
    accessSource: "FCE",
    targetArchitecture: "Cloud Native Migration geplant",
    operatingModel: pick(["On-Prem", "Cloud", "Hybrid"]),
    appType: pick(["Statischer Server", "Container"]),
    os: "Linux / Windows",
    deployment: "GitLab CI/CD Pipeline",
    initialAuth: "Manuell via Ticket",
    containerSelfBuilt: "Ja",
    k8sSecrets: "Nein",
    codeControl: "Ja",
    serverless: "Nein",
    cloudVMs: "Ja",
    codeControlServer: "Ja",
    configControlServer: "Ja",
    centralConfig: "Ja",
    javaWebserver: "Nein",
    secretTypes: ["Passwörter", "API-Keys"],
    secretStorage: ["Config-Files", "ENV-Variablen"],
    manualProcesses: "Nein",
    rotationResponsibility: "Ja",
    rotationLevel: ["Datenbankebene"],
    currentRotation: "Manuell",
    targetTool: "CyberArk CCP",
    targetVariant: "Vollautomatisiert zur Laufzeit",
    targetRotationMech: "Automatisch",
    targetTimeWindow: "Untertägig",
    targetFrequency: "Täglich"
});

async function seed() {
    try {
        const pool = await sql.connect(config);
        console.log('Verbunden mit Datenbank. Starte Seeding für 300 Anwendungen...');

        for (let i = 1; i <= 300; i++) {
            const govData = generateGovernance(i);
            
            // Insert Governance
            const govKeys = Object.keys(govData);
            const govCols = govKeys.map(k => `[${k}]`).join(', ');
            const govVals = govKeys.map(k => `@${k.replace(/[\s-]/g, '_')}`).join(', ');
            
            const req = pool.request();
            govKeys.forEach(k => req.input(k.replace(/[\s-]/g, '_'), sql.NVarChar, govData[k]));
            
            const res = await req.query(`INSERT INTO [${GOVERNANCE_TABLE}] (${govCols}) VALUES (${govVals}); SELECT SCOPE_IDENTITY() AS id`);
            const id = res.recordset[0].id;

            // Insert Onboarding
            const onbData = JSON.stringify(generateOnboarding(govData));
            await pool.request()
                .input('id', sql.Int, id)
                .input('data', sql.NVarChar(sql.MAX), onbData)
                .input('user', sql.NVarChar, 'SeedScript')
                .query(`INSERT INTO [${ONBOARDING_TABLE}] (governance_id, data, updated_by) VALUES (@id, @data, @user)`);

            // Insert Technical
            const techData = JSON.stringify(generateTechnical());
            await pool.request()
                .input('id', sql.Int, id)
                .input('data', sql.NVarChar(sql.MAX), techData)
                .input('user', sql.NVarChar, 'SeedScript')
                .query(`INSERT INTO [${TECHNICAL_TABLE}] (governance_id, data, updated_by) VALUES (@id, @data, @user)`);

            // Insert Secrets
            const secData = JSON.stringify(generateSecrets());
            await pool.request()
                .input('id', sql.Int, id)
                .input('data', sql.NVarChar(sql.MAX), secData)
                .input('user', sql.NVarChar, 'SeedScript')
                .query(`INSERT INTO [${SECRETS_TABLE}] (governance_id, data, updated_by) VALUES (@id, @data, @user)`);

            // Insert Secrets Onboarding
            const secOnbData = JSON.stringify(generateSecretsOnboarding());
            await pool.request()
                .input('id', sql.Int, id)
                .input('data', sql.NVarChar(sql.MAX), secOnbData)
                .input('user', sql.NVarChar, 'SeedScript')
                .query(`INSERT INTO [${SECRETS_ONBOARDING_TABLE}] (governance_id, data, updated_by) VALUES (@id, @data, @user)`);

            if (i % 10 === 0) process.stdout.write(`.`);
        }

        console.log('\nSeeding erfolgreich abgeschlossen!');
        await pool.close();
    } catch (err) {
        console.error('\nFehler beim Seeding:', err);
    }
}

seed();
