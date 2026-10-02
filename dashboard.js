(() => {

    "use strict";


    const SHEETS_API =
        "https://sheets.googleapis.com/v4/spreadsheets";

    const DRIVE_API =
        "https://www.googleapis.com/drive/v3/files";


    const SESSION_KEY = "fmAssetSession";
    let tokenRequestPromise = null;
    let movementSubmitBusy = false;

    const state = {

        idTokenPayload: null,

        accessToken: null,

        workbookId: null,

        workbookName: null,

        inventory: [],

        transactions: [],

        // The live inventory/transactions above remain the source for
        // movement recording and inventory management. Dashboard-only
        // views can switch to an archived snapshot without changing them.
        dashboardInventory: [],
        dashboardTransactions: [],
        dashboardClients: [],
        historicalLedgerFiles: [],
        dashboardSource: { kind: "current", id: null, name: "Current live ledger" },

        todayLoads: [],
        routines: [],
        alerts: [],
        alertHeaders: [],

        clients: [],
        auditVisibleRows: [],
        clientDetailsRows: []
    };


    const $ =
        id => document.getElementById(id);

document.addEventListener(
    "DOMContentLoaded",
    () => {

        bindEvents();

        bindQuickMenu();

        setDefaultTimestamp();

        handleMovementChange();

        waitForGoogle();

    }
);


    function bindEvents() {

        $("grant-access")
            .addEventListener(
                "click",
                requestSheetAccess
            );


        $("refresh")
            .addEventListener(
                "click",
                loadDashboard
            );


        $("sign-out")
            .addEventListener(
                "click",
                signOut
            );


        $("movement-form")
            .addEventListener(
                "submit",
                recordMovement
            );

        $("add-asset-row").addEventListener("click", () => addAssetRow());
        $("asset-rows").addEventListener("input", updateMovementPreview);
        $("asset-rows").addEventListener("change", updateMovementPreview);
        $("asset-rows").addEventListener("click", event => {
            const button = event.target.closest("[data-remove-asset]");
            if (!button) return;
            removeAssetRow(Number(button.dataset.removeAsset));
        });


        $("movement")
            .addEventListener(
                "change",
                handleMovementChange
            );

        $("movement-photo")?.addEventListener("change", handleDashboardPhotoChange);
        $("clear-photo")?.addEventListener("click", clearDashboardPhoto);
        $("take-photo")?.addEventListener("click", () => $("movement-photo-camera")?.click());
        $("movement-photo-camera")?.addEventListener("change", event => {
            const file = event.target.files?.[0];
            const main = $("movement-photo");
            if (!file || !main) return;
            try {
                const transfer = new DataTransfer();
                transfer.items.add(file);
                main.files = transfer.files;
            } catch (_) {
                // Older mobile browsers may not expose DataTransfer for file inputs.
                window._fmCameraFile = file;
            }
            handleDashboardPhotoChange();
        });


        $("client-cards")
            .addEventListener(
                "click",
                event => {
                    const card = event.target.closest("[data-client]");
                    if (card) {
                        openClientDetails(card.dataset.client);
                    }
                }
            );

        $("client-cards")
            .addEventListener(
                "keydown",
                event => {
                    if (event.key !== "Enter" && event.key !== " ") return;
                    const card = event.target.closest("[data-client]");
                    if (!card) return;
                    event.preventDefault();
                    openClientDetails(card.dataset.client);
                }
            );

        $("client-search")
            .addEventListener(
                "input",
                renderClientCards
            );

        
        const dashboardFilterPairs = [
            ["dashboard-filter-client", "dashboard-filter-client-audit"],
            ["dashboard-filter-date", "dashboard-filter-date-audit"],
            ["dashboard-filter-asset", "dashboard-filter-asset-audit"]
        ];

        dashboardFilterPairs.forEach(([mainId, auditId]) => {
            $(mainId)?.addEventListener("change", () => {
                if ($(auditId)) $(auditId).value = $(mainId).value;
                renderClientCards();
                renderWarehouseAssets();
                renderAudit();
                renderFilteredLoadTables();
            });
            $(auditId)?.addEventListener("change", () => {
                if ($(mainId)) $(mainId).value = $(auditId).value;
                renderClientCards();
                renderWarehouseAssets();
                renderAudit();
                renderFilteredLoadTables();
            });
        });

        $("clear-dashboard-filters")?.addEventListener("click", clearDashboardFilters);
        $("clear-audit-filters")?.addEventListener("click", clearDashboardFilters);

        $("dashboard-data-version")?.addEventListener("change", handleDashboardDataVersionChange);

        $("transactions-body")?.addEventListener("click", event => {
            const button = event.target.closest("[data-info-index]");
            if (button) openInfoModal(Number(button.dataset.infoIndex));
        });
        $("client-details-transactions")?.addEventListener("click", event => {
            const button = event.target.closest("[data-client-info-index]");
            if (button) openClientTransactionInfo(Number(button.dataset.clientInfoIndex));
        });
        $("close-info")?.addEventListener("click", closeInfoModal);
        document.addEventListener("click", event => {
            const summaryToggle = event.target.closest("[data-summary-toggle]");
            if (summaryToggle) {
                toggleTransactionSummary(summaryToggle);
                return;
            }
            if (event.target.matches("[data-close-info]")) closeInfoModal();
        });
        $("clear-client-details-filters")?.addEventListener("click", () => {
            const client = $("client-details-title")?.textContent || "";
            if ($("client-details-filter-date")) $("client-details-filter-date").value = "";
            if ($("client-details-filter-asset")) $("client-details-filter-asset").value = "";
            if (client) openClientDetails(client);
        });
$("client-details-summary")
            .addEventListener(
                "click",
                event => {
                    const button = event.target.closest("[data-client-asset]");
                    if (!button) return;

                    const client = $("client-details-title").textContent;
                    renderClientAssetGraph(client, button.dataset.clientAsset);

                    document
                        .querySelectorAll("[data-client-asset]")
                        .forEach(item => item.classList.remove("selected"));

                    button.classList.add("selected");
                }
            );


        $("close-client-details")
            .addEventListener(
                "click",
                closeClientDetails
            );


        document.addEventListener(
            "click",
            event => {

                if (
                    event.target.matches(
                        "[data-close-client-details]"
                    )
                ) {

                    closeClientDetails();

                }

            }
        );


        $("manage-inventory")
            .addEventListener(
                "click",
                openInventoryManager
            );


        $("close-inventory")
            .addEventListener(
                "click",
                closeInventoryManager
            );


        document.addEventListener(
            "click",
            event => {

                if (
                    event.target.matches(
                        "[data-close-modal]"
                    )
                ) {

                    closeInventoryManager();

                }

            }
        );


        $("add-asset")
            .addEventListener(
                "click",
                addAssetType
            );


        $("manage-inventory-body")
            .addEventListener(
                "click",
                event => {

                    const button =
                        event.target.closest(
                            "button[data-action]"
                        );

                    if (!button) {
                        return;
                    }


                    const rowNumber =
                        Number(button.dataset.row);


                    if (
                        button.dataset.action ===
                        "delete"
                    ) {

                        deleteAssetType(
                            rowNumber
                        );

                    }


                    if (
                        button.dataset.action ===
                        "rename"
                    ) {

                        renameAssetType(
                            rowNumber
                        );

                    }

                }
            );

    }



function bindQuickMenu() {
    const nav = document.querySelector(".quick-nav");
    const toggle = document.getElementById("menu-toggle");
    const links = document.getElementById("quick-links");

    if (!nav || !toggle || !links) {
        console.warn("Quick Links menu elements not found.");
        return;
    }

    const storageKey = "fmQuickMenuOpen";

    function setOpen(open) {
        nav.classList.toggle("menu-open", open);
        toggle.setAttribute("aria-expanded", String(open));

        try {
            localStorage.setItem(storageKey, open ? "1" : "0");
        } catch (error) {
            // Ignore storage errors
        }
    }

    // Mobile starts closed, desktop starts open unless the user has
    // previously chosen a state.
    let saved = null;

    try {
        saved = localStorage.getItem(storageKey);
    } catch (error) {
        saved = null;
    }

    const mobile = window.matchMedia("(max-width: 760px)").matches;

    if (saved === "1") {
        setOpen(true);
    } else if (saved === "0") {
        setOpen(false);
    } else {
        setOpen(!mobile);
    }

    // Open / close menu
    toggle.addEventListener("click", function (event) {
        event.preventDefault();
        event.stopPropagation();

        const isOpen = nav.classList.contains("menu-open");
        setOpen(!isOpen);
    });

    links.addEventListener("click", function (event) {
        const link = event.target.closest("a");

        if (!link) return;

        const href = link.getAttribute("href");

        if (!href) return;

        if (window.matchMedia("(max-width: 760px)").matches) {
            setOpen(false);
        }

        if (href.startsWith("#")) {
            const target = document.querySelector(href);

            if (target) {
                event.preventDefault();

                setTimeout(() => {
                    target.scrollIntoView({
                        behavior: "smooth",
                        block: "start"
                    });

                    history.replaceState(null, "", href);
                }, 0);
            }
        }
    });
}

    function waitForGoogle() {
        if (window.google?.accounts?.id && window.google?.accounts?.oauth2) {
            initializeGoogle();
            return;
        }

        let finished = false;
        const started = Date.now();

        const onLoaded = () => {
            if (finished) return;
            finished = true;
            window.removeEventListener("fm-google-loaded", onLoaded);
            initializeGoogle();
        };

        window.addEventListener("fm-google-loaded", onLoaded, { once: true });

        const timer = setInterval(() => {
            if (finished) {
                clearInterval(timer);
                return;
            }

            if (window.google?.accounts?.id && window.google?.accounts?.oauth2) {
                clearInterval(timer);
                finished = true;
                window.removeEventListener("fm-google-loaded", onLoaded);
                initializeGoogle();
                return;
            }

            if (Date.now() - started >= 8000) {
                clearInterval(timer);
                finished = true;
                window.removeEventListener("fm-google-loaded", onLoaded);
                setAuthStatus("Google services could not be loaded. Check your internet connection.", true);
            }
        }, 50);
    }


    function initializeGoogle() {
        if (!CONFIG.GOOGLE_CLIENT_ID || CONFIG.GOOGLE_CLIENT_ID.includes("PASTE_YOUR")) {
            setAuthStatus("Add your existing Google Web Client ID to config.js.", true);
            return;
        }

        google.accounts.id.initialize({
            client_id: CONFIG.GOOGLE_CLIENT_ID,
            callback: handleCredentialResponse,
            auto_select: true,
            cancel_on_tap_outside: false
        });

        google.accounts.id.renderButton($("google-signin-button"), {
            theme: "outline", size: "large", text: "signin_with", shape: "rectangular", width: 280
        });

        const saved = readSavedSession();
        if (saved) {
            setUserProfile(saved);
            $("grant-access").classList.add("hidden");
            setAuthStatus("Restoring your Sheets connection…");
            window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…", "Restoring your saved Google session.");
            attemptSilentAccess(saved.email);
        } else {
            window.FM_CONNECTION_UI?.hide();
            google.accounts.id.prompt();
        }
    }


    function handleCredentialResponse(response) {
        try {
            state.idTokenPayload = decodeJwtPayload(response.credential);
            saveSession();
            setUserProfile(state.idTokenPayload);
            $("grant-access").classList.remove("hidden");
            setAuthStatus("Signed in. Connecting to Google Sheets…");
            window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…", "Requesting access for this dashboard.");
            attemptSilentAccess(state.idTokenPayload.email);
        } catch (error) {
            console.error(error);
            setAuthStatus("Google sign-in response could not be read.", true);
        }
    }


    function requestSheetAccess(silent = false) {
        if (!state.idTokenPayload && !readSavedSession()) {
            setAuthStatus("Sign in with Google first.", true);
            return;
        }

        window.FM_CONNECTION_UI?.show(
            silent ? "Connecting to Google Sheets…" : "Authorising Google Sheets & Drive…",
            silent ? "Restoring your saved connection." : "Approve access in the Google prompt, then the dashboard will continue automatically."
        );
        acquireAccessToken(silent ? "none" : "consent")
            .then(async () => {
                hideLogin();
                await loadDashboard();
                window.FM_CONNECTION_UI?.hide();
            })
            .catch(error => {
                window.FM_CONNECTION_UI?.hide();
                $("grant-access").classList.remove("hidden");
                if (silent) {
                    setAuthStatus("Your Google account is signed in. Allow Sheets & Drive access to continue.");
                } else {
                    setAuthStatus(error.message || "Google authorization failed. Tap Allow Sheets & Drive access to try again.", true);
                }
            });
    }


    function attemptSilentAccess(email) {
        window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…", "Checking your current Google authorisation.");
        acquireAccessToken("none", email)
            .then(async () => {
                hideLogin();
                await loadDashboard();
                window.FM_CONNECTION_UI?.hide();
            })
            .catch((error) => {
                console.warn("Silent Google Sheets/Drive authorization failed:", error);
                window.FM_CONNECTION_UI?.hide();
                $("grant-access").classList.remove("hidden");
                setAuthStatus("Your Google session is available. Allow Sheets & Drive access to continue.");
            });
    }


    function acquireAccessToken(prompt = "none", email, options = {}) {
        const forceRefresh = Boolean(options.forceRefresh);
        const expectedEmail = email || state.idTokenPayload?.email || readSavedSession()?.email || "";
        if (!forceRefresh) {
            const cached = window.FM_AUTH_CACHE?.read?.(expectedEmail);
            if (cached?.token) {
                state.accessToken = cached.token;
                return Promise.resolve(cached.token);
            }
        }
        if (tokenRequestPromise) return tokenRequestPromise;
        tokenRequestPromise = new Promise((resolve, reject) => {
            let settled = false;
            const finish = (fn, value) => {
                if (settled) return;
                settled = true;
                tokenRequestPromise = null;
                clearTimeout(timeoutId);
                fn(value);
            };
            const timeoutMs = prompt === "none" ? 3000 : 15000;
            const timeoutId = setTimeout(() => finish(reject, new Error("Google authorisation did not complete.")), timeoutMs);
            const tokenClient = google.accounts.oauth2.initTokenClient({
                client_id: CONFIG.GOOGLE_CLIENT_ID,
                scope: CONFIG.OAUTH_SCOPES,
                callback: response => {
                    if (response.error) { finish(reject, new Error(`Google authorization failed: ${response.error}`)); return; }
                    state.accessToken = response.access_token;
                    window.FM_AUTH_CACHE?.write?.(response.access_token, response.expires_in, expectedEmail);
                    finish(resolve, response.access_token);
                }
            });
            tokenClient.requestAccessToken({
                prompt,
                login_hint: expectedEmail || undefined
            });
        });
        return tokenRequestPromise;
    }

    async function ensureWriteAccess() {
        const email = state.idTokenPayload?.email || readSavedSession()?.email || "";
        const cached = window.FM_AUTH_CACHE?.read?.(email);
        if (cached?.token) {
            state.accessToken = cached.token;
            return true;
        }
        try {
            window.FM_CONNECTION_UI?.show("Checking Google access…", "Refreshing access before saving this movement.");
            await acquireAccessToken("none", email, {forceRefresh: true});
            window.FM_CONNECTION_UI?.hide();
            return true;
        } catch (silentError) {
            console.warn("Silent token refresh failed before movement save:", silentError);
            try {
                window.FM_CONNECTION_UI?.show("Google permission needed", "Approve the Google access prompt to finish saving this movement.");
                await acquireAccessToken("consent", email, {forceRefresh: true});
                window.FM_CONNECTION_UI?.hide();
                return true;
            } catch (consentError) {
                window.FM_CONNECTION_UI?.hide();
                throw consentError;
            }
        }
    }

    async function loadDashboard() {

        if (!state.accessToken) {
            return;
        }

        setSyncStatus("Syncing with Assets Inventory Ledger...");

        try {
            const ledgerId = CONFIG.INVENTORY_LEDGER_SHEET_ID;
            const metadata = await sheetsGet(`/${encodeURIComponent(ledgerId)}`);
            let sheetTitles = (metadata.sheets || []).map(sheet => sheet.properties.title);

            const requiredSheets = [
                CONFIG.INVENTORY_SHEET_NAME,
                CONFIG.TRANSACTIONS_SHEET_NAME,
                CONFIG.ROUTINE_SHEET_NAME,
                CONFIG.ALERTS_SHEET_NAME
            ];
            const missingSheets = requiredSheets.filter(name => !sheetTitles.includes(name));

            if (missingSheets.length) {
                await createSheets(missingSheets);
                sheetTitles = sheetTitles.concat(missingSheets);
            }

            const [inventoryRows, transactionRows, clientListRows, routineRows, alertRows] = await Promise.all([
                getValues(ledgerId, CONFIG.INVENTORY_SHEET_NAME),
                getValues(ledgerId, CONFIG.TRANSACTIONS_SHEET_NAME),
                getValues(ledgerId, CONFIG.CLIENT_LIST_SHEET_NAME),
                getValues(ledgerId, CONFIG.ROUTINE_SHEET_NAME),
                getValues(ledgerId, CONFIG.ALERTS_SHEET_NAME)
            ]);

            await ensureTransactionColumns(ledgerId, transactionRows);
            await ensureOperationalSheetHeaders(ledgerId, CONFIG.ROUTINE_SHEET_NAME, routineRows, FM_ROUTINE_ALERTS.ROUTINE_HEADERS);
            await ensureOperationalSheetHeaders(ledgerId, CONFIG.ALERTS_SHEET_NAME, alertRows, FM_ROUTINE_ALERTS.ALERT_HEADERS);
            state.alertHeaders = alertRows[0]?.slice() || FM_ROUTINE_ALERTS.ALERT_HEADERS.slice();

            state.inventory = parseInventory(inventoryRows);
            state.transactions = parseTransactions(transactionRows);
            state.clients = parseClients(clientListRows);
            state.routines = FM_ROUTINE_ALERTS.parseRoutineRows(routineRows);
            state.alerts = FM_ROUTINE_ALERTS.parseAlertRows(alertRows);
            state.todayLoads = FM_ROUTINE_ALERTS.routineLoadRows(state.routines, new Date());

            state.workbookId = ledgerId;
            state.workbookName = CONFIG.INVENTORY_LEDGER_NAME || "Assets Inventory Ledger";
            $("workbook-name").textContent = state.workbookName;

            state.dashboardInventory = state.inventory.slice();
            state.dashboardTransactions = state.transactions.slice();
            state.dashboardClients = state.clients.slice();
            state.dashboardSource = { kind: "current", id: ledgerId, name: state.workbookName };

            // Refresh the alert snapshot: stale open alerts are removed;
            // attended/ignored history remains and today's candidates are rebuilt.
            try {
                await rebuildAlertSnapshot();
            } catch (alertError) {
                console.warn("Alerts could not be synchronised:", alertError);
            }

            try {
                await loadHistoricalLedgerFiles();
            } catch (archiveError) {
                console.warn("Historical Dashboard snapshots could not be listed:", archiveError);
                state.historicalLedgerFiles = [];
            }

            populateDashboardDataVersionOptions();
            renderDashboard();
            updateAlertNavBadges();

            setSyncStatus(`Synced at ${new Date().toLocaleTimeString()}`);
        } catch (error) {
            console.error(error);
            setSyncStatus(error.message || "Unable to load spreadsheet.", true);
        }
    }


    function parseHistoricalLedgerName(name) {
        const match = String(name || "").trim().match(/^(\d{1,2})_(\d{4})_(\d{4})$/);
        if (!match) return null;

        const week = Number(match[1]);
        const year = Number(match[2]);
        const hhmm = match[3];
        const hour = Number(hhmm.slice(0, 2));
        const minute = Number(hhmm.slice(2));

        if (!Number.isInteger(week) || week < 1 || week > 53 ||
            !Number.isInteger(year) || !Number.isInteger(hour) || hour > 23 ||
            !Number.isInteger(minute) || minute > 59) {
            return null;
        }

        return {
            week,
            year,
            hhmm,
            timeLabel: `${hhmm.slice(0, 2)}:${hhmm.slice(2)}`,
            label: `Week ${week} · ${year} · ${hhmm.slice(0, 2)}:${hhmm.slice(2)}`
        };
    }


    async function loadHistoricalLedgerFiles() {
        const parent = CONFIG.DRIVE_PHOTOS_PARENT_FOLDER_ID;
        if (!parent) {
            state.historicalLedgerFiles = [];
            return [];
        }

        const files = [];
        let pageToken = "";

        do {
            const params = new URLSearchParams({
                q: `'${parent}' in parents and trashed = false`,
                pageSize: "100",
                orderBy: "name desc",
                fields: "nextPageToken,files(id,name,mimeType,modifiedTime,webViewLink)",
                supportsAllDrives: "true",
                includeItemsFromAllDrives: "true"
            });
            if (pageToken) params.set("pageToken", pageToken);

            const data = await fetchJson(`${DRIVE_API}?${params.toString()}`, {
                headers: authHeaders()
            });

            (data.files || []).forEach(file => {
                const parsed = parseHistoricalLedgerName(file.name);
                if (!parsed) return;

                // Historical dashboard snapshots are expected to be Google Sheets.
                // Do not offer other Drive files that the Sheets API cannot read.
                if (file.mimeType !== "application/vnd.google-apps.spreadsheet") return;

                files.push({ ...file, ...parsed });
            });

            pageToken = data.nextPageToken || "";
        } while (pageToken);

        files.sort((a, b) =>
            (b.year - a.year) ||
            (b.week - a.week) ||
            String(b.hhmm).localeCompare(String(a.hhmm)) ||
            String(b.name).localeCompare(String(a.name))
        );

        state.historicalLedgerFiles = files;
        return files;
    }


    function populateDashboardDataVersionOptions() {
        const select = $("dashboard-data-version");
        if (!select) return;

        const currentValue = select.value || "current";
        const archived = state.historicalLedgerFiles || [];

        select.innerHTML = [
            `<option value="current">Current live ledger</option>`,
            ...archived.map(file =>
                `<option value="${escapeAttr(file.id)}">${escapeHtml(file.label)} · ${escapeHtml(file.name)}</option>`
            )
        ].join("");

        if (currentValue === "current" || archived.some(file => file.id === currentValue)) {
            select.value = currentValue;
        } else {
            select.value = "current";
        }

        updateDashboardDataVersionStatus();
    }


    function updateDashboardDataVersionStatus(message) {
        const status = $("dashboard-data-version-status");
        if (!status) return;

        if (message) {
            status.textContent = message;
            return;
        }

        if (state.dashboardSource?.kind === "historical") {
            status.textContent = `Dashboard snapshot: ${state.dashboardSource.name}`;
        } else {
            status.textContent = "Dashboard snapshot: live inventory ledger";
        }
    }


    async function handleDashboardDataVersionChange(event) {
        const value = String(event.target?.value || "current");

        if (value === "current") {
            state.dashboardInventory = state.inventory.slice();
            state.dashboardTransactions = state.transactions.slice();
            state.dashboardClients = state.clients.slice();
            state.dashboardSource = {
                kind: "current",
                id: CONFIG.INVENTORY_LEDGER_SHEET_ID,
                name: CONFIG.INVENTORY_LEDGER_NAME || "Current live ledger"
            };
            updateDashboardDataVersionStatus();
            renderDashboardViews();
            return;
        }

        const file = state.historicalLedgerFiles.find(item => item.id === value);
        if (!file) {
            event.target.value = "current";
            return;
        }

        try {
            event.target.disabled = true;
            updateDashboardDataVersionStatus(`Loading ${file.label}…`);

            const metadata = await sheetsGet(`/${encodeURIComponent(file.id)}`);
            const sheetTitles = (metadata.sheets || []).map(sheet => sheet.properties.title);
            const missing = [CONFIG.INVENTORY_SHEET_NAME, CONFIG.TRANSACTIONS_SHEET_NAME]
                .filter(name => !sheetTitles.includes(name));

            if (missing.length) {
                throw new Error(`The archive ${file.name} is missing: ${missing.join(", ")}.`);
            }

            const [inventoryRows, transactionRows, clientRows] = await Promise.all([
                getValues(file.id, CONFIG.INVENTORY_SHEET_NAME),
                getValues(file.id, CONFIG.TRANSACTIONS_SHEET_NAME),
                sheetTitles.includes(CONFIG.CLIENT_LIST_SHEET_NAME)
                    ? getValues(file.id, CONFIG.CLIENT_LIST_SHEET_NAME)
                    : Promise.resolve([])
            ]);

            const dashboardInventory = parseInventory(inventoryRows);
            const dashboardTransactions = parseTransactions(transactionRows);
            let dashboardClients = parseClients(clientRows);

            if (!dashboardClients.length) {
                dashboardClients = [...new Set(
                    dashboardTransactions
                        .map(item => String(item.client || "").trim())
                        .filter(Boolean)
                )].sort((a, b) => a.localeCompare(b));
            }

            state.dashboardInventory = dashboardInventory;
            state.dashboardTransactions = dashboardTransactions;
            state.dashboardClients = dashboardClients;
            state.dashboardSource = {
                kind: "historical",
                id: file.id,
                name: file.name,
                modifiedTime: file.modifiedTime || ""
            };

            updateDashboardDataVersionStatus();
            renderDashboardViews();
        } catch (error) {
            console.error(error);
            event.target.value = state.dashboardSource?.kind === "historical"
                ? state.dashboardSource.id
                : "current";
            updateDashboardDataVersionStatus(error.message || "Unable to load that archive.");
        } finally {
            event.target.disabled = false;
        }
    }


    // Sheet-creation only ever targets the Assets Inventory Ledger -
    // the routine and alerts sheets are maintained in the Assets Inventory Ledger.
    async function createSheets(
        names
    ) {

        const ledgerId =
            CONFIG.INVENTORY_LEDGER_SHEET_ID;


        const requests =
            names.map(
                title => ({

                    addSheet: {

                        properties: {

                            title

                        }

                    }

                })
            );


        await sheetsPost(

            `/${encodeURIComponent(
                ledgerId
            )}:batchUpdate`,

            {
                requests
            }

        );


        if (
            names.includes(
                CONFIG.INVENTORY_SHEET_NAME
            )
        ) {

            await updateValues(

                ledgerId,

                CONFIG.INVENTORY_SHEET_NAME,

                [
                    [
                        "Asset",
                        "Balance"
                    ]
                ]

            );

        }


        if (names.includes(CONFIG.ROUTINE_SHEET_NAME)) {
            await updateValues(ledgerId, CONFIG.ROUTINE_SHEET_NAME, [FM_ROUTINE_ALERTS.ROUTINE_HEADERS]);
        }

        if (names.includes(CONFIG.ALERTS_SHEET_NAME)) {
            await updateValues(ledgerId, CONFIG.ALERTS_SHEET_NAME, [FM_ROUTINE_ALERTS.ALERT_HEADERS]);
        }


        if (
            names.includes(
                CONFIG.TRANSACTIONS_SHEET_NAME
            )
        ) {

            await updateValues(

                ledgerId,

                CONFIG.TRANSACTIONS_SHEET_NAME,

                [
                    [
                        "Timestamp",
                        "Client",
                        "Movement",
                        "Asset",
                        "Quantity",
                        "User",
                        "Comment",
                        "Image Link"
                    ]
                ]

            );

        }

    }


    async function ensureOperationalSheetHeaders(ledgerId, sheetName, rows, expectedHeaders) {
        if (!rows.length) {
            await updateValues(ledgerId, sheetName, [expectedHeaders]);
            rows.push(expectedHeaders.slice());
            return;
        }
        const existing = rows[0].slice();
        const normalized = existing.map(normalizeHeader);
        const missing = expectedHeaders.filter(header => !normalized.includes(normalizeHeader(header)));
        if (missing.length) {
            existing.push(...missing);
            await updateValues(ledgerId, sheetName, [existing]);
            rows[0] = existing;
        }
    }

    function alertRowFromEntity(alert, status) {
        const headers = state.alertHeaders.length ? state.alertHeaders : FM_ROUTINE_ALERTS.ALERT_HEADERS;
        const normalized = headers.map(normalizeHeader);
        const row = Array.isArray(alert.rawRow) ? headers.map((_, index) => alert.rawRow[index] ?? "") : new Array(headers.length).fill("");
        const put = (aliases, value) => {
            const wanted = aliases.map(normalizeHeader);
            const index = normalized.findIndex(item => wanted.includes(item));
            if (index >= 0) row[index] = value ?? "";
        };
        put(["alert key","key"], alert.key);
        put(["created at","timestamp"], alert.createdAt || new Date().toISOString());
        put(["alert type","type"], alert.type);
        put(["status"], status || alert.status || "OPEN");
        put(["client"], alert.client); put(["asset"], alert.asset);
        put(["required qty","required"], alert.requiredQty); put(["available qty","available"], alert.availableQty);
        put(["shortfall","short by"], alert.shortfall); put(["destination","location"], alert.destination);
        put(["message","alert"], alert.message); put(["attended at"], alert.attendedAt);
        put(["attended by","attended user"], alert.attendedBy); put(["resolution comment","resolution note"], alert.resolutionComment);
        put(["routine key"], alert.routineKey); put(["ignored at","ignore at"], alert.ignoredAt);
        put(["ignored by","ignored user"], alert.ignoredBy); put(["ignore comment","ignore note"], alert.ignoreComment);
        return row;
    }

    async function rebuildAlertSnapshot() {
        const candidates = FM_ROUTINE_ALERTS.buildAlertCandidates({
            inventory: state.inventory,
            transactions: state.transactions,
            routines: state.routines,
            clients: state.clients,
            date: new Date()
        });
        const retainedMap = new Map();
        state.alerts.filter(alert => alert.status === "ATTENDED" || alert.status === "IGNORED").forEach(alert => {
            if (alert.key) retainedMap.set(alert.key, { ...alert, rawRow: Array.isArray(alert.rawRow) ? alert.rawRow.slice() : undefined });
        });
        const retained = [...retainedMap.values()];
        const retainedKeys = new Set(retained.map(alert => alert.key));
        const createdAt = new Date().toISOString();
        const fresh = candidates.filter(item => item.key && !retainedKeys.has(item.key)).map(item => ({
            rowNumber: 0, key: item.key, createdAt, type: item.type, status: "OPEN", client: item.client || "", asset: item.asset || "",
            requiredQty: item.requiredQty || 0, availableQty: item.availableQty || 0, shortfall: item.shortfall || 0, destination: item.destination || "",
            message: item.message || "", attendedAt: "", attendedBy: "", resolutionComment: "", routineKey: item.routineKey || "", ignoredAt: "", ignoredBy: "", ignoreComment: ""
        }));
        const finalAlerts = [...retained, ...fresh].map((alert, index) => ({ ...alert, rowNumber: index + 2 }));
        const headers = state.alertHeaders.length ? state.alertHeaders : FM_ROUTINE_ALERTS.ALERT_HEADERS;
        const lastCol = columnLetter(headers.length);
        const clearRange = `${quoteSheetName(CONFIG.ALERTS_SHEET_NAME)}!A2:${lastCol}`;
        await sheetsPost(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(clearRange)}:clear`, {});
        const writeRange = `${quoteSheetName(CONFIG.ALERTS_SHEET_NAME)}!A1:${lastCol}${finalAlerts.length + 1}`;
        await sheetsPut(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(writeRange)}?valueInputOption=USER_ENTERED`, {
            range: writeRange, majorDimension: "ROWS", values: [headers, ...finalAlerts.map(alert => alertRowFromEntity(alert, alert.status))]
        });
        state.alerts = finalAlerts;
        return finalAlerts;
    }

    // Compatibility aliases for older dashboard code.
    function buildAlertRow(alert, user) {
        return alertRowFromEntity({ ...alert, createdAt: new Date().toISOString() }, "OPEN");
    }
    async function syncGeneratedAlerts() { return rebuildAlertSnapshot(); }
    async function updateAlertStatus(alert, comment) {
        const attendedAt = new Date().toISOString();
        const attendedBy = state.idTokenPayload?.email || readSavedSession()?.email || "Google user";
        Object.assign(alert, { status: "ATTENDED", attendedAt, attendedBy, resolutionComment: comment });
        const headers = state.alertHeaders.length ? state.alertHeaders : FM_ROUTINE_ALERTS.ALERT_HEADERS;
        const lastCol = columnLetter(headers.length);
        const range = `${quoteSheetName(CONFIG.ALERTS_SHEET_NAME)}!A${alert.rowNumber}:${lastCol}${alert.rowNumber}`;
        await sheetsPut(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, { range, majorDimension: "ROWS", values: [alertRowFromEntity(alert, "ATTENDED")] });
    }

    async function getValues(
        spreadsheetId,
        sheetName
    ) {

        const range =
            `${quoteSheetName(
                sheetName
            )}!A:AE`;


        const data =
            await sheetsGet(

                `/${encodeURIComponent(
                    spreadsheetId
                )}/values/${encodeURIComponent(
                    range
                )}`

            );


        return data.values || [];

    }


    async function updateValues(
        spreadsheetId,
        sheetName,
        rows
    ) {

        const range =
            `${quoteSheetName(
                sheetName
            )}!A1`;


        return sheetsPut(

            `/${encodeURIComponent(
                spreadsheetId
            )}/values/${encodeURIComponent(
                range
            )}?valueInputOption=USER_ENTERED`,

            {

                range,

                majorDimension:
                    "ROWS",

                values:
                    rows

            }

        );

    }


    function parseInventory(
        rows
    ) {

        if (!rows.length) {
            return [];
        }


        const header =
            rows[0].map(
                normalizeHeader
            );


        const assetIdx =
            findColumn(

                header,

                [
                    "asset",
                    "asset name",
                    "item",
                    "type"
                ]

            );


        const balanceIdx =
            findColumn(

                header,

                [
                    "balance",
                    "current balance",
                    "stock",
                    "quantity"
                ]

            );

        // Optional audit columns. Older workbooks may not have them.
        const verifiedIdx =
            findColumn(
                header,
                ["verified", "verified status", "is verified"]
            );
        const verificationDateIdx =
            findColumn(
                header,
                ["date", "verification date", "verified date", "audit date"]
            );


        if (assetIdx < 0) {
            return [];
        }


        return rows

            .slice(1)

            .map(
                (row, index) => ({

                    rowNumber:
                        index + 2,

                    asset:
                        String(
                            row[assetIdx] ?? ""
                        ).trim(),

                    balance:
                        balanceIdx >= 0
                            ? numericValue(
                                row[balanceIdx]
                            )
                            : 0,

                    verified:
                        verifiedIdx >= 0
                            ? isTruthyVerificationValue(row[verifiedIdx])
                            : false,

                    verificationDate:
                        verificationDateIdx >= 0
                            ? String(row[verificationDateIdx] ?? "").trim()
                            : "",

                    assetColumn:
                        assetIdx + 1,

                    balanceColumn:
                        balanceIdx >= 0
                            ? balanceIdx + 1
                            : 2

                })
            )

            .filter(
                item =>
                    item.asset
            );

    }


    function parseTransactions(rows) {
        if (!rows.length) return [];
        const header = rows[0].map(normalizeHeader);
        const idx = {
            timestamp: findColumn(header, ["timestamp", "date", "datetime"]),
            client: findColumn(header, ["client", "client name"]),
            movement: findColumn(header, ["movement", "type", "direction"]),
            asset: findColumn(header, ["asset", "asset name", "item"]),
            quantity: findColumn(header, ["quantity", "qty"]),
            user: findColumn(header, ["user", "entered by", "email"]),
            comment: findColumn(header, ["comment", "comments", "notes"]),
            image: findColumn(header, ["image link", "image", "photo", "picture", "photo link", "attachment", "drive link"])
        };

        return rows.slice(1).map(row => ({
            timestamp: idx.timestamp >= 0 ? row[idx.timestamp] ?? "" : "",
            client: idx.client >= 0 ? row[idx.client] ?? "" : "",
            movement: idx.movement >= 0 ? row[idx.movement] ?? "" : "",
            asset: idx.asset >= 0 ? row[idx.asset] ?? "" : "",
            quantity: idx.quantity >= 0 ? numericValue(row[idx.quantity]) : 0,
            user: idx.user >= 0 ? row[idx.user] ?? "" : "",
            comment: idx.comment >= 0 ? String(row[idx.comment] ?? "").trim() : "",
            image: idx.image >= 0 ? String(row[idx.image] ?? "").trim() : ""
        })).filter(item => item.asset || item.client);
    }

    async function ensureTransactionColumns(ledgerId, rows) {
        const header = rows.length
            ? rows[0].slice()
            : ["Timestamp", "Client", "Movement", "Asset", "Quantity", "User"];

        const normalized = header.map(normalizeHeader);
        let changed = false;

        if (!normalized.some(h => ["comment", "comments", "notes"].includes(h))) {
            header.push("Comment");
            normalized.push("comment");
            changed = true;
        }
        if (!normalized.some(h => ["image link", "image", "photo", "picture", "photo link", "attachment", "drive link"].includes(h))) {
            header.push("Image Link");
            normalized.push("image link");
            changed = true;
        }

        if (changed) {
            const range = `${quoteSheetName(CONFIG.TRANSACTIONS_SHEET_NAME)}!A1`;
            await sheetsPut(`/${encodeURIComponent(ledgerId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, {
                range, majorDimension: "ROWS", values: [header]
            });
            if (rows.length) rows[0] = header;
            else rows.push(header);
        }

        return {header, length: header.length};
    }


    function findMainHeaderRow(
        rows
    ) {

        return rows.findIndex(
            row => {

                const headers =
                    row.map(
                        normalizeHeader
                    );


                const hasClient =
                    headers.includes("cleint") ||
                    headers.includes("client") ||
                    headers.includes("collection client");


                const hasLoadType =
                    headers.includes("load type") ||
                    headers.includes("loadtype");


                const hasArrival =
                    headers.includes("planned arrival") ||
                    headers.includes("plannedarrival");


                return (
                    hasClient &&
                    hasLoadType &&
                    hasArrival
                );

            }
        );

    }


    function parseClients(rows) {

        if (!rows.length) {
            return [];
        }

        const header =
            rows[0].map(normalizeHeader);

        const clientIdx =
            findColumn(
                header,
                [
                    "client name",
                    "client",
                    "name"
                ]
            );

        if (clientIdx < 0) {
            return [];
        }

        const clients = new Set();

        rows.slice(1).forEach(row => {
            const value =
                String(row[clientIdx] ?? "").trim();

            if (value) {
                clients.add(value);
            }
        });

        return [...clients].sort(
            (a, b) => a.localeCompare(b)
        );
    }


    function isTodayTransaction(
        timestamp
    ) {

        if (!timestamp) {
            return false;
        }


        const parsed =
            new Date(timestamp);


        if (
            Number.isNaN(
                parsed.getTime()
            )
        ) {

            return (
                String(
                    timestamp
                ).slice(0, 10) ===
                new Date()
                    .toISOString()
                    .slice(0, 10)
            );

        }


        const now =
            new Date();


        return (

            parsed.getFullYear() ===
                now.getFullYear() &&

            parsed.getMonth() ===
                now.getMonth() &&

            parsed.getDate() ===
                now.getDate()

        );

    }


    function openClientDetails(client) {
        $("client-details-title").textContent = client;

        const dateSelect = $("client-details-filter-date");
        const assetSelect = $("client-details-filter-asset");
        populateClientDetailsAssetFilter(client);

        const render = () => {
            const dateFilter = dateSelect?.value || "";
            const assetFilter = assetSelect?.value || "";

            const transactions = getDashboardTransactions()
                .filter(item =>
                    String(item.client || "").trim().toLowerCase() === String(client).trim().toLowerCase() &&
                    (!dateFilter || transactionDateKey(item.timestamp) === dateFilter) &&
                    (!assetFilter || String(item.asset || "").trim() === assetFilter)
                )
                .sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));

            const assets = getClientAssetSummary(transactions);

            $("client-details-subtitle").textContent =
                `${assets.length} asset type${assets.length === 1 ? "" : "s"} · ${transactions.length} transaction${transactions.length === 1 ? "" : "s"}`;

            $("client-details-summary").innerHTML = assets.length
                ? assets.map(item => `<button type="button"
                    class="client-asset-summary-row ${
                        item.net < 0 ? "asset-net-negative" : item.net > 0 ? "asset-net-positive" : "asset-net-zero"
                    }" data-client-asset="${escapeAttr(item.asset)}">
                    <span class="client-asset-name">${escapeHtml(item.asset)}</span>
                    <span class="client-asset-quantity">${formatSignedNumber(item.net)}</span>
                    <span class="client-asset-meta">Sent − Received</span>
                </button>`).join("")
                : `<div class="client-assets-empty"><strong>No asset movements match these filters</strong><span>Try another day or asset type.</span></div>`;

            renderClientAssetGraph(client, assets[0]?.asset || "");

            $("client-details-transactions").innerHTML = transactions.length
                ? transactions.map((item, i) => `<tr>
                    <td>${formatTransactionTime(item.timestamp)}</td>
                    <td><span class="movement-badge ${movementClass(item.movement)}">${escapeHtml(String(item.movement))}</span></td>
                    <td><strong>${escapeHtml(String(item.asset))}</strong></td>
                    <td class="num">${formatNumber(item.quantity)}</td>
                    <td>${escapeHtml(String(item.user || ""))}</td>
                    <td><button type="button" class="secondary info-button" data-client-info-index="${i}">View</button></td>
                </tr>`).join("")
                : emptyRow(6, "No asset transactions recorded for this client.");

            renderTransactionSummary($("client-details-transaction-summary"), transactions, "Client totals");
            state.clientDetailsRows = transactions;
        };

        if (dateSelect) dateSelect.onchange = render;
        if (assetSelect) assetSelect.onchange = render;

        // Start the details modal on the same day filter used by the dashboard.
        if (dateSelect && $("dashboard-filter-date")) dateSelect.value = $("dashboard-filter-date").value || "";
        render();
        $("client-details-modal").classList.remove("hidden");
    }

    function populateClientDetailsAssetFilter(client) {
        const select = $("client-details-filter-asset");
        if (!select) return;

        const current = select.value || "";
        const assets = [...new Set(getDashboardTransactions()
            .filter(t => String(t.client || "").trim().toLowerCase() === String(client).trim().toLowerCase())
            .map(t => String(t.asset || "").trim())
            .filter(Boolean))]
            .sort((a,b) => a.localeCompare(b));

        select.innerHTML = `<option value="">All asset types</option>${assets.map(a => `<option value="${escapeAttr(a)}">${escapeHtml(a)}</option>`).join("")}`;
        if (assets.includes(current)) select.value = current;
    }

    function getClientAssetSummary(transactions) {

        const map = new Map();

        transactions.forEach(item => {

            const asset =
                String(item.asset || "").trim();

            if (!asset) {
                return;
            }

            const key =
                asset.toLowerCase();

            if (!map.has(key)) {
                map.set(
                    key,
                    {
                        asset,
                        sent: 0,
                        received: 0
                    }
                );
            }

            const entry =
                map.get(key);

            const movement =
                String(item.movement || "")
                    .trim()
                    .toUpperCase();

            const quantity =
                Number(item.quantity) || 0;

            if (movement === "SENT") {
                entry.sent += quantity;
            }

            if (movement === "RECEIVED") {
                entry.received += quantity;
            }
        });

        return [...map.values()]
            .map(
                item => ({
                    ...item,
                    net:
                        item.sent -
                        item.received
                })
            )
            .sort(
                (a, b) =>
                    a.asset.localeCompare(b.asset)
            );
    }


    function renderClientAssetGraph(
        client,
        asset
    ) {

        const title =
            $("client-graph-title");

        const subtitle =
            $("client-graph-subtitle");

        const graph =
            $("client-sent-graph");

        if (!asset) {

            title.textContent =
                "Daily sent vs received";

            subtitle.textContent =
                "Select an asset type above to see its daily sent and received history.";

            graph.innerHTML =
                `<div class="client-graph-empty">
                    No asset transaction history to graph.
                 </div>`;

            return;
        }

        title.textContent =
            `${asset} · sent vs received daily`;

        subtitle.textContent =
            "Daily quantity sent and received for this asset, from the Asset Transactions ledger.";

        const rows =
            getDashboardTransactions().filter(
                item =>
                    String(item.client).trim().toLowerCase() ===
                    String(client).trim().toLowerCase() &&
                    String(item.asset).trim().toLowerCase() ===
                    String(asset).trim().toLowerCase() &&
                    ["SENT", "RECEIVED"].includes(
                        String(item.movement).trim().toUpperCase()
                    )
            );

        const daily = new Map();

        rows.forEach(item => {

            const date =
                transactionDateKey(
                    item.timestamp
                );

            if (!date) {
                return;
            }

            if (!daily.has(date)) {
                daily.set(date, {
                    sent: 0,
                    received: 0
                });
            }

            const entry = daily.get(date);
            const quantity = Number(item.quantity) || 0;
            const movement = String(item.movement).trim().toUpperCase();

            if (movement === "SENT") {
                entry.sent += quantity;
            } else if (movement === "RECEIVED") {
                entry.received += quantity;
            }
        });

        const points =
            [...daily.entries()]
                .sort((a, b) => a[0].localeCompare(b[0]));

        if (!points.length) {

            graph.innerHTML =
                `<div class="client-graph-empty">
                    No sent or received transactions recorded for ${escapeHtml(asset)}.
                 </div>`;

            return;
        }

        const width =
            Math.max(
                680,
                points.length * 82
            );

        const height = 300;
        const left = 56;
        const right = 24;
        const top = 42;
        const bottom = 58;
        const chartWidth = width - left - right;
        const chartHeight = height - top - bottom;

        const maxValue =
            Math.max(
                ...points.flatMap(([, value]) => [value.sent, value.received]),
                1
            );

        const yTicks = 4;
        const yGrid = Array.from(
            { length: yTicks + 1 },
            (_, index) => {
                const value =
                    maxValue * (1 - index / yTicks);

                const y =
                    top + (chartHeight * index / yTicks);

                return { value, y };
            }
        );

        const xFor = index =>
            points.length === 1
                ? left + chartWidth / 2
                : left + (index / (points.length - 1)) * chartWidth;

        const yFor = value =>
            top + chartHeight -
            (value / maxValue) * chartHeight;

        const makePath = key =>
            points.map(
                ([, value], index) =>
                    `${index === 0 ? "M" : "L"} ${xFor(index).toFixed(1)} ${yFor(value[key]).toFixed(1)}`
            ).join(" ");

        const sentPath = makePath("sent");
        const receivedPath = makePath("received");

        const gridLines = yGrid.map(tick => `
            <line
                x1="${left}"
                y1="${tick.y.toFixed(1)}"
                x2="${width - right}"
                y2="${tick.y.toFixed(1)}"
                class="graph-grid"
            ></line>
            <text
                x="${left - 9}"
                y="${(tick.y + 4).toFixed(1)}"
                text-anchor="end"
                class="graph-y-label"
            >
                ${formatNumber(Math.round(tick.value))}
            </text>
        `).join("");

        const xLabels = points.map(
            ([date], index) => `
                <text
                    x="${xFor(index).toFixed(1)}"
                    y="${height - 20}"
                    text-anchor="middle"
                    class="graph-label"
                >
                    ${escapeHtml(formatGraphDate(date))}
                </text>
            `
        ).join("");

        const pointsMarkup = points.map(
            ([date, value], index) => {
                const x = xFor(index);
                const label = formatGraphDate(date);

                return `
                    <g>
                        <title>
                            ${escapeHtml(label)}: ${formatNumber(value.sent)} sent, ${formatNumber(value.received)} received
                        </title>
                        <circle
                            cx="${x.toFixed(1)}"
                            cy="${yFor(value.sent).toFixed(1)}"
                            r="4"
                            class="graph-point-sent"
                        ></circle>
                        <circle
                            cx="${x.toFixed(1)}"
                            cy="${yFor(value.received).toFixed(1)}"
                            r="4"
                            class="graph-point-received"
                        ></circle>
                    </g>
                `;
            }
        ).join("");

        graph.innerHTML = `
            <div class="client-graph-legend">
                <span class="graph-legend-item">
                    <span class="graph-legend-dot graph-legend-sent"></span>
                    Sent
                </span>
                <span class="graph-legend-item">
                    <span class="graph-legend-dot graph-legend-received"></span>
                    Received
                </span>
            </div>

            <div class="client-graph-scroll">
                <svg
                    class="client-graph-svg"
                    viewBox="0 0 ${width} ${height}"
                    role="img"
                    aria-label="Daily sent versus received line graph for ${escapeHtml(asset)}"
                >
                    ${gridLines}

                    <line
                        x1="${left}"
                        y1="${top + chartHeight}"
                        x2="${width - right}"
                        y2="${top + chartHeight}"
                        class="graph-axis"
                    ></line>

                    <path
                        d="${sentPath}"
                        class="graph-line graph-line-sent"
                        fill="none"
                    ></path>

                    <path
                        d="${receivedPath}"
                        class="graph-line graph-line-received"
                        fill="none"
                    ></path>

                    ${pointsMarkup}
                    ${xLabels}
                </svg>
            </div>
        `;
    }

    function transactionDateKey(timestamp) {

        if (!timestamp) {
            return "";
        }

        const parsed =
            new Date(timestamp);

        if (
            Number.isNaN(
                parsed.getTime()
            )
        ) {

            const match =
                String(timestamp)
                    .match(/(\d{4}-\d{2}-\d{2})/);

            return match
                ? match[1]
                : "";
        }

        const year =
            parsed.getFullYear();

        const month =
            String(
                parsed.getMonth() + 1
            ).padStart(2, "0");

        const day =
            String(
                parsed.getDate()
            ).padStart(2, "0");

        return `${year}-${month}-${day}`;
    }


    function formatGraphDate(dateKey) {

        const parsed =
            new Date(
                `${dateKey}T00:00:00`
            );

        if (
            Number.isNaN(
                parsed.getTime()
            )
        ) {
            return dateKey;
        }

        return parsed.toLocaleDateString(
            "en-GB",
            {
                day: "2-digit",
                month: "short"
            }
        );
    }


    function formatSignedNumber(value) {

        const number =
            Number(value) || 0;

        if (number > 0) {
            return `+${formatNumber(number)}`;
        }

        return formatNumber(number);
    }


    function closeClientDetails() {

        $("client-details-modal")
            .classList
            .add("hidden");

    }


    function formatTransactionTime(
        timestamp
    ) {

        const parsed =
            new Date(timestamp);


        if (
            Number.isNaN(
                parsed.getTime()
            )
        ) {

            return escapeHtml(
                String(
                    timestamp || ""
                )
            );

        }


        return escapeHtml(

            parsed.toLocaleTimeString(
                "en-GB",
                {

                    hour:
                        "2-digit",

                    minute:
                        "2-digit"

                }
            )

        );

    }


    function movementClass(
        movement
    ) {

        const value =
            String(
                movement
            )
                .trim()
                .toUpperCase();


        if (
            value ===
            "RECEIVED"
        ) {

            return "movement-received";

        }


        if (
            value ===
            "SENT"
        ) {

            return "movement-sent";

        }


        if (
            value ===
            "DISCARD"
        ) {

            return "movement-discard";

        }


        return "movement-other";

    }


    function getDashboardFilters() {
        return {
            client: String($("dashboard-filter-client")?.value || "").trim(),
            date: String($("dashboard-filter-date")?.value || "").trim(),
            asset: String($("dashboard-filter-asset")?.value || "").trim()
        };
    }

    function getDashboardInventory() {
        return Array.isArray(state.dashboardInventory)
            ? state.dashboardInventory
            : state.inventory;
    }

    function getDashboardTransactions() {
        return Array.isArray(state.dashboardTransactions)
            ? state.dashboardTransactions
            : state.transactions;
    }

    function getDashboardClients() {
        const source = Array.isArray(state.dashboardClients)
            ? state.dashboardClients
            : state.clients;
        return source.filter(client => client !== "HSC London(Self)");
    }

    function renderDashboardViews() {
        populateDashboardFilters();
        renderWarehouseAssets();
        renderClientCards();
        renderAudit();
        updateDashboardDataVersionStatus();
    }

    function filteredDashboardTransactions() {
        const filters = getDashboardFilters();
        return getDashboardTransactions().filter(item =>
            (!filters.client || String(item.client || "").trim() === filters.client) &&
            (!filters.asset || String(item.asset || "").trim() === filters.asset) &&
            (!filters.date || transactionDateKey(item.timestamp) === filters.date)
        );
    }

    function transactionSummaryGroups(transactions) {
        const groups = new Map();
        transactions.forEach(item => {
            const client = String(item.client || "Unknown client").trim() || "Unknown client";
            const asset = String(item.asset || "Unknown asset").trim() || "Unknown asset";
            const movement = String(item.movement || "").trim().toUpperCase();
            const direction = movementLabel(movement) || movement || "Other";
            const key = `${client.toLowerCase()}\u0000${asset.toLowerCase()}\u0000${movement}`;
            if (!groups.has(key)) groups.set(key, {client, asset, movement, direction, quantity: 0});
            groups.get(key).quantity += Number(item.quantity) || 0;
        });
        return [...groups.values()].sort((a, b) =>
            a.client.localeCompare(b.client) ||
            a.asset.localeCompare(b.asset) ||
            a.direction.localeCompare(b.direction)
        );
    }

    function renderTransactionSummary(container, transactions, label = "Matching totals") {
        if (!container) return;
        const groups = transactionSummaryGroups(transactions);
        if (!groups.length) {
            container.innerHTML = "";
            container.classList.add("hidden");
            return;
        }
        const rows = groups.map(group => `<div class="transaction-total-row">
            <strong class="transaction-total-client">${escapeHtml(group.client)}</strong>
            <span class="transaction-total-asset">${escapeHtml(group.asset)}</span>
            <span class="movement-badge ${movementClass(group.movement)}">${escapeHtml(group.direction)}</span>
            <strong class="transaction-total-quantity">${formatNumber(group.quantity)}</strong>
        </div>`).join("");
        container.innerHTML = `<div class="transaction-summary-heading">
            <div class="transaction-summary-heading-copy">
                <div class="eyebrow">${escapeHtml(label)}</div>
                <strong>${formatNumber(transactions.length)} matching transaction${transactions.length === 1 ? "" : "s"}</strong>
            </div>
            <div class="transaction-summary-heading-actions">
                <span class="muted transaction-summary-hint">Grouped by client → asset type → direction</span>
                <button type="button" class="secondary transaction-summary-toggle" data-summary-toggle aria-expanded="false">Expand totals</button>
            </div>
        </div>
        <div class="transaction-summary-body hidden" data-summary-body>
            <div class="transaction-summary-table">
                <div class="transaction-summary-head"><span>Client</span><span>Asset type</span><span>Direction</span><span class="num">Total</span></div>
                ${rows}
            </div>
        </div>`;
        container.classList.remove("hidden");
    }

    function toggleTransactionSummary(button) {
        const container = button.closest(".transaction-summary");
        const body = container?.querySelector("[data-summary-body]");
        if (!container || !body) return;
        const expanded = button.getAttribute("aria-expanded") === "true";
        body.classList.toggle("hidden", expanded);
        button.setAttribute("aria-expanded", String(!expanded));
        button.textContent = expanded ? "Expand totals" : "Hide totals";
        container.classList.toggle("is-expanded", !expanded);
    }

    function populateDashboardFilters() {
        const filters = getDashboardFilters();
        const clients = [...new Set(getDashboardTransactions().map(t => String(t.client || "").trim()).filter(Boolean))]
            .sort((a,b) => a.localeCompare(b));
        const assets = [...new Set(getDashboardInventory().map(i => String(i.asset || "").trim()).filter(Boolean))]
            .sort((a,b) => a.localeCompare(b));

        const clientSelect = $("dashboard-filter-client");
        const clientAuditSelect = $("dashboard-filter-client-audit");
        const assetSelect = $("dashboard-filter-asset");
        const assetAuditSelect = $("dashboard-filter-asset-audit");

        const clientOptions = `<option value="">All clients</option>${clients.map(c => `<option value="${escapeAttr(c)}">${escapeHtml(c)}</option>`).join("")}`;
        const assetOptionsHtml = `<option value="">All asset types</option>${assets.map(a => `<option value="${escapeAttr(a)}">${escapeHtml(a)}</option>`).join("")}`;

        if (clientSelect) {
            clientSelect.innerHTML = clientOptions;
            clientSelect.value = clients.includes(filters.client) ? filters.client : "";
        }
        if (clientAuditSelect) {
            clientAuditSelect.innerHTML = clientOptions;
            clientAuditSelect.value = clients.includes(filters.client) ? filters.client : "";
        }
        if (assetSelect) {
            assetSelect.innerHTML = assetOptionsHtml;
            assetSelect.value = assets.includes(filters.asset) ? filters.asset : "";
        }
        if (assetAuditSelect) {
            assetAuditSelect.innerHTML = assetOptionsHtml;
            assetAuditSelect.value = assets.includes(filters.asset) ? filters.asset : "";
        }

        if ($("dashboard-filter-date-audit")) {
            $("dashboard-filter-date-audit").value = filters.date;
        }
    }

    function clearDashboardFilters() {
        ["dashboard-filter-client","dashboard-filter-date","dashboard-filter-asset",
         "dashboard-filter-client-audit","dashboard-filter-date-audit","dashboard-filter-asset-audit"]
            .forEach(id => { if ($(id)) $(id).value = ""; });
        renderClientCards();
        renderWarehouseAssets();
        renderAudit();
        renderFilteredLoadTables();
    }

    function renderFilteredLoadTables() {
        const filters = getDashboardFilters();
        const todayKey = transactionDateKey(new Date().toISOString());
        const dateOk = !filters.date || filters.date === todayKey;
        const clientOk = item => !filters.client || String(item.client || "").trim() === filters.client;
        const assetOk = item => !filters.asset || String(item.asset || "").trim() === filters.asset;

        const due = dateOk
            ? state.todayLoads.filter(item => clientOk(item) && assetOk(item))
            : [];
        const inbound = due.filter(item => item.direction === "Inbound");
        const outbound = due.filter(item => item.direction === "Outbound");

        $("inbound-badge").textContent = inbound.length.toLocaleString();
        $("outbound-badge").textContent = outbound.length.toLocaleString();
        renderLoadTable($("inbound-body"), inbound);
        renderLoadTable($("outbound-body"), outbound);

        [document.querySelector("#inbound-body")?.closest(".load-section"),
         document.querySelector("#outbound-body")?.closest(".load-section")]
            .filter(Boolean)
            .forEach(section => {
                let el = section.querySelector(".filter-note");
                if (!el) {
                    el = document.createElement("p");
                    el.className = "muted filter-note";
                    section.appendChild(el);
                }
                el.textContent = !dateOk && filters.date
                    ? "Routine movements are displayed for today. Clear the Day filter to return to today's routine schedule."
                    : "Routine schedule · no external tracker data is used.";
                el.classList.remove("hidden");
            });
    }


    function renderAudit() {
        const filteredRows = filteredDashboardTransactions();
        const rows = filteredRows.slice(-200).reverse();
        state.auditVisibleRows = rows;

        const count = $("audit-count");
        if (count) count.textContent = filteredRows.length.toLocaleString();

        const body = $("transactions-body");
        if (!body) return;

        body.innerHTML = rows.length
            ? rows.map((item, i) => `<tr>
                <td>${escapeHtml(formatTimestamp(item.timestamp))}</td>
                <td><strong>${escapeHtml(item.client || "")}</strong></td>
                <td><span class="movement-badge ${movementClass(item.movement)}">${escapeHtml(item.movement || "")}</span></td>
                <td>${escapeHtml(item.asset || "")}</td>
                <td class="num">${formatNumber(item.quantity)}</td>
                <td>${escapeHtml(item.user || "")}</td>
                <td><button type="button" class="secondary info-button" data-info-index="${i}">View</button></td>
            </tr>`).join("")
            : emptyRow(7, "No asset movements match the current filters.");

        const filters = getDashboardFilters();
        const filtered = Boolean(filters.client || filters.date || filters.asset);
        renderTransactionSummary($("transaction-summary"), filteredRows, filtered ? "Filtered totals" : "All transaction totals");
    }

    function driveFileIdFromLink(url) {
        const text = String(url || "");
        const m = text.match(/\/d\/([a-zA-Z0-9_-]+)/) || text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
        return m ? m[1] : "";
    }

    async function fillInfoModal(item) {
        $("info-client").textContent = item.client || "—";
        $("info-movement").textContent = movementLabel(item.movement) || "—";
        $("info-asset").textContent = item.asset || "—";
        $("info-quantity").textContent = formatNumber(item.quantity);
        if ($("info-timestamp")) $("info-timestamp").textContent = formatTimestamp(item.timestamp) || "—";
        if ($("info-user")) $("info-user").textContent = item.user || "—";

        const photoWrap = $("info-photo-wrap");
        const commentEl = $("info-comment");
        const emptyEl = $("info-empty");

        if (item.image) {
            const image = $("info-photo");
            FM_MEDIA?.revokeObjectUrl?.(image);
            photoWrap.classList.remove("hidden");
            image.removeAttribute("src");
            image.classList.add("is-loading");
            $("info-photo-link").href = item.image;
            const photoStatus=$("info-photo-status");
            photoStatus?.classList.remove("hidden");
            if(photoStatus) photoStatus.textContent="Loading photo...";
            let loaded=false;
            try {
                loaded=await FM_MEDIA.loadDriveImage(image,item.image,state.accessToken);
                if(!loaded && photoStatus) photoStatus.textContent="Preview unavailable here. Use Open full size in Drive.";
            } finally {
                image.classList.remove("is-loading");
                if(photoStatus && loaded) photoStatus.classList.add("hidden");
            }
        } else {
            photoWrap.classList.add("hidden");
            $("info-photo").removeAttribute("src");
            $("info-photo-status")?.classList.add("hidden");
            $("info-photo-link").removeAttribute("href");
        }

        if (item.comment) {
            commentEl.textContent = item.comment;
            commentEl.classList.remove("hidden");
        } else {
            commentEl.textContent = "";
            commentEl.classList.add("hidden");
        }

        emptyEl.classList.toggle("hidden", Boolean(item.image || item.comment));
        $("info-modal").classList.remove("hidden");
    }

    function openInfoModal(index) {
        const item = state.auditVisibleRows[index];
        if (item) fillInfoModal(item);
    }

    function openClientTransactionInfo(index) {
        const item = state.clientDetailsRows[index];
        if (item) fillInfoModal(item);
    }

    function closeInfoModal() {
        const image = $("info-photo");
        FM_MEDIA?.revokeObjectUrl?.(image);
        $("info-modal")?.classList.add("hidden");
    }

    /* =====================================================
       DASHBOARD
       ===================================================== */


    function renderDashboard() {

        const inbound =
            state.todayLoads.filter(
                item =>
                    item.direction ===
                    "Inbound"
            );


        const outbound =
            state.todayLoads.filter(
                item =>
                    item.direction ===
                    "Outbound"
            );


        $("inbound-badge")
            .textContent =
            inbound.length
                .toLocaleString();


        $("outbound-badge")
            .textContent =
            outbound.length
                .toLocaleString();


        /*
         * NEW:
         * Warehouse asset cards
         */

        renderWarehouseAssets();

        renderDashboardViews();

        renderLoadTable(
            $("inbound-body"),
            inbound
        );


        renderLoadTable(
            $("outbound-body"),
            outbound
        );


        /*
         * Multi-asset movement form
         */
        const clients = state.clients.filter(client => client !== "HSC London(Self)");
        $("client").innerHTML = clients.length ? `<option value="">Select client</option>${clients.map(client => `<option value="${escapeAttr(client)}">${escapeHtml(client)}</option>`).join("")}` : `<option value="">No clients found</option>`;
        const existingRows = [...document.querySelectorAll(".asset-row")];
        if (!existingRows.length) addAssetRow(false);
        document.querySelectorAll(".asset-select").forEach(select => { const value = select.value; select.innerHTML = assetOptions(value); });
        updateAssetRows();
        handleMovementChange();
        updateMovementPreview();


        $("dashboard-date")
            .textContent =
            new Date().toLocaleDateString(
                "en-GB",
                {

                    weekday:
                        "long",

                    day:
                        "numeric",

                    month:
                        "long",

                    year:
                        "numeric"

                }
            );


        renderInventoryManager();

    }


    /*
     * =====================================================
     * NEW WAREHOUSE ASSET CARDS
     * =====================================================
     *
     * Current number:
     *     state.inventory balance
     *
     * Movement indicator:
     *
     * RECEIVED = +
     * SENT     = -
     * DISCARD  = -
     *
     * Only today's transactions are included.
     */


    function getWarehouseAssetMovements() {

        const movements = {};


        getDashboardTransactions()

            .filter(
                item =>
                    isTodayTransaction(
                        item.timestamp
                    )
            )

            .forEach(
                item => {

                    const asset =
                        String(
                            item.asset ?? ""
                        ).trim();


                    if (!asset) {
                        return;
                    }


                    const movement =
                        String(
                            item.movement ?? ""
                        )
                        .trim()
                        .toUpperCase();


                    const quantity =
                        Number(
                            item.quantity
                        ) || 0;


                    if (
                        !movements[
                            asset.toLowerCase()
                        ]
                    ) {

                        movements[
                            asset.toLowerCase()
                        ] = {

                            asset,

                            net:
                                0,

                            received:
                                0,

                            sent:
                                0,

                            discarded:
                                0

                        };

                    }


                    const entry =
                        movements[
                            asset.toLowerCase()
                        ];


                    if (
                        movement ===
                        "RECEIVED"
                    ) {

                        entry.net +=
                            quantity;

                        entry.received +=
                            quantity;

                    }


                    if (
                        movement ===
                        "SENT"
                    ) {

                        entry.net -=
                            quantity;

                        entry.sent +=
                            quantity;

                    }


                    if (
                        movement ===
                        "DISCARD"
                    ) {

                        entry.net -=
                            quantity;

                        entry.discarded +=
                            quantity;

                    }

                }
            );


        return movements;

    }


    function renderWarehouseAssets() {
        const container = $("warehouse-assets");
        const filters = getDashboardFilters();
        const assets = getDashboardInventory().filter(item => !filters.asset || String(item.asset).trim() === filters.asset);

        if (!assets.length) {
            container.innerHTML = `<div class="warehouse-empty">No assets match the current filter.</div>`;
            return;
        }

        const activeDate = filters.date || transactionDateKey(new Date().toISOString());

        container.innerHTML = assets.map(item => {
            const assetName = String(item.asset).trim();
            const isHistoricalSnapshot = state.dashboardSource?.kind === "historical";
            const hasExplicitDay = Boolean(filters.date);
            const movementTotal = getDashboardTransactions()
                .filter(t =>
                    String(t.asset || "").trim().toLowerCase() === assetName.toLowerCase() &&
                    (!activeDate || transactionDateKey(t.timestamp) === activeDate) &&
                    (!filters.client || String(t.client || "").trim() === filters.client)
                )
                .reduce((total, t) => {
                    const movement = String(t.movement || "").trim().toUpperCase();
                    const qty = Number(t.quantity) || 0;
                    return movement === "RECEIVED" ? total + qty :
                        ["SENT","DISCARD"].includes(movement) ? total - qty : total;
                }, 0);

            const indicator = isHistoricalSnapshot && !hasExplicitDay
                ? `<span class="asset-movement neutral">—</span>`
                : movementTotal > 0
                    ? `<span class="asset-movement positive">+${formatNumber(movementTotal)}</span>`
                    : movementTotal < 0
                        ? `<span class="asset-movement negative">${formatNumber(movementTotal)}</span>`
                        : `<span class="asset-movement neutral">0</span>`;

            const verifiedBadge = item.verified
                ? (() => {
                    const verificationDate = formatVerificationDate(item.verificationDate);
                    const tooltip = verificationDate
                        ? `Verified on ${verificationDate}`
                        : "Verified";
                    return `<span class="inventory-verified-badge" title="${escapeAttr(tooltip)}" aria-label="${escapeAttr(tooltip)}">✓ Verified</span>`;
                })()
                : "";

            return `<div class="warehouse-asset-card">
                <div class="warehouse-asset-name-row">
                    <div class="warehouse-asset-name">${escapeHtml(assetName)}</div>
                    ${verifiedBadge}
                </div>
                <div class="warehouse-asset-count">${formatNumber(item.balance)}</div>
                <div class="warehouse-asset-movement">
                    ${indicator}
                    <span class="movement-label">${filters.date ? formatGraphDate(activeDate) : (state.dashboardSource?.kind === "historical" ? "snapshot" : "today")}</span>
                </div>
            </div>`;
        }).join("");
    }


    function getClientSummary(
        client
    ) {

        const loads =
            state.todayLoads.filter(
                item =>
                    item.client.toLowerCase() ===
                    client.toLowerCase()
            );


        const transactions =
            state.transactions.filter(
                item => {

                    if (
                        String(
                            item.client
                        )
                            .trim()
                            .toLowerCase() !==
                        client.toLowerCase()
                    ) {

                        return false;

                    }


                    return isTodayTransaction(
                        item.timestamp
                    );

                }
            );


        const inboundLoads =
            loads.filter(
                item =>
                    item.direction ===
                    "Inbound"
            );


        const outboundLoads =
            loads.filter(
                item =>
                    item.direction ===
                    "Outbound"
            );


        const received =
            transactions.filter(
                item =>
                    String(
                        item.movement
                    )
                        .trim()
                        .toUpperCase() ===
                    "RECEIVED"
            );


        const sent =
            transactions.filter(
                item =>
                    String(
                        item.movement
                    )
                        .trim()
                        .toUpperCase() ===
                    "SENT"
            );


        const discarded =
            transactions.filter(
                item =>
                    String(
                        item.movement
                    )
                        .trim()
                        .toUpperCase() ===
                    "DISCARD"
            );


        const expectedPallets =
            sumNumericText(
                inboundLoads,
                "pallets"
            );


        const expectedLoose =
            sumNumericText(
                inboundLoads,
                "loose"
            );


        const outboundPallets =
            sumNumericText(
                outboundLoads,
                "pallets"
            );


        const outboundLoose =
            sumNumericText(
                outboundLoads,
                "loose"
            );


        const actualReceived =
            received.reduce(
                (
                    sum,
                    item
                ) =>
                    sum +
                    item.quantity,
                0
            );


        const actualSent =
            sent.reduce(
                (
                    sum,
                    item
                ) =>
                    sum +
                    item.quantity,
                0
            );


        const actualDiscarded =
            discarded.reduce(
                (
                    sum,
                    item
                ) =>
                    sum +
                    item.quantity,
                0
            );


        return {

            client,

            loads,

            inboundLoads,

            outboundLoads,

            transactions,

            received,

            sent,

            discarded,

            expectedPallets,

            expectedLoose,

            outboundPallets,

            outboundLoose,

            actualReceived,

            actualSent,

            actualDiscarded,

            expectedInboundTotal:
                expectedPallets +
                expectedLoose,

            expectedOutboundTotal:
                outboundPallets +
                outboundLoose,

            actualInboundTotal:
                actualReceived,

            actualOutboundTotal:
                actualSent +
                actualDiscarded

        };

    }


    function renderClientCards() {
        const container = $("client-cards");
        const query = String($("client-search")?.value || "").trim().toLowerCase();
        const filters = getDashboardFilters();
        const todayKey = transactionDateKey(new Date().toISOString());

        const clients = getDashboardClients().filter(client => {
            if (query && !client.toLowerCase().includes(query)) return false;
            if (filters.client && client !== filters.client) return false;

            const cardDate = filters.date || todayKey;
            const matchingTransactions = getDashboardTransactions().filter(item =>
                String(item.client || "").trim().toLowerCase() === client.toLowerCase() &&
                transactionDateKey(item.timestamp) === cardDate &&
                (!filters.asset || String(item.asset || "").trim() === filters.asset)
            );

            const canUseLoads = state.dashboardSource?.kind !== "historical" && (!filters.date || filters.date === todayKey);
            const hasLoad = canUseLoads && state.todayLoads.some(load =>
                String(load.client || "").trim().toLowerCase() === client.toLowerCase()
            );

            return !filters.asset || matchingTransactions.length > 0 || hasLoad;
        }).sort((a,b) => a.localeCompare(b));

        container.innerHTML = clients.length ? clients.map(client => {
            const cardDate = filters.date || todayKey;
            const transactions = getDashboardTransactions().filter(item =>
                String(item.client || "").trim().toLowerCase() === client.toLowerCase() &&
                transactionDateKey(item.timestamp) === cardDate &&
                (!filters.asset || String(item.asset || "").trim() === filters.asset)
            );

            const received = transactions.filter(t => String(t.movement || "").toUpperCase() === "RECEIVED")
                .reduce((s,t) => s + (Number(t.quantity) || 0), 0);
            const sent = transactions.filter(t => ["SENT","DISCARD"].includes(String(t.movement || "").toUpperCase()))
                .reduce((s,t) => s + (Number(t.quantity) || 0), 0);
            const assets = [...new Set(transactions.map(t => String(t.asset || "").trim()).filter(Boolean))];

            const canUseLoads = !filters.date || filters.date === todayKey;
            const loads = canUseLoads
                ? state.todayLoads.filter(load => String(load.client || "").trim().toLowerCase() === client.toLowerCase())
                : [];
            const expected = loads.reduce((s, load) => s + sumNumericText([load], "pallets") + sumNumericText([load], "loose"), 0);

            const directionClass = getClientCardDirectionClass({
                inboundLoads: loads.filter(l => l.direction === "Inbound"),
                outboundLoads: loads.filter(l => l.direction === "Outbound")
            });

            return `<article class="client-card client-ledger-card ${directionClass}"
                data-client="${escapeAttr(client)}" tabindex="0" role="button">
                <div class="client-card-top">
                    <div>
                        <div class="eyebrow">CLIENT</div>
                        <h3>${escapeHtml(client)}</h3>
                    </div>
                    <span class="today-pill">${transactions.length} movement${transactions.length === 1 ? "" : "s"}</span>
                </div>
                <div class="client-card-stats">
                    <div><span>Received</span><strong>${formatNumber(received)}</strong></div>
                    <div><span>Sent / discard</span><strong>${formatNumber(sent)}</strong></div>
                    <div><span>Expected today</span><strong>${canUseLoads ? formatNumber(expected) : "—"}</strong></div>
                </div>
                <div class="client-card-assets">
                    ${assets.length
                        ? assets.map(a => `<span class="asset-chip">${escapeHtml(a)}</span>`).join("")
                        : `<span class="muted">No matching asset movements</span>`}
                </div>
                <div class="client-card-footer client-ledger-footer">
                    <span>Asset inventory &amp; transactions</span>
                    <button type="button" class="secondary small-button client-show-more"
                        data-client="${escapeAttr(client)}">Show more</button>
                </div>
            </article>`;
        }).join("") : `
            <div class="client-cards-empty">
                <strong>No clients match the current filters</strong>
                <span>Try clearing a filter or changing the client search.</span>
            </div>`;
    }

    function getClientCardDirectionClass(
        summary
    ) {

        if (
            summary.inboundLoads.length &&
            summary.outboundLoads.length
        ) {

            return "direction-mixed";

        }


        if (
            summary.outboundLoads.length
        ) {

            return "direction-outbound";

        }


        if (
            summary.inboundLoads.length
        ) {

            return "direction-inbound";

        }


        return "direction-neutral";

    }


    function sumNumericText(
        items,
        key
    ) {

        return items.reduce(
            (
                sum,
                item
            ) => {

                const value =
                    String(
                        item[key] ?? ""
                    )
                        .replace(
                            /,/g,
                            ""
                        )
                        .trim();


                const parsed =
                    Number(value);


                return Number.isFinite(
                    parsed
                )
                    ? sum + parsed
                    : sum;

            },
            0
        );

    }


    function renderLoadTable(body, items) {
        body.innerHTML = items.length
            ? items.map(item => `
                <tr>
                    <td><strong>${escapeHtml(item.client || "Unknown")}</strong></td>
                    <td>
                        <strong>${escapeHtml(item.plannedArrival || "Routine")}</strong>
                        ${item.loadType ? `<span class="subtext">${escapeHtml(item.loadType)}</span>` : ""}
                    </td>
                    <td><strong>${escapeHtml(item.asset || (/pallet/i.test(item.asset || "") ? "Pallets" : "Assets"))}</strong></td>
                    <td class="num">${formatNumber(item.quantity || 0)}</td>
                    <td>${escapeHtml(item.destination || "—")}</td>
                </tr>`).join("")
            : emptyRow(5, "No routine movements scheduled for today.");
    }


    async function recordMovement(event) {
        event.preventDefault();
        if (movementSubmitBusy) return;
        const movement = $("movement").value;
        const client = movement === "DISCARD" ? "HSC London (Self)" : $("client").value.trim();
        const entries = getAssetEntries();
        const user = state.idTokenPayload?.email || readSavedSession()?.email || state.idTokenPayload?.name || "Google user";
        const comment = $("movement-comment")?.value.trim() || "";
        const photoFile = $("movement-photo")?.files?.[0] || window._fmCameraFile || null;

        if (!["RECEIVED", "SENT", "DISCARD"].includes(movement)) {
            setMovementStatus("Select a movement direction before recording.", true);
            return;
        }

        if (!client || !entries.length || entries.some(x => !x.asset || !Number.isInteger(x.quantity) || x.quantity <= 0)) {
            setMovementStatus("Select a client and add at least one asset with a whole quantity greater than zero.", true);
            return;
        }

        const seen = new Set();
        const balances = new Map(state.inventory.map(item => [item.asset.toLowerCase(), item.balance]));
        const prepared = [];

        for (const entry of entries) {
            const key = entry.asset.toLowerCase();
            if (seen.has(key)) {
                setMovementStatus(`You have selected ${entry.asset} more than once. Combine the quantities.`, true);
                return;
            }
            seen.add(key);

            const item = state.inventory.find(x => x.asset.toLowerCase() === key);
            if (!item) {
                setMovementStatus(`Asset "${entry.asset}" is not present in Inventory.`, true);
                return;
            }

            const current = balances.get(key) || 0;
            const next = movement === "RECEIVED" ? current + entry.quantity : current - entry.quantity;
            if (movement !== "RECEIVED" && next < 0) {
                setMovementStatus(`Cannot remove ${entry.quantity} ${item.asset}. Current balance is ${formatNumber(current)}.`, true);
                return;
            }

            balances.set(key, next);
            prepared.push({ ...entry, item, newBalance: next });
        }

        movementSubmitBusy = true;
        setMovementSubmitting(true, "Checking Google access...");

        try {
            await ensureWriteAccess();
        } catch (authError) {
            movementSubmitBusy = false;
            setMovementSubmitting(false);
            showReconnectUI("Google access expired. Allow Sheets & Drive access, then submit the form again.");
            setMovementStatus(authError.message || "Google access could not be refreshed.", true);
            return;
        }

        setMovementSubmitting(true, photoFile ? "Uploading photo..." : "Recording movement...");

        try {
            setMovementStatus(photoFile ? "Uploading photo..." : "Recording movements...");
            const ledgerId = CONFIG.INVENTORY_LEDGER_SHEET_ID;
            const timestamp = new Date().toISOString();
            let imageLink = "";

            if (photoFile) {
                const folderId = await getOrCreateDailyFolder(transactionDateKey(timestamp));
                const uploadFile = await FM_MEDIA.optimizeImageForUpload(photoFile);
                const uploaded = await uploadImageToDrive(
                    uploadFile,
                    folderId,
                    buildPhotoFilename({client, movement, photoFile})
                );
                imageLink = uploaded.webViewLink || `https://drive.google.com/file/d/${uploaded.id}/view`;
                setMovementSubmitting(true, "Recording movement...");
            }

            const transactionRange = `${quoteSheetName(CONFIG.TRANSACTIONS_SHEET_NAME)}!A:H`;
            await sheetsPost(
                `/${encodeURIComponent(ledgerId)}/values/${encodeURIComponent(transactionRange)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
                {values: prepared.map(x => [timestamp, client, movement, x.asset, x.quantity, user, comment, imageLink])}
            );

            for (const entry of prepared) {
                const balanceCell = columnLetter(entry.item.balanceColumn) + entry.item.rowNumber;
                const inventoryRange = `${quoteSheetName(CONFIG.INVENTORY_SHEET_NAME)}!${balanceCell}`;
                await sheetsPut(
                    `/${encodeURIComponent(ledgerId)}/values/${encodeURIComponent(inventoryRange)}?valueInputOption=USER_ENTERED`,
                    {range: inventoryRange, majorDimension: "ROWS", values: [[entry.newBalance]]}
                );
            }

            resetMovementForm();
            setMovementStatus(`${prepared.length} movement${prepared.length === 1 ? "" : "s"} recorded successfully.`);
            await loadDashboard();
        } catch (error) {
            console.error(error);
            setMovementStatus(error.message || "Unable to record movement.", true);
        } finally {
            movementSubmitBusy = false;
            setMovementSubmitting(false);
        }
    }

    function setMovementSubmitting(busy, detail = "") {
        const overlay = $("movement-submit-loading");
        if (overlay) {
            overlay.classList.toggle("hidden", !busy);
            overlay.setAttribute("aria-hidden", String(!busy));
        }
        if ($("movement-submit-loading-detail") && detail) $("movement-submit-loading-detail").textContent = detail;
        const submit = $("movement-form")?.querySelector("button[type=submit]");
        if (submit) {
            if (!submit.dataset.originalText) submit.dataset.originalText = submit.textContent;
            submit.disabled = busy;
            submit.textContent = busy ? "Saving..." : submit.dataset.originalText;
        }
        $("add-asset-row")?.toggleAttribute("disabled", busy);
    }

    function handleDashboardPhotoChange() {
        const file = $("movement-photo")?.files?.[0];
        const wrap = $("photo-preview-wrap");
        const clear = $("clear-photo");
        if (!file) {
            wrap?.classList.add("hidden");
            clear?.classList.add("hidden");
            return;
        }
        $("photo-preview").src = URL.createObjectURL(file);
        wrap?.classList.remove("hidden");
        clear?.classList.remove("hidden");
    }

    function clearDashboardPhoto() {
        const input = $("movement-photo");
        if (input) input.value = "";
        if ($("movement-photo-camera")) $("movement-photo-camera").value = "";
        $("photo-preview-wrap")?.classList.add("hidden");
        $("photo-preview")?.removeAttribute("src");
        $("clear-photo")?.classList.add("hidden");
    }

    async function getOrCreateDailyFolder(dateStr) {
        if (!window._fmDriveFolderCache) window._fmDriveFolderCache = {};
        if (window._fmDriveFolderCache[dateStr]) return window._fmDriveFolderCache[dateStr];

        const parent = CONFIG.DRIVE_PHOTOS_PARENT_FOLDER_ID;
        const query = [
            `name = '${escapeDriveQuery(dateStr)}'`,
            `mimeType = 'application/vnd.google-apps.folder'`,
            `'${parent}' in parents`,
            'trashed = false'
        ].join(" and ");

        const found = await fetchJson(
            `${DRIVE_API}?q=${encodeURIComponent(query)}&pageSize=1&fields=files(id,name)`,
            {headers: authHeaders()}
        );

        let id = found.files?.[0]?.id;
        if (!id) {
            const created = await fetchJson(DRIVE_API, {
                method: "POST",
                headers: {...authHeaders(), "Content-Type": "application/json"},
                body: JSON.stringify({name: dateStr, mimeType: "application/vnd.google-apps.folder", parents: [parent]})
            });
            id = created.id;
        }
        window._fmDriveFolderCache[dateStr] = id;
        return id;
    }

    async function uploadImageToDrive(file, folderId, filename) {
        const metadata = {name: filename, parents: [folderId], mimeType: file.type || "image/jpeg"};
        const boundary = "fmdashboard" + Math.random().toString(36).slice(2);
        const delimiter = `--${boundary}\r\n`;
        const closeDelim = `\r\n--${boundary}--`;
        const metaPart = delimiter + "Content-Type: application/json; charset=UTF-8\r\n\r\n" + JSON.stringify(metadata) + "\r\n";
        const mediaHeader = delimiter + `Content-Type: ${metadata.mimeType}\r\n\r\n`;
        const body = new Blob([metaPart, mediaHeader, file, closeDelim]);

        return fetchJson(
            `https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink`,
            {
                method: "POST",
                headers: {...authHeaders(), "Content-Type": `multipart/related; boundary=${boundary}`},
                body
            }
        );
    }

    function buildPhotoFilename(data) {
        const safeClient = String(data.client || "client").replace(/[^a-z0-9]+/gi, "-").replace(/^-+|-+$/g, "") || "client";
        const stamp = new Date().toISOString().replace(/[:.]/g, "-");
        const ext = data.photoFile?.type === "image/jpeg" ? ".jpg" : ((data.photoFile?.name?.match(/\.[a-zA-Z0-9]+$/) || [".jpg"])[0]);
        return `${safeClient}-${data.movement}-${stamp}${ext}`;
    }

    function assetOptions(selected = "") {
        return state.inventory.length ? `<option value="">Select asset type</option>${state.inventory.map(item => `<option value="${escapeAttr(item.asset)}" ${item.asset === selected ? "selected" : ""}>${escapeHtml(item.asset)}</option>`).join("")}` : `<option value="">No assets configured</option>`;
    }

    function addAssetRow(focus = true) {
        const wrap = $("asset-rows");
        const row = document.createElement("div");
        row.className = "asset-row";
        row.innerHTML = `<div class="asset-row-number"></div><label><span>Asset type</span><select class="asset-select" required>${assetOptions()}</select></label><label><span>Quantity</span><input class="asset-quantity" type="number" min="1" step="1" inputmode="numeric" placeholder="0" required></label><button type="button" class="secondary remove-asset" data-remove-asset="0">Remove</button>`;
        wrap.appendChild(row);
        updateAssetRows();
        if (focus) row.querySelector(".asset-select").focus();
    }

    function removeAssetRow(index) {
        const rows = [...document.querySelectorAll(".asset-row")];
        const row = rows[index];
        if (!row) return;

        if (rows.length === 1) {
            row.querySelector(".asset-select").value = "";
            row.querySelector(".asset-quantity").value = "";
        } else {
            row.remove();
        }

        updateAssetRows();
        updateMovementPreview();
    }

    function updateAssetRows() {
        const rows = [...document.querySelectorAll(".asset-row")];
        rows.forEach((row, index) => {
            const button = row.querySelector("[data-remove-asset]");
            row.querySelector(".asset-row-number").textContent = String(index + 1).padStart(2, "0");
            button.dataset.removeAsset = index;
            button.disabled = false;
            button.textContent = rows.length === 1 ? "Clear" : "Remove";
            button.setAttribute("aria-label", rows.length === 1 ? "Clear asset row" : "Remove asset");
        });
    }

    function getAssetEntries() {
        return [...document.querySelectorAll(".asset-row")].map(row => ({ asset: row.querySelector(".asset-select").value, quantity: Number(row.querySelector(".asset-quantity").value) }));
    }

    function movementLabel(value) {
        switch (String(value || '').toUpperCase()) {
            case 'RECEIVED':
                return 'Received';
            case 'SENT':
                return 'Sent';
            case 'DISCARD':
                return 'Discard';
            default:
                return String(value || '');
        }
    }


    function updateMovementPreview() {
        const movement = $("movement")?.value || "";
        const client = movement === "DISCARD" ? "HSC London (Self)" : $("client")?.value.trim() || "";
        const entries = getAssetEntries().filter(x => x.asset && Number.isInteger(x.quantity) && x.quantity > 0);
        if (!movement) {
            $("preview-text").textContent = "Select a movement, client and add one or more asset types with quantities.";
            return;
        }
        if (!client || !entries.length) {
            $("preview-text").textContent = "Select a client and add one or more asset types with quantities.";
            return;
        }
        $("preview-text").textContent = `${movementLabel(movement)} ${entries.map(x => `${formatNumber(x.quantity)} × ${x.asset}`).join(" • ")} for ${client}`;
    }

    function resetMovementForm() {
        $("movement-form").reset();
        $("client").disabled = false;
        $("asset-rows").innerHTML = "";
        addAssetRow(false);
        setDefaultTimestamp();
        handleMovementChange();
        clearDashboardPhoto();
        window._fmCameraFile = null;
    }


    async function addAssetType() {

        const input =
            $("new-asset-name");


        const balanceInput =
            $("new-asset-balance");


        const assetName =
            input.value.trim();


        const balance =
            Number(
                balanceInput.value
            );


        if (!assetName) {

            setInventoryManageStatus(
                "Enter an asset type name.",
                true
            );

            return;

        }


        if (
            !Number.isInteger(
                balance
            ) ||
            balance < 0
        ) {

            setInventoryManageStatus(
                "Starting balance must be a whole number of zero or more.",
                true
            );

            return;

        }


        const exists =
            state.inventory.some(
                item =>
                    item.asset
                        .toLowerCase() ===
                    assetName.toLowerCase()
            );


        if (exists) {

            setInventoryManageStatus(
                "That asset type already exists.",
                true
            );

            return;

        }


        try {

            setInventoryManageStatus(
                "Adding asset type..."
            );


            const range =
                `${quoteSheetName(
                    CONFIG.INVENTORY_SHEET_NAME
                )}!A:B`;


            await sheetsPost(

                `/${encodeURIComponent(
                    CONFIG.INVENTORY_LEDGER_SHEET_ID
                )}/values/${encodeURIComponent(
                    range
                )}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,

                {

                    values: [[

                        assetName,

                        balance

                    ]]

                }

            );


            input.value = "";

            balanceInput.value =
                "0";


            setInventoryManageStatus(
                "Asset type added successfully."
            );


            await loadDashboard();


        } catch (error) {

            console.error(error);


            setInventoryManageStatus(
                error.message,
                true
            );

        }

    }


    async function deleteAssetType(
        rowNumber
    ) {

        const item =
            state.inventory.find(
                inventoryItem =>
                    inventoryItem.rowNumber ===
                    rowNumber
            );


        if (!item) {
            return;
        }


        if (
            item.balance !== 0
        ) {

            setInventoryManageStatus(
                `Cannot remove ${item.asset} because its current balance is ${item.balance}. Set the balance to zero first.`,
                true
            );

            return;

        }


        const confirmed =
            window.confirm(
                `Remove the asset type "${item.asset}" from inventory?`
            );


        if (!confirmed) {
            return;
        }


        try {

            setInventoryManageStatus(
                "Removing asset type..."
            );


            const metadata =
                await sheetsGet(
                    `/${encodeURIComponent(
                        CONFIG.INVENTORY_LEDGER_SHEET_ID
                    )}`
                );


            const sheet =
                (metadata.sheets || [])
                    .find(
                        sheet =>
                            sheet.properties.title ===
                            CONFIG.INVENTORY_SHEET_NAME
                    );


            if (!sheet) {

                throw new Error(
                    `Sheet "${CONFIG.INVENTORY_SHEET_NAME}" was not found.`
                );

            }


            const sheetId =
                sheet.properties.sheetId;


            await sheetsPost(

                `/${encodeURIComponent(
                    CONFIG.INVENTORY_LEDGER_SHEET_ID
                )}:batchUpdate`,

                {

                    requests: [

                        {

                            deleteDimension: {

                                range: {

                                    sheetId,

                                    dimension:
                                        "ROWS",

                                    startIndex:
                                        rowNumber - 1,

                                    endIndex:
                                        rowNumber

                                }

                            }

                        }

                    ]

                }

            );


            setInventoryManageStatus(
                "Asset type removed successfully."
            );


            await loadDashboard();


        } catch (error) {

            console.error(error);


            setInventoryManageStatus(
                error.message,
                true
            );

        }

    }


    async function renameAssetType(
        rowNumber
    ) {

        const item =
            state.inventory.find(
                inventoryItem =>
                    inventoryItem.rowNumber ===
                    rowNumber
            );


        if (!item) {
            return;
        }


        const newName =
            window.prompt(
                `Rename "${item.asset}" to:`,
                item.asset
            );


        if (
            newName === null
        ) {
            return;
        }


        const cleanName =
            newName.trim();


        if (!cleanName) {

            setInventoryManageStatus(
                "Asset name cannot be empty.",
                true
            );

            return;

        }


        const duplicate =
            state.inventory.some(
                other =>
                    other.rowNumber !==
                        rowNumber &&

                    other.asset
                        .toLowerCase() ===
                    cleanName.toLowerCase()
            );


        if (duplicate) {

            setInventoryManageStatus(
                "Another asset already has that name.",
                true
            );

            return;

        }


        try {

            setInventoryManageStatus(
                "Renaming asset type..."
            );


            const assetCell =
                columnLetter(
                    item.assetColumn
                ) +
                item.rowNumber;


            const range =
                `${quoteSheetName(
                    CONFIG.INVENTORY_SHEET_NAME
                )}!${assetCell}`;


            await sheetsPut(

                `/${encodeURIComponent(
                    CONFIG.INVENTORY_LEDGER_SHEET_ID
                )}/values/${encodeURIComponent(
                    range
                )}?valueInputOption=USER_ENTERED`,

                {

                    range,

                    majorDimension:
                        "ROWS",

                    values: [[
                        cleanName
                    ]]

                }

            );


            setInventoryManageStatus(
                "Asset type renamed successfully."
            );


            await loadDashboard();


        } catch (error) {

            console.error(error);


            setInventoryManageStatus(
                error.message,
                true
            );

        }

    }


    function openInventoryManager() {

        renderInventoryManager();


        $("inventory-modal")
            .classList
            .remove("hidden");

    }


    function closeInventoryManager() {

        $("inventory-modal")
            .classList
            .add("hidden");

    }


    function renderInventoryManager() {

        const body =
            $("manage-inventory-body");


        if (!body) {
            return;
        }


        body.innerHTML =

            state.inventory.length

                ? state.inventory.map(
                    item => `

                        <tr>

                            <td>

                                <strong>
                                    ${escapeHtml(
                                        item.asset
                                    )}
                                </strong>

                            </td>

                            <td class="num">

                                ${formatNumber(
                                    item.balance
                                )}

                            </td>

                            <td
                                class="actions-cell"
                            >

                                <button
                                    class="table-action"
                                    data-action="rename"
                                    data-row="${item.rowNumber}"
                                >
                                    Rename
                                </button>


                                <button
                                    class="table-action danger"
                                    data-action="delete"
                                    data-row="${item.rowNumber}"
                                >
                                    Remove
                                </button>

                            </td>

                        </tr>

                    `
                ).join("")

                : emptyRow(
                    3,
                    "No asset types configured."
                );

    }


    function handleMovementChange() {
        const movement = $("movement")?.value || "";
        const form = $("movement-form");
        const strip = $("movement-mode-strip");
        if (strip) {
            const copy = { RECEIVED: ["RECEIVED", "Assets coming into warehouse", "mode-received"], SENT: ["SENT", "Assets leaving warehouse", "mode-sent"], DISCARD: ["DISCARDED", "Assets removed from inventory", "mode-discard"] };
            const mode = copy[movement];
            strip.textContent = mode ? `${mode[0]} · ${mode[1]}` : "Select a movement to set the form mode.";
            strip.className = `movement-mode-strip ${mode ? mode[2] : ""}`;
        }
        if (form) {
            form.classList.remove("mode-received", "mode-sent", "mode-discard");
            if (["RECEIVED", "SENT", "DISCARD"].includes(movement)) form.classList.add(`mode-${movement.toLowerCase()}`);
        }

        const client = $("client");
        if (!client) return;

        if (movement === "DISCARD") {
            client.value = "HSC London (Self)";
            client.disabled = true;
            if ($("client-help")) $("client-help").textContent = "Discarded assets are automatically recorded against HSC London (Self).";
        } else {
            client.disabled = false;
            if (client.value === "HSC London (Self)") client.value = "";
            if ($("client-help")) $("client-help").textContent = "";
        }

        updateMovementPreview();
    }


    function setDefaultTimestamp() {

        const input =
            $("timestamp");


        if (!input) {
            return;
        }


        const now =
            new Date();


        input.value =

            `${String(
                now.getHours()
            ).padStart(2, "0")}:${String(
                now.getMinutes()
            ).padStart(2, "0")}`;

    }


    function updateAlertNavBadge() {
        const count = state.alerts.filter(alert => alert?.status === "OPEN").length;
        document.querySelectorAll("#alerts-nav-badge").forEach(el => {
            el.textContent = count.toLocaleString();
            el.classList.toggle("hidden", count === 0);
        });
    }

    function updateAlertNavBadges() { updateAlertNavBadge(); }

    function signOut() {
        const saved = readSavedSession();
        if (state.idTokenPayload?.sub) { try { google.accounts.id.revoke(state.idTokenPayload.sub, () => {}); } catch (_) {} }
        if (saved?.sub && saved.sub !== state.idTokenPayload?.sub) { try { google.accounts.id.revoke(saved.sub, () => {}); } catch (_) {} }
        localStorage.removeItem(SESSION_KEY);
        window.FM_AUTH_CACHE?.clear?.();
        state.idTokenPayload = null; state.accessToken = null; state.workbookId = null; state.workbookName = null; state.inventory = []; state.transactions = []; state.dashboardInventory = []; state.dashboardTransactions = []; state.dashboardClients = []; state.historicalLedgerFiles = []; state.dashboardSource = {kind:"current", id:null, name:"Current live ledger"}; state.todayLoads = []; state.routines = []; state.alerts = []; state.clients = [];
        $("dashboard").classList.add("hidden"); $("login-card").classList.remove("hidden"); $("google-signin-button").classList.remove("hidden"); $("grant-access").classList.add("hidden"); $("sign-out").classList.add("hidden"); $("user-photo").classList.add("hidden"); $("user-name").textContent = "Not signed in"; $("user-email").textContent = "";
        closeInventoryManager();
    }


    function saveSession() { if (state.idTokenPayload) localStorage.setItem(SESSION_KEY, JSON.stringify({ name: state.idTokenPayload.name || "Google user", email: state.idTokenPayload.email || "", picture: state.idTokenPayload.picture || "", sub: state.idTokenPayload.sub || "" })); }
    function readSavedSession() { try { return JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); } catch (_) { return null; } }
    function setUserProfile(profile) { $("user-name").textContent = profile?.name || "Google user"; $("user-email").textContent = profile?.email || ""; if (profile?.picture) { $("user-photo").src = profile.picture; $("user-photo").classList.remove("hidden"); } }
    function hideLogin() { $("google-signin-button").classList.add("hidden"); $("grant-access").classList.add("hidden"); $("sign-out").classList.remove("hidden"); $("login-card").classList.add("hidden"); $("dashboard").classList.remove("hidden"); }

    function showReconnectUI(message = "Google access expired. Allow Sheets & Drive access to continue.") {
        window.FM_CONNECTION_UI?.hide();
        $("login-card")?.classList.remove("hidden");
        $("google-signin-button")?.classList.add("hidden");
        $("grant-access")?.classList.remove("hidden");
        const title = $("login-card")?.querySelector("h2");
        if (title) title.textContent = "Reconnect Google Sheets";
        setAuthStatus(message, true);
    }

    function refreshSavedSessionIfNeeded() {
        const saved = readSavedSession();
        if (!saved || !window.google?.accounts?.oauth2) return;
        if (window.FM_AUTH_CACHE?.read?.(saved.email)?.token) return;
        attemptSilentAccess(saved.email);
    }

    window.addEventListener("pageshow", refreshSavedSessionIfNeeded);
    document.addEventListener("visibilitychange", () => {
        if (!document.hidden) refreshSavedSessionIfNeeded();
    });


    function authHeaders() {

        return {

            Authorization:
                `Bearer ${state.accessToken}`

        };

    }


    async function sheetsGet(
        path
    ) {

        return fetchJson(

            SHEETS_API + path,

            {

                headers:
                    authHeaders()

            }

        );

    }


    async function sheetsPost(
        path,
        body
    ) {

        return fetchJson(

            SHEETS_API + path,

            {

                method:
                    "POST",

                headers: {

                    ...authHeaders(),

                    "Content-Type":
                        "application/json"

                },

                body:
                    JSON.stringify(
                        body
                    )

            }

        );

    }


    async function sheetsPut(
        path,
        body
    ) {

        return fetchJson(

            SHEETS_API + path,

            {

                method:
                    "PUT",

                headers: {

                    ...authHeaders(),

                    "Content-Type":
                        "application/json"

                },

                body:
                    JSON.stringify(
                        body
                    )

            }

        );

    }


    async function fetchJson(url, options = {}) {
        let response = await fetch(url, options);
        if (response.status === 401 && !options.__retried) {
            try {
                await acquireAccessToken("none", state.idTokenPayload?.email || readSavedSession()?.email, {forceRefresh: true});
                return fetchJson(url, { ...options, __retried: true, headers: { ...(options.headers || {}), ...authHeaders() } });
            } catch (_) {
                state.accessToken = null;
                window.FM_AUTH_CACHE?.clear?.();
                showReconnectUI("Google Sheets access expired. Allow Sheets & Drive access to reconnect.");
                throw new Error("Google Sheets access expired. Allow Sheets & Drive access to reconnect.");
            }
        }
        const text = await response.text(); let data = {}; try { data = text ? JSON.parse(text) : {}; } catch (_) {}
        if (!response.ok) throw new Error(data?.error?.message || `Request failed (${response.status})`);
        return data;
    }


    function decodeJwtPayload(
        jwt
    ) {

        const part =
            jwt.split(".")[1];


        const normalized =
            part
                .replace(
                    /-/g,
                    "+"
                )
                .replace(
                    /_/g,
                    "/"
                );


        const padded =
            normalized +
            "=".repeat(

                (
                    4 -
                    normalized.length % 4
                ) % 4

            );


        return JSON.parse(

            decodeURIComponent(

                Array.from(
                    atob(padded)
                )

                .map(
                    character =>
                        `%${character
                            .charCodeAt(0)
                            .toString(16)
                            .padStart(2, "0")}`
                )

                .join("")

            )

        );

    }


    function normalizeHeader(
        value
    ) {

        return String(
            value ?? ""
        )
            .trim()
            .toLowerCase()
            .replace(
                /\s+/g,
                " "
            );

    }


    function findColumn(
        header,
        names
    ) {

        const normalized =
            names.map(
                normalizeHeader
            );


        return header.findIndex(
            value =>
                normalized.includes(
                    value
                )
        );

    }


    function isTruthyVerificationValue(value) {
        if (value === true || value === 1) return true;
        const normalized = String(value ?? "").trim().toLowerCase();
        return ["true", "1", "yes", "y"].includes(normalized);
    }


    function formatVerificationDate(value) {
        if (value === null || value === undefined || String(value).trim() === "") return "";

        const raw = String(value).trim();
        const numeric = Number(raw);
        const date = Number.isFinite(numeric) && numeric > 20000 && numeric < 100000
            ? new Date((numeric - 25569) * 86400 * 1000)
            : new Date(raw);

        if (Number.isNaN(date.getTime())) return raw;

        return date.toLocaleDateString("en-GB", {
            day: "2-digit",
            month: "short",
            year: "numeric"
        });
    }


    function numericValue(
        value
    ) {

        const text =
            String(
                value ?? ""
            )
                .replace(
                    /,/g,
                    ""
                )
                .trim();


        if (!text) {
            return 0;
        }


        const number =
            Number(text);


        return Number.isFinite(
            number
        )
            ? number
            : 0;

    }


    function formatNumber(
        value
    ) {

        return Number(
            value || 0
        )
            .toLocaleString(
                "en-GB"
            );

    }


    function quoteSheetName(
        sheetName
    ) {

        return "'" +

            String(
                sheetName
            )
                .replace(
                    /'/g,
                    "''"
                ) +

            "'";

    }


    function columnLetter(
        number
    ) {

        let result =
            "";


        while (
            number > 0
        ) {

            const remainder =
                (
                    number - 1
                ) % 26;


            result =
                String.fromCharCode(
                    65 + remainder
                ) +
                result;


            number =
                Math.floor(
                    (
                        number - 1
                    ) / 26
                );

        }


        return result;

    }


    function escapeDriveQuery(
        value
    ) {

        return String(
            value
        )
            .replace(
                /\\/g,
                "\\\\"
            )
            .replace(
                /'/g,
                "\\'"
            );

    }


    function escapeHtml(
        value
    ) {

        return String(
            value ?? ""
        )
            .replace(
                /[&<>'"]/g,
                character =>
                    ({

                        "&":
                            "&amp;",

                        "<":
                            "&lt;",

                        ">":
                            "&gt;",

                        "'":
                            "&#39;",

                        '"':
                            "&quot;"

                    }[
                        character
                    ])

            );

    }


    function escapeAttr(
        value
    ) {

        return escapeHtml(
            value
        );

    }


    function emptyRow(
        span,
        text
    ) {

        return `

            <tr>

                <td
                    colspan="${span}"
                    class="empty"
                >
                    ${escapeHtml(
                        text
                    )}
                </td>

            </tr>

        `;

    }


    function setAuthStatus(
        text,
        error = false
    ) {

        const element =
            $("auth-status");


        element.textContent =
            text;


        element.className =
            `status ${
                error
                    ? "error"
                    : ""
            }`;

    }


    function setMovementStatus(
        text,
        error = false
    ) {

        const element =
            $("movement-status");


        element.textContent =
            text;


        element.className =
            `status ${
                error
                    ? "error"
                    : ""
            }`;

    }


    function setInventoryManageStatus(
        text,
        error = false
    ) {

        const element =
            $("inventory-manage-status");


        element.textContent =
            text;


        element.className =
            `status ${
                error
                    ? "error"
                    : ""
            }`;

    }


    function setSyncStatus(
        text,
        error = false
    ) {

        const element =
            $("sync-status");


        element.textContent =
            text;


        element.className =
            `muted ${
                error
                    ? "error"
                    : ""
            }`;

    }


})();