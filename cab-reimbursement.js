// ─────────────────────────────────────────────────────────────────────────────
// CAB REIMBURSEMENT — Optimised JS
// ─────────────────────────────────────────────────────────────────────────────
const DEBUG = false;

document.addEventListener("DOMContentLoaded", () => {

    pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";


    const circle = document.querySelector(".progress-ring-circle");

    let circumference = 0;

    if (circle) {
        const radius = circle.r.baseVal.value;
        circumference = 2 * Math.PI * radius;

        circle.style.strokeDasharray = circumference;
        circle.style.strokeDashoffset = circumference;
    }

    // ─── STATE ────────────────────────────────────────────────────────────────
    const state = {
        receipts: [],
        monthCounts: {},
        activeReceiptId: null,
        lastScrollPosition: 0,
        pdfViewer: { currentPage: 1, totalPages: 1 },
    };

    // ─── CONSTANTS ────────────────────────────────────────────────────────────
    const DPR = Math.min(window.devicePixelRatio || 1, 2);
    const TOTAL_REGEX = /Total\s*(?:₹|Rs\.?|INR)?\s*(\d+(?:\.\d{1,2})?)/i;
    const RAPIDO_REGEX = /Selected\s*Price\s*(?:₹|Rs\.?|INR)?\s*(\d+(?:\.\d{1,2})?)/i;
    const CURRENCY_REGEX = /(?:₹|Rs\.?|INR)\s*(\d+(?:\.\d{1,2})?)/g;

    const locationKeywords = [
        { keys: ["Jorbagh", "Sri Aurobindo Marg", "Jor Bagh", "Safdarjung Airport", "Safdarjung", "safdarjung Airport Area", "Satya Sadan"], short: "Jor Bagh" },
        { keys: ["Rajaji Marg Vijay Chowk Area", "Vijay Chowk Area"], short: "Sena Bhawan" },
        { keys: ["Electronics Niketan", "CGO complex"], short: "Electronics Niketan" },
        { keys: ["Scope Complex"], short: "Scope Complex" },
        { keys: ["Punjabi Bagh", "Punjabi Bagh Enclave", "West Punjabi Bagh"], short: "Punjabi Bagh" },
        { keys: ["GPO Complex", "Barapullah Rd", "Aviation Colony", "INA Colony"], short: "Ayush Bhavan" },
        { keys: ["INA Metro Station"], short: "INA Metro Station" },
        { keys: ["Directorate General of Information Systems", "Shankar Vihar"], short: "DGIS" },
        { keys: ["Central Secretariat", "Udyog Bhawan"], short: "Central Secretariat" },
        { keys: ["Subroto Park"], short: "Subroto Park" },
        { keys: ["Shalimar Bagh"], short: "Shalimar Bagh" },
        { keys: ["IGNOU", "Indira Gandhi National Open University"], short: "IGNOU" },
        { keys: ["Defence Colony"], short: "Defence Colony" },
        { keys: ["KG Marg", "KG M arg"], short: "KG Marg" },
    ];

    // ─── DOM CACHE ────────────────────────────────────────────────────────────
    const $ = id => document.getElementById(id);
    const els = {
        dropZone: $("dropZone"),
        fileInput: $("fileInput"),
        stepUpload: $("step-upload"),
        stepEditor: $("step-editor"),
        stepSuccess: $("step-success"),
        receiptGrid: $("receiptGrid"),
        receiptContainer: $("receiptCardsContainer"),
        pdfDetailView: $("pdfDetailView"),
        pdfPagesContainer: $("pdfPagesContainer"),
        detailFileName: $("detailFileName"),
        pdfPagePill: $("pdfPagePill"),
        pdfProgressFill: document.querySelector(".pdf-progress-fill"),
        cabTableBody: document.querySelector("#cabTable tbody"),
        loadingOverlay: $("loadingOverlay"),
        loadingText: $("loadingText"),
        itemCount: $("itemCount"),
        uiTotalAmount: $("uiTotalAmount"),
        monthInput: $("month"),
        empNameInput: $("empName"),
        designationInput: $("designation"),
        signatureInput: $("signature"),
        signaturePreview: $("signaturePreview"),
        signatureImg: $("signatureImg"),
        generateBtn: $("generateBtn"),
    };

    // ─── TOAST ────────────────────────────────────────────────────────────────
    let toastTimer = null;
    const toastEl = (() => {
        const t = document.createElement("div");
        t.className = "cab-toast";
        document.body.appendChild(t);
        return t;
    })();

    function showToast(msg, type = "info", duration = 3000) {
        clearTimeout(toastTimer);
        const icons = { success: "fa-circle-check", error: "fa-circle-exclamation", info: "fa-circle-info" };
        toastEl.className = `cab-toast toast-${type}`;
        toastEl.innerHTML = `<i class="fa-solid ${icons[type] || icons.info}"></i><span>${msg}</span>`;
        // Force reflow so transition fires
        void toastEl.offsetWidth;
        toastEl.classList.add("show");
        toastTimer = setTimeout(() => toastEl.classList.remove("show"), duration);
    }

    // ─── VALIDATION ALERT — highlights the exact field ───────────────────────
    /**
     * Shows a polished SweetAlert2 error.
     * If rowIndex & fieldName are provided, also highlights the offending input
     * and scrolls the table row into view.
     *
     * @param {string}  title
     * @param {string}  message
     * @param {number}  [rowIndex]    0-based row in cabTableBody
     * @param {string}  [fieldName]   'date'|'time'|'from'|'to'|'amount'|'purpose'
     * @param {Element} [focusEl]     fallback element to focus (side-panel inputs)
     */
    function showValidationError(title, message, rowIndex, fieldName, focusEl) {
        // Highlight table cell if specified
        if (rowIndex !== undefined && fieldName !== undefined) {
            highlightTableField(rowIndex, fieldName);
        }
        // Highlight side-panel input
        if (focusEl) {
            focusEl.classList.add("input-error");
            focusEl.focus();
            focusEl.scrollIntoView({ behavior: "smooth", block: "center" });
            // Auto-clear error on next input
            focusEl.addEventListener("input", () => focusEl.classList.remove("input-error"), { once: true });
        }

        return Swal.fire({
            icon: "error",
            title,
            html: `<span style="font-size:14px;color:#374151;">${message}</span>`,
            confirmButtonColor: "#6366f1",
            confirmButtonText: "Got it",
            customClass: {
                popup: "swal-cab-popup",
                title: "swal-cab-title",
                actions: "swal-cab-actions",
            },
            showClass: { popup: "swal__fadeIn" },
            hideClass: { popup: "swal__fadeOut" },
        });
    }

    /**
     * Map fieldName to the correct input inside a table row and apply error style.
     */
    function highlightTableField(rowIndex, fieldName) {
        const row = els.cabTableBody.rows[rowIndex];
        if (!row) return;

        const fieldMap = { date: 0, time: 1, from: 2, to: 3, amount: 4 };
        let input;

        if (fieldName === "purpose") {
            input = row.children[6]?.querySelector("input:not(.hidden)");
        } else if (fieldMap[fieldName] !== undefined) {
            input = row.children[fieldMap[fieldName] + 1]?.querySelector("input");
        }

        if (!input) return;

        input.classList.add("input-error");
        row.classList.add("row-error");

        // Scroll to that row smoothly
        row.scrollIntoView({ behavior: "smooth", block: "nearest" });

        // Focus after a tick (to let SweetAlert take over)
        setTimeout(() => input.focus(), 600);

        // Auto-clear highlight on fix
        input.addEventListener("input", () => {
            input.classList.remove("input-error");
            row.classList.remove("row-error");
        }, { once: true });
    }

    function showError(title, text) {
        return showValidationError(title, text);
    }

    // ─── GENERATE BUTTON WIRE-UP ──────────────────────────────────────────────
    els.generateBtn.innerHTML = `
        <span class="btn-label"><i class="fa-solid fa-file-pdf"></i>&nbsp; Generate Annexure PDF</span>
        <span class="btn-spinner"></span>`;
    els.generateBtn.addEventListener("click", generatePDF);

    // ─── DRAG & DROP ──────────────────────────────────────────────────────────
    if (els.dropZone) {
        els.dropZone.addEventListener("dragover", e => { e.preventDefault(); els.dropZone.classList.add("drag-over"); });
        els.dropZone.addEventListener("dragleave", () => els.dropZone.classList.remove("drag-over"));
        els.dropZone.addEventListener("drop", e => {
            e.preventDefault();
            els.dropZone.classList.remove("drag-over");
            handleFiles(e.dataTransfer.files);
        });
    }
    els.fileInput.addEventListener("change", e => handleFiles(e.target.files));

    window.addMoreReceipts = () => { els.fileInput.value = ""; els.fileInput.click(); };

    // ESC key closes viewer
    document.addEventListener("keydown", e => {
        if (e.key === "Escape" && !els.pdfDetailView.classList.contains("hidden")) {
            closePdfViewer();
        }
    });

    // ─── SIGNATURE PREVIEW ────────────────────────────────────────────────────
    els.signatureInput.addEventListener("change", e => {
        const file = e.target.files[0];
        if (!file) return;

        if (!["image/jpeg", "image/jpg", "image/png"].includes(file.type)) {
            showError("Invalid Format", "Signature must be JPG or PNG.");
            e.target.value = "";
            return;
        }
        if (file.size > 1024 * 1024) {
            showError("File Too Large", "Signature must be under 1 MB.");
            e.target.value = "";
            return;
        }
        const reader = new FileReader();
        reader.onload = ev => {
            els.signatureImg.src = ev.target.result;
            els.signaturePreview.classList.remove("hidden");
        };
        reader.onerror = () => showError("Read Error", "Could not read signature file.");
        reader.readAsDataURL(file);
    });

    // ─── SORTABLE ─────────────────────────────────────────────────────────────
    new Sortable(els.receiptGrid, {
        animation: 160,
        ghostClass: "drag-ghost",
        chosenClass: "drag-chosen",
        forceFallback: false,
        delay: 200,
        delayOnTouchOnly: true,
        touchStartThreshold: 4,
        onEnd() {
            [...els.receiptGrid.querySelectorAll(".receipt-card")].forEach((card, i) => {
                const r = state.receipts.find(r => r.id === card.dataset.receiptId);
                if (r) r.order = i;
            });
            showToast("Order updated", "info", 1800);
        },
    });

    // ─── NAV HEIGHT HELPER ────────────────────────────────────────────────────
    function getNavHeight() {
        const selectors = ["nav", "header", ".navbar", ".nav-bar", ".site-header", ".top-nav",
            '[role="navigation"]', "#navbar", "#header", "#topbar"];
        for (const sel of selectors) {
            const el = document.querySelector(sel);
            if (el) {
                const r = el.getBoundingClientRect();
                if (r.top <= 4 && r.height > 20) return Math.round(r.bottom);
            }
        }
        return 0;
    }

    // ─── VIEW RECEIPT ─────────────────────────────────────────────────────────
    window.viewReceipt = async function (receiptId) {
        state.activeReceiptId = receiptId;
        state.lastScrollPosition = window.scrollY;

        const receipt = state.receipts.find(r => r.id === receiptId);
        if (!receipt) return;

        const detailView = els.pdfDetailView;
        if (detailView.parentElement !== document.body) document.body.appendChild(detailView);

        const { pdfPagesContainer: pagesContainer, detailFileName: titleEl, pdfPagePill: pagePill, pdfProgressFill: progressFill } = els;

        const navH = getNavHeight();
        document.documentElement.style.setProperty("--pdf-viewer-top", navH + "px");

        document.body.classList.add("pdf-viewer-open");
        detailView.classList.remove("hidden", "closing");
        pagesContainer.scrollTop = 0;

        if (titleEl) titleEl.textContent = receipt.name;
        if (pagePill) pagePill.textContent = "…";
        if (progressFill) progressFill.style.width = "0%";

        pagesContainer.innerHTML = `
            <div class="pdf-loading-msg">
                <i class="fa-solid fa-spinner fa-spin"></i>
                <span>Loading PDF…</span>
            </div>`;

        try {
            const pdf = receipt.pdfDoc;
            const totalPages = pdf.numPages;

            // Two RAF ticks + small delay to let the viewer paint before measuring
            await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
            await new Promise(r => setTimeout(r, 40));

            const cs = window.getComputedStyle(pagesContainer);
            const horizPad = parseFloat(cs.paddingLeft) + parseFloat(cs.paddingRight);
            const vw = document.documentElement.clientWidth;
            const cssPageWidth = Math.min(vw - horizPad - 8, 760);

            const firstPage = await pdf.getPage(1);
            const baseViewport = firstPage.getViewport({ scale: 1 });
            const scale = (cssPageWidth / baseViewport.width) * DPR;

            pagesContainer.innerHTML = "";

            for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
                if (totalPages > 1) {
                    const label = document.createElement("div");
                    label.className = "pdf-page-label";
                    label.textContent = `Page ${pageNum} of ${totalPages}`;
                    pagesContainer.appendChild(label);
                }

                const page = await pdf.getPage(pageNum);
                const viewport = page.getViewport({ scale });

                const canvas = document.createElement("canvas");
                canvas.width = Math.round(viewport.width);
                canvas.height = Math.round(viewport.height);
                canvas.style.width = cssPageWidth + "px";
                canvas.style.animationDelay = (pageNum - 1) * 55 + "ms";
                pagesContainer.appendChild(canvas);

                await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;

                if (progressFill) {
                    progressFill.style.width = Math.round((pageNum / totalPages) * 100) + "%";
                }
            }

            if (pagePill) pagePill.textContent = totalPages + " page" + (totalPages !== 1 ? "s" : "");

            if (progressFill) {
                progressFill.style.width = "100%";
                setTimeout(() => { progressFill.style.width = "0%"; }, 600);
            }
        } catch (err) {
            console.error("PDF render error:", err);
            pagesContainer.innerHTML = `
                <div class="pdf-error-msg">
                    <i class="fa-solid fa-circle-exclamation"></i>
                    Could not load PDF preview.
                    <span>${err.message || ""}</span>
                </div>`;
        }
    };

    // ─── CLOSE PDF VIEWER ─────────────────────────────────────────────────────
    window.closePdfViewer = function () {
        const detailView = $("pdfDetailView");
        const pagesContainer = $("pdfPagesContainer");

        detailView.classList.add("closing");

        const onEnd = () => {
            detailView.removeEventListener("animationend", onEnd);
            detailView.classList.add("hidden");
            detailView.classList.remove("closing");
        };
        detailView.addEventListener("animationend", onEnd);

        document.body.classList.remove("pdf-viewer-open");
        window.scrollTo({ top: state.lastScrollPosition, behavior: "instant" });

        if (state.activeReceiptId) {
            setTimeout(() => {
                const card = document.querySelector(`.receipt-card[data-receipt-id="${state.activeReceiptId}"]`);
                if (card) {
                    card.scrollIntoView({ behavior: "smooth", block: "nearest", inline: "center" });
                    card.style.boxShadow = "0 0 0 3px #6366f1, 0 4px 16px rgba(99,102,241,0.3)";
                    setTimeout(() => (card.style.boxShadow = ""), 1000);
                }
            }, 50);
        }

        setTimeout(() => { pagesContainer.innerHTML = ""; }, 350);
    };

    // ─── DATE PARSER ──────────────────────────────────────────────────────────
    function parseReceiptDate(date, time) {
        if (!date) return new Date(0);
        const [dd, mm, yyyy] = date.split("-");
        let hh = 0, min = 0;
        if (time) {
            const mx = time.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
            if (mx) {
                hh = parseInt(mx[1]);
                min = parseInt(mx[2]);
                if (mx[3].toUpperCase() === "PM" && hh !== 12) hh += 12;
                if (mx[3].toUpperCase() === "AM" && hh === 12) hh = 0;
            }
        }
        return new Date(yyyy, mm - 1, dd, hh, min);
    }

    // ─── HANDLE FILES ─────────────────────────────────────────────────────────
    async function handleFiles(fileList) {
        const files = Array.from(fileList);
        if (!files.length) return;

        // Validate before any heavy work
        for (const file of files) {
            if (file.type !== "application/pdf") {
                showError("Invalid File", `"${file.name}" is not a PDF.`);
                return;
            }
            if (file.size > 1024 * 1024) {
                showError("File Too Large", `"${file.name}" exceeds 1 MB.`);
                return;
            }
        }

        showLoading(`Processing ${files.length} receipt${files.length > 1 ? "s" : ""}…`);

        try {
            // Process sequentially to avoid memory spikes on mobile
            // for (const file of files) await processReceipt(file);
            for (let i = 0; i < files.length; i++) {

                const file = files[i];

                updateLoadingText(`Processing receipt ${i + 1} of ${files.length}...`);

                // optional progress bar
                const percent = Math.round(((i + 1) / files.length) * 100);

                if (els.pdfProgressFill) {
                    els.pdfProgressFill.style.width = percent + "%";
                }

                updateCircleProgress(percent);

                await processReceipt(file);

                // allow UI to repaint (important!)
                await new Promise(requestAnimationFrame);
            }
        } catch (err) {
            hideLoading();
            showError("Processing Failed", err.message || "Could not read one or more PDFs.");
            return;
        }

        // Sort by date/time ascending
        state.receipts.sort((a, b) =>
            parseReceiptDate(a.data.date, a.data.time) -
            parseReceiptDate(b.data.date, b.data.time)
        );
        state.receipts.forEach((r, i) => (r.order = i));

        els.stepUpload.classList.add("hidden");
        els.stepEditor.classList.remove("hidden");
        els.stepEditor.classList.add("flex");

        renderReceiptCards();
        renderTable();
        updateMonthField();
        hideLoading();
        showToast(`${files.length} receipt${files.length > 1 ? "s" : ""} added`, "success");
    }

    // ─── PROCESS RECEIPT ──────────────────────────────────────────────────────
    async function processReceipt(file) {
        let arrayBuffer;
        try { arrayBuffer = await file.arrayBuffer(); }
        catch { throw new Error(`Could not read "${file.name}". Try again.`); }

        let pdf;
        try { pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise; }
        catch { throw new Error(`Could not parse "${file.name}". File may be corrupted.`); }

        let text = "";
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            text += content.items.map(i => i.str).join(" ") + " ";
        }

        const cleanText = text.replace(/\s+/g, " ").replace(/:\s+/g, ":");
        const normalizedText = cleanText
            .replace(/\b(\d{1,2})\s*:\s*(\d{2})\s*(AM|PM|am|pm)\b/g, "$1:$2 $3")
            .replace(/₹\s*((?:\d\s*){1,3})(?=\D)/g, (_, d) => "₹ " + d.replace(/\s+/g, ""));

        if (DEBUG) console.log("normalizedText:", normalizedText);

        const locations = extractLocationsAccurately(normalizedText);
        const isRapido = /selected\s*price/i.test(normalizedText);

        const extractedData = {
            date: extractDate(normalizedText),
            time: extractTime(normalizedText),
            amount: extractAmount(normalizedText),
            fromLoc: isRapido ? locations.toLoc : locations.fromLoc,
            toLoc: isRapido ? locations.fromLoc : locations.toLoc,
        };

        state.receipts.push({
            id: `r-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
            file,
            name: file.name,
            size: (file.size / 1024).toFixed(2) + " KB",
            pdfDoc: pdf,
            data: extractedData,
        });
    }

    // ─── EXTRACTION ───────────────────────────────────────────────────────────
    function extractDate(text) {
        const monthNames = ["January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"];
        const dateRegex = /\b(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}|\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{1,2}(?:st|nd|rd|th)?\s*,?\s*\d{4}|\d{1,2}(?:st|nd|rd|th)?\s*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*-?\s*\d{4})\b/gi;
        const matches = [...text.matchAll(dateRegex)];
        if (!matches.length) return "";

        const wordRegex = /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i;
        const best = matches.find(m => wordRegex.test(m[1])) || matches[0];
        const standardized = standardizeDate(best[1]);
        const parts = standardized.split("-");

        if (parts.length === 3) {
            const mi = parseInt(parts[1], 10) - 1;
            const yr = parts[2];
            if (mi >= 0 && mi < 12) {
                const key = `${monthNames[mi]} ${yr}`;
                state.monthCounts[key] = (state.monthCounts[key] || 0) + 1;
            }
        }
        return standardized;
    }

    function extractTime(text) {
        const timeRegex = /\b(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)?)\b/g;
        const matches = [...text.matchAll(timeRegex)];
        if (!matches.length) return "";
        const withMeridiem = matches.find(m => /AM|PM/i.test(m[1]));
        return standardizeTime((withMeridiem || matches[0])[1]);
    }

    function extractAmount(text) {
        let match = text.match(TOTAL_REGEX);
        if (match) return parseFloat(match[1]);
        match = text.match(RAPIDO_REGEX);
        if (match) return parseFloat(match[1]);
        const all = [...text.matchAll(CURRENCY_REGEX)];
        return all.length ? parseFloat(all[0][1]) : 0;
    }

    function extractLocationsAccurately(text) {
        const normalized = text.replace(/\n/g, " ").replace(/\s+/g, " ").toLowerCase();
        const blocks = normalized.split(/india/i).map(b => b.trim()).filter(b => b.length > 25);
        const results = [];

        for (const block of blocks) {
            let bestScore = 0, bestLocation = "";
            for (const loc of locationKeywords) {
                let score = 0;
                for (const key of loc.keys) {
                    const k = key.toLowerCase();
                    if (block.includes(k)) score += k.length;
                }
                if (score > bestScore) { bestScore = score; bestLocation = loc.short; }
            }
            if (bestLocation) results.push(bestLocation);
        }

        const unique = [...new Map(results.map(l => [l, l])).values()];
        return { fromLoc: unique[0] || "", toLoc: unique[1] || "" };
    }

    // ─── FORMATTING HELPERS ───────────────────────────────────────────────────
    function standardizeDate(raw) {
        if (!raw) return "";
        const clean = raw.replace(/\b(\d{1,2})(st|nd|rd|th)\b/gi, "$1").trim();
        let d, m, y;

        const dmY = clean.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
        const Ymd = clean.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);

        if (dmY) {
            d = parseInt(dmY[1], 10); m = parseInt(dmY[2], 10) - 1; y = parseInt(dmY[3], 10);
            if (d <= 12 && m + 1 > 12) [d, m] = [m + 1, d - 1];
        } else if (Ymd) {
            y = parseInt(Ymd[1], 10); m = parseInt(Ymd[2], 10) - 1; d = parseInt(Ymd[3], 10);
        } else {
            const parsed = new Date(clean);
            if (!isNaN(parsed.getTime())) { d = parsed.getDate(); m = parsed.getMonth(); y = parsed.getFullYear(); }
            else return raw;
        }

        if (y < 100) y += 2000;
        return `${String(d).padStart(2, "0")}-${String(m + 1).padStart(2, "0")}-${y}`;
    }

    function standardizeTime(raw) {
        if (!raw) return "";
        const match = raw.match(/(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)?/i);
        if (!match) return raw;
        let h = parseInt(match[1], 10);
        const mins = match[2];
        let mod = match[3] ? match[3].toUpperCase() : null;
        if (!mod) { mod = h >= 12 ? "PM" : "AM"; if (h > 12) h -= 12; if (h === 0) h = 12; }
        else if (h === 0) h = 12;
        return `${String(h).padStart(2, "0")}:${mins} ${mod}`;
    }

    // ─── RENDER RECEIPT CARDS ─────────────────────────────────────────────────
    function renderReceiptCards() {
        els.receiptGrid.innerHTML = "";
        els.itemCount.textContent = `${state.receipts.length} Receipt${state.receipts.length !== 1 ? "s" : ""}`;

        const frag = document.createDocumentFragment();

        state.receipts.forEach((receipt, index) => {
            const card = document.createElement("div");
            card.className = "receipt-card";
            card.dataset.receiptId = receipt.id;

            card.innerHTML = `
                <div class="receipt-number-badge">${index + 1}</div>
                <button onclick="deleteReceipt('${receipt.id}')" class="btn-delete-receipt" title="Remove">
                    <i class="fa-solid fa-xmark"></i>
                </button>
                <button onclick="viewReceipt('${receipt.id}')" class="btn-view-receipt" title="View PDF">
                    <i class="fa-solid fa-eye"></i>
                </button>
                <div class="receipt-thumb loading">
                    <canvas class="receipt-canvas"></canvas>
                </div>
                <div class="receipt-info">
                    <div class="receipt-name" title="${receipt.name}">${receipt.name}</div>
                </div>`;

            frag.appendChild(card);
        });

        els.receiptGrid.appendChild(frag);

        // Render thumbnails lazily via requestIdleCallback (or fallback rAF)
        const scheduleThumb = window.requestIdleCallback
            ? cb => requestIdleCallback(cb, { timeout: 800 })
            : cb => requestAnimationFrame(cb);

        state.receipts.forEach((receipt, index) => {
            scheduleThumb(() => {
                const card = els.receiptGrid.querySelectorAll(".receipt-card")[index];
                const thumb = card?.querySelector(".receipt-thumb");
                const canvas = card?.querySelector("canvas");
                if (!canvas) return;
                renderThumbnail(receipt.pdfDoc, 1, canvas).then(() => {
                    thumb?.classList.remove("loading");
                });
            });
        });
    }

    async function renderThumbnail(pdfDoc, pageNum, canvas) {
        try {
            const page = await pdfDoc.getPage(pageNum);
            const THUMB = 84;
            const base = page.getViewport({ scale: 1 });
            const vp = page.getViewport({ scale: THUMB / base.width });

            canvas.width = vp.width;
            canvas.height = vp.height;

            await page.render({ canvasContext: canvas.getContext("2d"), viewport: vp }).promise;
        } catch {
            canvas.width = 84; canvas.height = 112;
            const ctx = canvas.getContext("2d");
            ctx.fillStyle = "#e5e7eb";
            ctx.fillRect(0, 0, 84, 112);
            ctx.fillStyle = "#9ca3af";
            ctx.font = "10px sans-serif";
            ctx.textAlign = "center";
            ctx.fillText("No preview", 42, 56);
        }
    }

    // ─── RENDER TABLE ─────────────────────────────────────────────────────────
    function renderTable() {
        els.cabTableBody.innerHTML = "";
        let total = 0;
        const frag = document.createDocumentFragment();

        state.receipts.forEach((receipt, index) => {
            const amount = receipt.data.amount;
            total += amount;

            const row = document.createElement("tr");
            row.innerHTML = `
                <td style="color:#9ca3af;font-size:12px;">${index + 1}</td>
                <td><input type="text"   placeholder="DD-MM-YYYY"  class="table-input" value="${receipt.data.date || ""}"></td>
                <td><input type="text"   placeholder="HH:MM AM/PM" class="table-input" value="${receipt.data.time || ""}"></td>
                <td><input type="text"   placeholder="From location" class="table-input" value="${receipt.data.fromLoc || ""}"></td>
                <td><input type="text"   placeholder="To location"   class="table-input" value="${receipt.data.toLoc || ""}"></td>
                <td><input type="number" placeholder="0.00" class="table-input"
                        value="${amount > 0 ? amount.toFixed(2) : ""}"
                        min="0" step="0.01" oninput="updateTotal()"></td>
                <td>
                    <select onchange="togglePurpose(this)" class="table-select">
                        <option value="Official">Official</option>
                        <option value="Other">Other</option>
                    </select>
                    <input type="text" placeholder="Enter purpose" class="table-input hidden" style="margin-top:4px;">
                </td>`;
            frag.appendChild(row);
        });

        els.cabTableBody.appendChild(frag);
        els.uiTotalAmount.textContent = "Rs. " + total.toFixed(2);
    }

    // ─── DELETE RECEIPT ───────────────────────────────────────────────────────
    window.deleteReceipt = function (receiptId) {
        const idx = state.receipts.findIndex(r => r.id === receiptId);
        if (idx === -1) return;
        const name = state.receipts[idx].name;
        state.receipts.splice(idx, 1);

        if (!state.receipts.length) {
            els.stepEditor.classList.add("hidden");
            els.stepUpload.classList.remove("hidden");
            state.monthCounts = {};
        } else {
            state.receipts.forEach((r, i) => (r.order = i));
            renderReceiptCards();
            renderTable();
            updateMonthField();
        }
        showToast(`Removed "${name}"`, "info", 2500);
    };

    // ─── MONTH AUTO-DETECT ────────────────────────────────────────────────────
    function updateMonthField() {
        let best = "", max = 0;
        for (const [month, count] of Object.entries(state.monthCounts)) {
            if (count > max) { max = count; best = month; }
        }
        if (best) els.monthInput.value = best;
    }

    // ─── TOGGLE PURPOSE ───────────────────────────────────────────────────────
    window.togglePurpose = function (select) {
        const input = select.nextElementSibling;
        input.classList.toggle("hidden", select.value !== "Other");
    };

    // ─── REMOVE SIGNATURE ────────────────────────────────────────────────────
    window.removeSignature = function () {
        els.signatureInput.value = "";
        els.signatureImg.src = "";
        els.signaturePreview.classList.add("hidden");
    };

    // ─── DEBOUNCED TOTAL UPDATE ───────────────────────────────────────────────
    let totalDebounceTimer = null;
    window.updateTotal = function () {
        clearTimeout(totalDebounceTimer);
        totalDebounceTimer = setTimeout(() => {
            let total = 0;
            Array.from(els.cabTableBody.rows).forEach(tr => {
                total += parseFloat(tr.children[5]?.querySelector("input")?.value) || 0;
            });
            els.uiTotalAmount.textContent = "Rs. " + total.toFixed(2);
        }, 80); // 80 ms debounce — imperceptible but prevents per-keystroke thrash
    };

    // ─── HELPERS ──────────────────────────────────────────────────────────────
    function getFirstName(name) {
        const first = (name || "").trim().split(/[\s_]+/)[0];
        return first.charAt(0).toUpperCase() + first.slice(1).toLowerCase();
    }

    function fileToBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    }

    function triggerDownload(url, filename) {
        const a = document.createElement("a");
        a.href = url; a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    // Loading overlay helpers
    function showLoading(msg) {
        if (els.loadingText) els.loadingText.textContent = msg || "Processing…";
        els.loadingOverlay.classList.remove("hidden");
        els.loadingOverlay.classList.add("flex");

        updateCircleProgress(0); // reset progress
        setGenerateBtnLoading(true);
    }
    function updateLoadingText(msg) {
        if (els.loadingText) els.loadingText.textContent = msg;
    }

    function hideLoading() {
        els.loadingOverlay.classList.add("hidden");
        els.loadingOverlay.classList.remove("flex");
        setGenerateBtnLoading(false);
    }

    function setGenerateBtnLoading(on) {
        els.generateBtn.disabled = on;
        els.generateBtn.classList.toggle("loading", on);
    }

    // ─── GENERATE PDF ─────────────────────────────────────────────────────────
    async function generatePDF() {
        if (!state.receipts.length) {
            showValidationError("No Receipts", "Please upload at least one cab receipt PDF.");
            return;
        }

        const nameVal = els.empNameInput.value.trim();
        if (!nameVal) {
            showValidationError("Missing Name", "Please enter the employee name.", undefined, undefined, els.empNameInput);
            return;
        }
        const desigVal = els.designationInput.value.trim();
        if (!desigVal) {
            showValidationError("Missing Designation", "Please enter the designation.", undefined, undefined, els.designationInput);
            return;
        }

        const dateRegex = /^\d{2}-\d{2}-\d{4}$/;
        const timeRegex = /^\d{2}:\d{2}\s*(AM|PM)$/i;
        let validationError = null;
        let errorRow = null;
        let errorField = null;

        Array.from(els.cabTableBody.rows).every((tr, i) => {
            const rowNum = i + 1;
            const inputs = tr.querySelectorAll("input");
            const date = inputs[0].value.trim();
            const time = inputs[1].value.trim();
            const from = inputs[2].value.trim();
            const to = inputs[3].value.trim();
            const amount = parseFloat(inputs[4].value);
            const purposeSelect = tr.children[6].querySelector("select");
            const purposeInput = tr.children[6].querySelector("input");

            if (!date) {
                validationError = `Row #${rowNum}: Date is missing.`;
                errorRow = i; errorField = "date"; return false;
            }
            if (!dateRegex.test(date)) {
                validationError = `Row #${rowNum}: Date "${date}" is invalid — use DD-MM-YYYY.`;
                errorRow = i; errorField = "date"; return false;
            }
            if (!time) {
                validationError = `Row #${rowNum}: Time is missing.`;
                errorRow = i; errorField = "time"; return false;
            }
            if (!timeRegex.test(time)) {
                validationError = `Row #${rowNum}: Time "${time}" is invalid — use HH:MM AM/PM.`;
                errorRow = i; errorField = "time"; return false;
            }
            if (!from) {
                validationError = `Row #${rowNum}: "From" location is missing.`;
                errorRow = i; errorField = "from"; return false;
            }
            if (!to) {
                validationError = `Row #${rowNum}: "To" location is missing.`;
                errorRow = i; errorField = "to"; return false;
            }
            if (!amount || amount <= 0) {
                validationError = `Row #${rowNum}: Amount must be greater than 0.`;
                errorRow = i; errorField = "amount"; return false;
            }
            if (purposeSelect.value === "Other" && !purposeInput.value.trim()) {
                validationError = `Row #${rowNum}: Custom purpose is required when "Other" is selected.`;
                errorRow = i; errorField = "purpose"; return false;
            }
            return true;
        });

        if (validationError) {
            showValidationError("Incomplete Details", validationError, errorRow, errorField);
            return;
        }

        try {
            showLoading("Generating annexure PDF…");

            const { jsPDF } = window.jspdf;
            const doc = new jsPDF("p", "mm", "a4");
            const pageW = doc.internal.pageSize.getWidth();   // 210 mm
            const pageH = doc.internal.pageSize.getHeight();  // 297 mm
            const LEFT = 14;   // left margin
            const RIGHT = 14;   // right margin
            const BOTTOM = 18;   // bottom safe margin
            const USABLE = pageW - LEFT - RIGHT;  // 182 mm usable width

            const month = (els.monthInput.value || "").trim().replace(/\s+/g, "_");
            const firstName = getFirstName(nameVal);

            // ── Header month badge ──
            doc.setFont("helvetica", "bold");
            const monthText = `Month: ${els.monthInput.value || "________"}`;
            doc.setFontSize(10);
            doc.setFillColor(255, 255, 0);
            doc.rect(LEFT - 1, 9, doc.getTextWidth(monthText) + 4, 6, "F");
            doc.setTextColor(0, 0, 0);
            doc.text(monthText, LEFT, 14);

            // ── Build table rows ──
            const rows = [];
            let pdfTotal = 0;

            Array.from(els.cabTableBody.rows).forEach(tr => {
                const tds = tr.children;
                const parsed = parseFloat(tds[5].querySelector("input").value) || 0;
                pdfTotal += parsed;

                const purposeSelect = tds[6].querySelector("select");
                const purposeInput = tds[6].querySelector("input");
                const purpose = purposeSelect.value === "Official"
                    ? "Official"
                    : (purposeInput.value.trim() || "Other");

                rows.push([
                    tds[0].textContent.trim(),
                    tds[1].querySelector("input").value.trim(),
                    tds[2].querySelector("input").value.trim(),
                    tds[3].querySelector("input").value.trim(),
                    tds[4].querySelector("input").value.trim(),
                    parsed > 0 ? "Rs. " + parsed.toFixed(2) : "-",
                    purpose,
                ]);
            });

            // ── Column widths (must sum to USABLE = 182 mm) ──
            // S.No(10) + Date(25) + Time(22) + From(37) + To(37) + Amount(27) + Purpose(24) = 182
            const colW = { sno: 10, date: 25, time: 22, from: 37, to: 37, amount: 27, purpose: 24 };

            const cellPad = { top: 3.5, bottom: 3.5, left: 3, right: 3 };

            doc.autoTable({
                startY: 22,
                margin: { left: LEFT, right: RIGHT },
                tableWidth: USABLE,
                head: [["S. No.", "Date", "Time", "From", "To", "Bill Amount", "Purpose"]],
                body: rows,
                theme: "grid",
                columnStyles: {
                    0: { cellWidth: colW.sno, halign: "center" },
                    1: { cellWidth: colW.date, halign: "center" },
                    2: { cellWidth: colW.time, halign: "center" },
                    3: { cellWidth: colW.from, halign: "left" },
                    4: { cellWidth: colW.to, halign: "left" },
                    5: { cellWidth: colW.amount, halign: "right" },
                    6: { cellWidth: colW.purpose, halign: "center" },
                },
                headStyles: {
                    fillColor: [240, 240, 240],
                    textColor: [0, 0, 0],
                    fontStyle: "bold",
                    halign: "center",
                    fontSize: 8.5,
                    lineWidth: 0.3,
                    lineColor: [180, 180, 180],
                    cellPadding: cellPad,
                },
                bodyStyles: {
                    textColor: [30, 30, 30],
                    fontSize: 8.5,
                    lineWidth: 0.2,
                    lineColor: [180, 180, 180],
                    cellPadding: cellPad,
                    valign: "middle",
                    overflow: "linebreak",
                },
                alternateRowStyles: {
                    fillColor: [250, 250, 252],
                },
                styles: {
                    font: "helvetica",
                    overflow: "linebreak",
                },
                // Repeat header on every new page
                showHead: "everyPage",
                // Add page number in footer of each page
                didDrawPage: (data) => {
                    const pgNum = doc.internal.getNumberOfPages();
                    doc.setFont("helvetica", "normal");
                    doc.setFontSize(8);
                    doc.setTextColor(150, 150, 150);
                    doc.text(
                        `Page ${pgNum}`,
                        pageW / 2, pageH - 8,
                        { align: "center" }
                    );
                    doc.setTextColor(0, 0, 0);
                },
            });

            let y = doc.lastAutoTable.finalY + 10;

            // ── Total badge ──
            const totalText = `Total Amount: Rs. ${pdfTotal.toFixed(2)}`;
            doc.setFontSize(10);
            doc.setFont("helvetica", "bold");

            // Estimate space needed for the footer block:
            // total(8) + gap(10) + decl_heading(6) + decl_body(16) + gap(8) + sig(18) + name(8) + desig(6) = ~80mm
            const FOOTER_HEIGHT = 82;
            if (y + FOOTER_HEIGHT > pageH - BOTTOM) {
                doc.addPage();
                y = 20;
            }

            doc.setFillColor(255, 255, 0);
            doc.rect(LEFT - 1, y - 4.5, doc.getTextWidth(totalText) + 6, 6.5, "F");
            doc.setTextColor(0, 0, 0);
            doc.text(totalText, LEFT, y);
            y += 12;

            // ── Declaration ──
            doc.setFontSize(10);
            doc.setFont("helvetica", "bold");
            doc.text("Declaration:", LEFT, y);
            y += 6;

            doc.setFont("helvetica", "normal");
            doc.setFontSize(9);
            const declaration =
                "I have not availed the cab services for any personal purposes and all the bills " +
                "submitted are true and original. I claim full responsibility for the details furnished in this annexure.";

            // Split manually so we know exact line count & height
            const declLines = doc.splitTextToSize(declaration, USABLE);
            const lineH = 5;   // 9pt at standard leading ≈ 5mm per line
            doc.text(declLines, LEFT, y);
            y += declLines.length * lineH + 10;

            // ── Signature & Name block ──
            const SIG_W = 42;
            const SIG_H = 16;
            const NAME_X = pageW - RIGHT;   // right-aligned to margin

            const signFile = els.signatureInput.files[0];
            if (signFile) {
                const imgData = await fileToBase64(signFile);
                const format = signFile.type === "image/png" ? "PNG" : "JPEG";
                doc.addImage(imgData, format, NAME_X - SIG_W, y, SIG_W, SIG_H);
                y += SIG_H + 3;
            } else {
                // blank line above name when no signature
                y += 6;
            }

            // Thin underline for signature area
            doc.setDrawColor(180, 180, 180);
            doc.setLineWidth(0.3);
            doc.line(NAME_X - SIG_W - 4, y - 1, NAME_X, y - 1);

            doc.setFont("helvetica", "bold");
            doc.setFontSize(9);
            doc.setTextColor(0, 0, 0);
            doc.text(nameVal, NAME_X, y + 4, { align: "right" });

            doc.setFont("helvetica", "normal");
            doc.setFontSize(8.5);
            doc.setTextColor(80, 80, 80);
            // Wrap designation if long
            const desigLines = doc.splitTextToSize(desigVal, SIG_W + 4);
            doc.text(desigLines, NAME_X, y + 10, { align: "right" });

            // ── Merge receipts ──
            updateLoadingText("Merging receipt PDFs…");
            const annexureBlob = doc.output("blob");
            const receiptsBlob = await generateMergedReceiptsPDF();

            // ── ZIP ──
            updateLoadingText("Creating ZIP archive…");
            const zip = new JSZip();
            zip.file(`${firstName}_Cab_Reimbursement_${month}.pdf`, annexureBlob);
            zip.file(`${firstName}_Cab_Bills_Merged_${month}.pdf`, receiptsBlob);
            const zipBlob = await zip.generateAsync({ type: "blob" });
            const zipName = `${firstName}_Cab_Reimbursement_${month}.zip`;

            const dlBtn = $("downloadBtn");
            if (dlBtn) dlBtn.onclick = () => triggerDownload(URL.createObjectURL(zipBlob), zipName);

            els.stepEditor.classList.add("hidden");
            els.stepSuccess.classList.remove("hidden");
            hideLoading();
            showToast("PDF generated successfully!", "success", 4000);

        } catch (err) {
            console.error("PDF generation error:", err);
            hideLoading();
            showError("Generation Failed", err.message || "An unexpected error occurred.");
        }
    }

    // ─── MERGE RECEIPTS PDF ───────────────────────────────────────────────────
    async function generateMergedReceiptsPDF() {
        const { PDFDocument } = PDFLib;
        const merged = await PDFDocument.create();

        for (const r of state.receipts) {
            const bytes = await r.file.arrayBuffer();
            const pdf = await PDFDocument.load(bytes);
            const pages = await merged.copyPages(pdf, pdf.getPageIndices());
            pages.forEach(p => merged.addPage(p));
        }

        return new Blob([await merged.save()], { type: "application/pdf" });
    }

    // ─── PAGE CACHE / RESET ───────────────────────────────────────────────────
    window.addEventListener("pageshow", e => {
        if (e.persisted || window.performance?.navigation?.type === 2) {
            window.location.reload();
        }
    });

    if (window.performance?.navigation?.type !== 1) {
        [els.monthInput, els.empNameInput, els.designationInput].forEach(el => { if (el) el.value = ""; });
        els.signatureInput.value = "";
        els.signatureImg.src = "";
        els.signaturePreview.classList.add("hidden");
        els.receiptGrid.innerHTML = "";
        els.cabTableBody.innerHTML = "";
        els.itemCount.textContent = "0 Receipts";
        els.uiTotalAmount.textContent = "Rs. 0.00";
        els.stepEditor.classList.add("hidden");
        els.stepSuccess?.classList.add("hidden");
        els.stepUpload.classList.remove("hidden");
    }

    function updateCircleProgress(percent) {

        const circle = document.querySelector(".progress-ring-circle");
        if (!circle) return;

        const offset = circumference - (percent / 100) * circumference;
        circle.style.strokeDashoffset = offset;
    }
});