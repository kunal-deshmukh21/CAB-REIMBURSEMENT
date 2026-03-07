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
        // Track active Flatpickr instances so we can destroy on re-render
        datePickerInstances: [],
        timePickerInstances: [],
    };

    // ─── CONSTANTS ────────────────────────────────────────────────────────────
    const DPR = Math.min(window.devicePixelRatio || 1, 2);
    const TOTAL_REGEX = /Total\s*(?:₹|Rs\.?|INR)?\s*(\d+(?:\.\d{1,2})?)/i;
    const RAPIDO_REGEX = /Selected\s*Price\s*(?:₹|Rs\.?|INR)?\s*(\d+(?:\.\d{1,2})?)/i;
    const CURRENCY_REGEX = /(?:₹|Rs\.?|INR)\s*(\d+(?:\.\d{1,2})?)/g;

    const MAX_PDF_FILES = 50;
    const MAX_PDF_SIZE = 1024 * 1024; // 1 MB
    const MAX_SIGNATURE_SIZE = 2 * 1024 * 1024; // 2 MB

    const NAME_MAX = 80;
    const LOCATION_MAX = 120;
    const PURPOSE_MAX = 120;

    const MAX_BILL_AMOUNT = 100000; // 1 lakh

    const MONTH_NAMES = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];

    /**
     * Parse "Month YYYY" → { monthIndex: 0-11, year: YYYY }
     * Returns null when the string is not yet a valid month/year.
     */
    function parseMonthYear(value) {
        const match = (value || "").match(
            /^(January|February|March|April|May|June|July|August|September|October|November|December)\s(\d{4})$/i
        );
        if (!match) return null;
        const monthIndex = MONTH_NAMES.findIndex(
            m => m.toLowerCase() === match[1].toLowerCase()
        );
        return { monthIndex, year: parseInt(match[2], 10) };
    }

    /**
     * Apply minDate / maxDate to every date picker based on the current
     * Month field value.  Called whenever the month field changes AND
     * right after date pickers are initialised.
     *
     * If the month field is blank / invalid → remove any range restriction.
     * If the currently selected date falls outside the new range → clear it
     * and show a toast so the user knows.
     */
    function updateDatePickerRanges() {
        const parsed = parseMonthYear(els.monthInput.value.trim());

        state.datePickerInstances.forEach(fp => {
            if (parsed) {
                const firstDay = new Date(parsed.year, parsed.monthIndex, 1);
                const lastDay = new Date(parsed.year, parsed.monthIndex + 1, 0);

                fp.set("minDate", firstDay);
                fp.set("maxDate", lastDay);

                // Jump the calendar view to the selected month
                fp.jumpToDate(firstDay, false);

                // If a date is already selected but now out of range → clear it
                if (fp.selectedDates.length) {
                    const sel = fp.selectedDates[0];
                    if (sel < firstDay || sel > lastDay) {
                        fp.clear();
                        showToast("Date cleared — outside selected month", "info", 3000);
                    }
                }
            } else {
                // No valid month → lift restrictions
                fp.set("minDate", null);
                fp.set("maxDate", null);
            }
        });
    }

    function isValidDate(dateStr) {

        if (!/^\d{2}-\d{2}-\d{4}$/.test(dateStr)) return false;

        const [dd, mm, yyyy] = dateStr.split("-").map(Number);

        if (yyyy < 2000 || yyyy > 2100) return false;
        if (mm < 1 || mm > 12) return false;

        const days = new Date(yyyy, mm, 0).getDate();

        if (dd < 1 || dd > days) return false;

        return true;
    }

    function isValidTime(timeStr) {

        // Accept both "02:30 PM" (manually typed) and "2:30 PM" (Flatpickr h format)
        const match = timeStr.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
        if (!match) return false;

        const hh = parseInt(match[1]);
        const mm = parseInt(match[2]);

        if (hh < 1 || hh > 12) return false;
        if (mm < 0 || mm > 59) return false;

        return true;
    }
    function isValidLocation(loc) {
        return loc && loc.length >= 3 && loc.length <= LOCATION_MAX;
    }

    function isValidName(name) {

        if (!name) return false;

        if (name.length > NAME_MAX) return false;

        return /^[a-zA-Z\s._-]+$/.test(name);
    }

    function isValidMonthYear(value) {

        const match = value.match(/^(January|February|March|April|May|June|July|August|September|October|November|December)\s\d{4}$/i);

        return !!match;
    }

    const locationKeywords = [
        { keys: ["Jorbagh", "Sri Aurobindo Marg", "Jor Bagh", "Safdarjung Airport", "Safdarjung", "safdarjung Airport Area", "Satya Sadan"], short: "Jor Bagh" },
        { keys: ["Rajaji Marg, Vijay Chowk Area", "Vijay Chowk Area", "Sena Bhawan", "Sena Bhavan", "Central Secretariat"], short: "Sena Bhawan" },
        { keys: ["Electronics Niketan", "CGO complex"], short: "Electronics Niketan" },
        { keys: ["Scope Complex"], short: "Scope Complex" },
        { keys: ["Punjabi Bagh", "Punjabi Bagh Enclave", "West Punjabi Bagh"], short: "Punjabi Bagh" },
        { keys: ["GPO Complex", "Barapullah Rd", "Aviation Colony", "INA Colony", "Ayush Bhavan", "Ayush Bhawan"], short: "Ayush Bhawan" },
        { keys: ["INA Metro Station"], short: "INA Metro Station" },
        { keys: ["Directorate General of Information Systems", "Shankar Vihar"], short: "DGIS" },
        { keys: ["Central Secretariat", "Udyog Bhawan"], short: "Central Secretariat" },
        { keys: ["Subroto Park"], short: "Subroto Park" },
        { keys: ["Shalimar Bagh"], short: "Shalimar Bagh" },
        { keys: ["IGNOU", "Indira Gandhi National Open University"], short: "IGNOU" },
        { keys: ["Defence Colony"], short: "Defence Colony" },
        { keys: ["KG Marg", "KG M arg"], short: "KG Marg" },
        { keys: ["Nausena Bhawan", "Nausena Bhavan", "Nau Sena Bhawan", "NCN Centre", "Near Raksha Sampda Bhawan"], short: "Nau Sena Bhawan" },
        { keys: ["DCN Palam"], short: "DCN Palam" },
        { keys: ["Rajaji Marg", "Meena Bagh", "Krishna Manon Lane Area","HQIDS Kashmir House","Kashmir House"], short: "Kashmir House" },
        { keys: ["UPSC Bhavan", "UPSC Bhawan", "Man Singh Road Area","UPSC"], short: "UPSC" },


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
    function showValidationError(title, message, rowIndex, fieldName, focusEl) {
        if (rowIndex !== undefined && fieldName !== undefined) {
            highlightTableField(rowIndex, fieldName);
        }
        if (focusEl) {
            focusEl.classList.add("input-error");
            focusEl.focus();
            focusEl.scrollIntoView({ behavior: "smooth", block: "center" });
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

        row.scrollIntoView({ behavior: "smooth", block: "nearest" });

        setTimeout(() => input.focus(), 600);

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
        const allowed = ["image/jpeg", "image/jpg", "image/png", "image/webp"];

        if (!allowed.includes(file.type)) {
            showError("Invalid Format", "Signature must be JPG, PNG or WEBP.");
            return;
        }

        if (file.size > MAX_SIGNATURE_SIZE) {
            showError("File Too Large", "Signature must be under 2MB.");
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

        if (files.length > MAX_PDF_FILES) {
            showError("Too Many Files", "Maximum 50 PDFs allowed at once.");
            return;
        }

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
            for (let i = 0; i < files.length; i++) {

                const file = files[i];

                updateLoadingText(`Processing receipt ${i + 1} of ${files.length}...`);

                const percent = Math.round(((i + 1) / files.length) * 100);

                if (els.pdfProgressFill) {
                    els.pdfProgressFill.style.width = percent + "%";
                }

                updateCircleProgress(percent);

                await processReceipt(file);

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
        // Destroy any existing Flatpickr instances before clearing the DOM
        destroyDateTimePickers();

        els.cabTableBody.innerHTML = "";
        let total = 0;
        const frag = document.createDocumentFragment();

        state.receipts.forEach((receipt, index) => {
            const amount = receipt.data.amount;
            total += amount;

            const row = document.createElement("tr");
            row.innerHTML = `
                <td style="color:#9ca3af;font-size:12px;">${index + 1}</td>
                <td>
                    <div class="picker-cell">
                        <input type="text" placeholder="DD-MM-YYYY" class="table-input date-input"
                            value="${receipt.data.date || ""}" autocomplete="off" readonly>
                        <i class="fa-regular fa-calendar picker-icon date-icon"></i>
                    </div>
                </td>
                <td>
                    <div class="picker-cell">
                        <input type="text" placeholder="HH:MM AM/PM" class="table-input time-input"
                            value="${receipt.data.time || ""}" autocomplete="off" readonly>
                        <i class="fa-regular fa-clock picker-icon time-icon"></i>
                    </div>
                </td>
                <td><input maxlength="120" type="text" placeholder="From location" class="table-input" value="${receipt.data.fromLoc || ""}"></td>
                <td><input maxlength="120" type="text" placeholder="To location"   class="table-input" value="${receipt.data.toLoc || ""}"></td>
                <td><input type="number" placeholder="0.00" class="table-input"
                        value="${amount > 0 ? amount.toFixed(2) : ""}"
                        min="0" step="0.01" oninput="updateTotal()"></td>
                <td>
                    <select onchange="togglePurpose(this)" class="table-select">
                        <option value="Official">Official</option>
                        <option value="Other">Other</option>
                    </select>
                    <input maxlength="120" type="text" placeholder="Enter purpose" class="table-input hidden" style="margin-top:4px;">
                </td>`;
            frag.appendChild(row);
        });

        els.cabTableBody.appendChild(frag);
        els.uiTotalAmount.textContent = "Rs. " + total.toFixed(2);

        // Initialise pickers after DOM is ready
        requestAnimationFrame(() => initDateTimePickers());
    }

    // ─── DATE / TIME PICKERS ──────────────────────────────────────────────────

    /**
     * Destroy all tracked Flatpickr instances (called before re-render).
     */
    function destroyDateTimePickers() {
        state.datePickerInstances.forEach(fp => { try { fp.destroy(); } catch (_) { } });
        state.timePickerInstances.forEach(fp => { try { fp.destroy(); } catch (_) { } });
        state.datePickerInstances = [];
        state.timePickerInstances = [];
    }

    /**
     * Attach Flatpickr to every date-input and time-input in the table.
     * Called once after the table is rendered.
     */
    function initDateTimePickers() {
        // Pre-compute month range so the calendar opens on the right month
        const parsedMonth = parseMonthYear(els.monthInput.value.trim());
        const rangeMin = parsedMonth
            ? new Date(parsedMonth.year, parsedMonth.monthIndex, 1)
            : new Date(2000, 0, 1);
        const rangeMax = parsedMonth
            ? new Date(parsedMonth.year, parsedMonth.monthIndex + 1, 0)
            : new Date(2100, 11, 31);

        // ── Date pickers ──────────────────────────────────────────────────────
        els.cabTableBody.querySelectorAll(".date-input").forEach(input => {
            const fp = flatpickr(input, {
                dateFormat: "d-m-Y",
                allowInput: true,
                disableMobile: true,
                todayHighlight: true,
                closeOnSelect: true,
                minDate: rangeMin,
                maxDate: rangeMax,
                // NOTE: no defaultDate here — that would overwrite extracted values.
                // Instead we jump the *view* (not the value) when the calendar opens.
                onOpen() {
                    // If no date is selected yet, scroll the calendar view to the
                    // selected month so the user doesn't have to navigate manually.
                    if (!fp.selectedDates.length && parsedMonth) {
                        fp.jumpToDate(rangeMin, false);
                    }
                },
                onClose(selectedDates) {
                    if (selectedDates.length) {
                        input.dispatchEvent(new Event("input", { bubbles: true }));
                    }
                },
            });
            state.datePickerInstances.push(fp);
        });

        // ── Time pickers ──────────────────────────────────────────────────────
        els.cabTableBody.querySelectorAll(".time-input").forEach(input => {
            const fp = flatpickr(input, {
                // Time-only mode
                enableTime: true,
                noCalendar: true,
                // 12-hour clock with AM/PM → "02:30 PM"
                time_24hr: false,
                // Output format → HH:MM AM/PM (matches isValidTime regex)
                dateFormat: "h:i K",
                allowInput: true,
                disableMobile: true,
                // 5-minute steps for convenience
                minuteIncrement: 5,
                onClose(selectedDates) {
                    if (selectedDates.length) {
                        input.dispatchEvent(new Event("input", { bubbles: true }));
                    }
                },
            });
            state.timePickerInstances.push(fp);
        });

        // Apply month range immediately so existing extracted dates are validated
        updateDatePickerRanges();
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
            destroyDateTimePickers();
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
        if (best) {
            els.monthInput.value = best;
            // Sync the calendar range to the newly detected month
            updateDatePickerRanges();
        }
    }

    // Re-apply range whenever the user manually edits the Month field
    let monthDebounceTimer = null;
    els.monthInput.addEventListener("input", () => {
        clearTimeout(monthDebounceTimer);
        monthDebounceTimer = setTimeout(() => {
            updateDatePickerRanges();
            // Visual feedback when a valid month is entered
            if (parseMonthYear(els.monthInput.value.trim())) {
                showToast(`Calendar locked to ${els.monthInput.value.trim()}`, "info", 2200);
            }
        }, 400);
    });

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
        }, 80);
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

    function showLoading(msg) {
        if (els.loadingText) els.loadingText.textContent = msg || "Processing…";
        els.loadingOverlay.classList.remove("hidden");
        els.loadingOverlay.classList.add("flex");
        updateCircleProgress(0);
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
        if (!isValidName(nameVal)) {
            showValidationError(
                "Invalid Name",
                "Please enter a valid employee name (letters only).",
                undefined, undefined, els.empNameInput
            );
            return;
        }

        const monthVal = els.monthInput.value.trim();
        if (!isValidMonthYear(monthVal)) {
            showValidationError(
                "Invalid Month",
                "Month must be like 'January 2026'.",
                undefined, undefined, els.monthInput
            );
            return;
        }

        const desigVal = els.designationInput.value.trim();
        if (!desigVal) {
            showValidationError("Missing Designation", "Please enter the designation.", undefined, undefined, els.designationInput);
            return;
        }

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
            if (!isValidDate(date)) {
                validationError = `Row #${rowNum}: Invalid date. Use DD-MM-YYYY.`;
                errorRow = i; errorField = "date"; return false;
            }
            if (!time) {
                validationError = `Row #${rowNum}: Time is missing.`;
                errorRow = i; errorField = "time"; return false;
            }
            if (!isValidTime(time)) {
                validationError = `Row #${rowNum}: Invalid time. Use HH:MM AM/PM.`;
                errorRow = i; errorField = "time"; return false;
            }
            if (!from) {
                validationError = `Row #${rowNum}: "From" location is missing.`;
                errorRow = i; errorField = "from"; return false;
            }
            if (!isValidLocation(from)) {
                validationError = `Row #${rowNum}: From location must be 3-120 characters.`;
                errorRow = i; errorField = "from"; return false;
            }
            if (!to) {
                validationError = `Row #${rowNum}: "To" location is missing.`;
                errorRow = i; errorField = "to"; return false;
            }
            if (!isValidLocation(to)) {
                validationError = `Row #${rowNum}: To location must be 3-120 characters.`;
                errorRow = i; errorField = "to"; return false;
            }
            if (!amount || amount <= 0 || amount > MAX_BILL_AMOUNT) {
                validationError = `Row #${rowNum}: Amount must be between ₹1 and ₹100000.`;
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
            const pageW = doc.internal.pageSize.getWidth();
            const pageH = doc.internal.pageSize.getHeight();
            const LEFT = 14;
            const RIGHT = 14;
            const BOTTOM = 18;
            const USABLE = pageW - LEFT - RIGHT;

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
                alternateRowStyles: { fillColor: [250, 250, 252] },
                styles: { font: "helvetica", overflow: "linebreak" },
                showHead: "everyPage",
                didDrawPage: (data) => {
                    const pgNum = doc.internal.getNumberOfPages();
                    doc.setFont("helvetica", "normal");
                    doc.setFontSize(8);
                    doc.setTextColor(150, 150, 150);
                    doc.text(`Page ${pgNum}`, pageW / 2, pageH - 8, { align: "center" });
                    doc.setTextColor(0, 0, 0);
                },
            });

            let y = doc.lastAutoTable.finalY + 10;

            const totalText = `Total Amount: Rs. ${pdfTotal.toFixed(2)}`;
            doc.setFontSize(10);
            doc.setFont("helvetica", "bold");

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

            doc.setFontSize(10);
            doc.setFont("helvetica", "bold");
            doc.text("Declaration:", LEFT, y);
            y += 6;

            doc.setFont("helvetica", "normal");
            doc.setFontSize(9);
            const declaration =
                "I have not availed the cab services for any personal purposes and all the bills " +
                "submitted are true and original. I claim full responsibility for the details furnished in this annexure.";

            const declLines = doc.splitTextToSize(declaration, USABLE);
            const lineH = 5;
            doc.text(declLines, LEFT, y);
            y += declLines.length * lineH + 10;

            const SIG_W = 42;
            const SIG_H = 16;
            const RIGHT_X = pageW - RIGHT;

            const signFile = els.signatureInput.files[0];

            // Prepare designation lines
            doc.setFont("helvetica", "normal");
            doc.setFontSize(8.5);
            const desigLines = doc.splitTextToSize(desigVal, SIG_W);

            // signature block start
            let blockTop = y;

            // draw signature if present
            if (signFile) {

                const imgData = await fileToBase64(signFile);
                const format = signFile.type === "image/png" ? "PNG" : "JPEG";

                doc.addImage(imgData, format, RIGHT_X - SIG_W, blockTop, SIG_W, SIG_H);

                blockTop += SIG_H + 2;
                // draw signature line
                doc.setDrawColor(180, 180, 180);
                doc.setLineWidth(0.3);
                doc.line(RIGHT_X - SIG_W, blockTop, RIGHT_X, blockTop);

            }



            // NAME
            doc.setFont("helvetica", "bold");
            doc.setFontSize(9);

            const nameWidth = doc.getTextWidth(nameVal);
            const nameX = RIGHT_X - (SIG_W / 2) - (nameWidth / 2);

            doc.text(nameVal, nameX, blockTop + 5);

            // DESIGNATION
            doc.setFont("helvetica", "normal");
            doc.setFontSize(8.5);

            let roleY = blockTop + 10;

            desigLines.forEach(line => {

                const lineWidth = doc.getTextWidth(line);
                const lineX = RIGHT_X - (SIG_W / 2) - (lineWidth / 2);

                doc.text(line, lineX, roleY);

                roleY += 4;

            });
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
        const sortedReceipts = [...state.receipts].sort((a, b) => a.order - b.order);


        for (const r of sortedReceipts) {
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

    document.addEventListener("input", e => {
        if (e.target.matches(".table-input")) {
            e.target.value = e.target.value.replace(/[<>]/g, "");
        }
    });

});