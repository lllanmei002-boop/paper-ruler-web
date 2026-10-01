(() => {
  "use strict";

  const PX_PER_MM = 96 / 25.4;
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const PAGE_SIZES = {
    A4: [210, 297],
    "A4-landscape": [297, 210],
    A3: [297, 420],
    Letter: [215.9, 279.4]
  };

  const DEFAULT_SETTINGS = {
    pageSize: "A4",
    width: 210,
    height: 297,
    margin: 12,
    columns: 20,
    rows: 20,
    pairingMode: "mixed",
    rulePunctuation: true,
    rulePairs: true,
    ruleHalfwidth: true,
    ruleSpaces: true,
    ruleStructure: true,
    gridColor: "#c98f85",
    showPageNumber: true,
    gridFont: "KaiTi, STKaiti, Kaiti SC, SimSun, serif",
    mergePunctuation: false
  };

  const tutorialBlocks = [
    block("title", "新手教学示例"),
    block("body", "尊敬的评审老师："),
    block("body", "本人申请参加青年人才交流计划，希望在项目中学习项目设计、资源整合与成果评估方法。"),
    block("body", "本次申请材料包含（“，”）连续标点示例，方便查看同格压缩效果。"),
    block("body", "我会认真完成各项学习任务，并在结束后提交可落地的实践方案。"),
    block("signature", "申请人：李明"),
    block("date", "2026年9月30日")
  ];

  const state = {
    settings: { ...DEFAULT_SETTINGS },
    blocks: [],
    fileName: "未导入文档",
    fileType: "空白",
    layout: null,
    issues: [],
    currentPage: 0,
    zoom: 0.45,
    autoFit: true,
    selectedIssue: "",
    issueFilter: "all",
    renderTimer: 0,
    syncLock: false
  };

  const elements = {};
  let segments = null;
  let tutorialReturnFocus = null;
  let crcTable;

  document.addEventListener("DOMContentLoaded", initialize);

  function initialize() {
    cacheElements();
    bindEvents();
    syncControls();
    renderAll();
    requestAnimationFrame(() => fitToWindow(false));
  }

  function block(type, text, extra = {}) {
    return { type, text: String(text ?? ""), align: "", bold: false, ...extra };
  }

  function cacheElements() {
    [
      "fileInput","dropZone","fileCard","fileIcon","fileName","fileMeta","replaceFileBtn","pasteInput","pasteBtn",
      "clearBtn","sampleBtn","printBtn","exportPngBtn","exportPdfBtn","documentTitle","outlineCount",
      "outlineList","summaryRing","summaryIssueCount","summaryTitle","summaryText","errorCount","warningCount",
      "compareStage","emptyState","emptyImportBtn","singleLayout",
      "previewScroll","previewPages","previewPageCount",
      "zoomOutBtn","zoomInBtn","zoomLabel","fitBtn","prevPageBtn","nextPageBtn","currentPageLabel","pageCountLabel",
      "capacityLabel","issueFilter","issuePanelTitle","issueList","toastRegion","loadingOverlay",
      "loadingTitle","loadingText","printPageStyle","customSizeFields","emptyTutorialBtn","tutorialOverlay","tutorialDialog","tutorialCloseBtn","tutorialUseBtn","tutorialDismissBtn","tutorialMiniPreview"
    ].forEach((id) => { elements[id] = document.getElementById(id); });
  }

  function getSegments() {
    if (segments) return segments;
    const supported = typeof Intl !== "undefined" && Intl.Segmenter;
    segments = supported ? new Intl.Segmenter("zh-CN", { granularity: "grapheme" }) : null;
    return segments;
  }

  function segmentGraphemes(text) {
    const value = String(text ?? "");
    const segmenter = getSegments();
    if (!segmenter) return Array.from(value);
    return Array.from(segmenter.segment(value), (item) => item.segment);
  }

  function prepareBlocks(blocks) {
    let cursor = 0;
    return normalizeBlocks(blocks).map((item, index) => {
      const graphemes = item.type === "pageBreak" ? [] : segmentGraphemes(item.text);
      const prepared = {
        ...item,
        blockIndex: index,
        startSeq: cursor,
        graphemes,
        endSeq: cursor + Math.max(0, graphemes.length - 1)
      };
      cursor += graphemes.length;
      return prepared;
    });
  }

  function bindEvents() {
    elements.fileInput.addEventListener("change", (event) => {
      const file = event.target.files && event.target.files[0];
      if (file) handleFile(file);
      event.target.value = "";
    });
    elements.dropZone.addEventListener("click", () => elements.fileInput.click());
    elements.replaceFileBtn.addEventListener("click", () => elements.fileInput.click());
    elements.emptyImportBtn.addEventListener("click", () => elements.fileInput.click());
    ["dragenter","dragover"].forEach((name) => elements.dropZone.addEventListener(name, (event) => {
      event.preventDefault(); elements.dropZone.classList.add("is-dragging");
    }));
    ["dragleave","drop"].forEach((name) => elements.dropZone.addEventListener(name, (event) => {
      event.preventDefault(); elements.dropZone.classList.remove("is-dragging");
    }));
    elements.dropZone.addEventListener("drop", (event) => {
      const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) handleFile(file);
    });
    document.addEventListener("dragover", (event) => event.preventDefault());
    document.addEventListener("drop", (event) => { if (!elements.dropZone.contains(event.target)) event.preventDefault(); });

    elements.pasteBtn.addEventListener("click", () => {
      const text = elements.pasteInput.value.trim();
      if (!text) { showToast("请先粘贴正文内容。", "error"); return; }
      try {
        acceptContent(blocksFromPlainText(text), "粘贴文本", "文本");
        showToast("稿纸预览与格式检查已更新。", "success");
      } catch (error) { showToast(`无法生成稿纸：${friendlyError(error)}`, "error"); }
    });
    elements.sampleBtn.addEventListener("click", () => openTutorial(elements.sampleBtn));
    elements.emptyTutorialBtn.addEventListener("click", () => openTutorial(elements.emptyTutorialBtn));
    elements.tutorialCloseBtn.addEventListener("click", closeTutorial);
    elements.tutorialDismissBtn.addEventListener("click", closeTutorial);
    elements.tutorialUseBtn.addEventListener("click", loadTutorialSample);
    elements.tutorialOverlay.addEventListener("click", (event) => { if (event.target === elements.tutorialOverlay) closeTutorial(); });
    document.addEventListener("keydown", handleTutorialKeydown);
    elements.clearBtn.addEventListener("click", clearDocument);

    document.querySelectorAll("[data-setting]").forEach((control) => {
      const name = control.dataset.setting;
      const eventName = ["text","number","color"].includes(control.type) || control.type === "radio" ? "input" : "change";
      control.addEventListener(eventName, () => {
        if (name === "pageSize") { applyPageSize(control.value); return; }
        state.settings[name] = readControlValue(control);
        if (name === "width" || name === "height") state.settings.pageSize = "custom";
        syncControls();
        scheduleRender();
      });
    });

    document.querySelectorAll("[data-grid-preset]").forEach((button) => {
      button.addEventListener("click", () => {
        const [columns, rows] = button.dataset.gridPreset.split("x").map(Number);
        state.settings.columns = columns; state.settings.rows = rows;
        syncControls(); scheduleRender();
      });
    });


    elements.zoomOutBtn.addEventListener("click", () => { state.autoFit = false; setZoom(state.zoom - 0.08); });
    elements.zoomInBtn.addEventListener("click", () => { state.autoFit = false; setZoom(state.zoom + 0.08); });
    elements.fitBtn.addEventListener("click", () => { state.autoFit = true; fitToWindow(true); });
    elements.prevPageBtn.addEventListener("click", () => scrollToPage(state.currentPage - 1));
    elements.nextPageBtn.addEventListener("click", () => scrollToPage(state.currentPage + 1));
    elements.issueFilter.addEventListener("change", () => { state.issueFilter = elements.issueFilter.value; renderIssueList(); });

    elements.printBtn.addEventListener("click", printReference);
    elements.exportPngBtn.addEventListener("click", exportImages);
    elements.exportPdfBtn.addEventListener("click", exportPdf);
    window.addEventListener("resize", debounce(() => {
      if (state.autoFit) fitToWindow(false);
    }, 160));
    elements.singleLayout.addEventListener("click", (event) => {
      const cell = event.target.closest(".grid-cell[data-issue-ids]");
      if (cell) selectIssue((cell.dataset.issueIds || "").split(" ")[0], true);
    });
    elements.issueList.addEventListener("click", (event) => {
      const item = event.target.closest("[data-issue-id]");
      if (item) selectIssue(item.dataset.issueId, true);
    });
  }

  function readControlValue(control) {
    if (control.type === "checkbox") return control.checked;
    if (control.type === "radio") return control.value;
    if (control.type === "number") {
      const value = Number(control.value);
      return Number.isFinite(value) ? value : 0;
    }
    return control.value;
  }

  function applyPageSize(value) {
    state.settings.pageSize = value;
    if (PAGE_SIZES[value]) {
      state.settings.width = PAGE_SIZES[value][0];
      state.settings.height = PAGE_SIZES[value][1];
    }
    syncControls();
    scheduleRender();
  }

  function syncControls() {
    document.querySelectorAll("[data-setting]").forEach((control) => {
      const key = control.dataset.setting;
      if (!(key in state.settings)) return;
      if (control.type === "checkbox") control.checked = Boolean(state.settings[key]);
      else if (control.type === "radio") control.checked = control.value === state.settings[key];
      else control.value = state.settings[key];
    });
    elements.capacityLabel.textContent = `${Math.max(0, state.settings.columns * state.settings.rows)} 格`;
    elements.customSizeFields.classList.toggle("is-hidden", state.settings.pageSize !== "custom");
    document.querySelectorAll("[data-grid-preset]").forEach((button) => {
      const [columns, rows] = button.dataset.gridPreset.split("x").map(Number);
      button.classList.toggle("is-active", columns === state.settings.columns && rows === state.settings.rows);
    });
  }

  function scheduleRender() {
    window.clearTimeout(state.renderTimer);
    state.renderTimer = window.setTimeout(renderAll, 60);
  }

  function getGridMetrics() {
    const widthMm = clamp(Number(state.settings.width) || 210, 80, 1000);
    const heightMm = clamp(Number(state.settings.height) || 297, 80, 1000);
    const margin = clamp(Number(state.settings.margin) || 0, 0, Math.min(widthMm, heightMm) * 0.35);
    const columns = clamp(Math.round(Number(state.settings.columns) || 20), 8, 40);
    const rows = clamp(Math.round(Number(state.settings.rows) || 20), 8, 50);
    const availableWidth = Math.max(40, widthMm - margin * 2);
    const availableHeight = Math.max(40, heightMm - margin * 2);
    const cellMm = Math.min(availableWidth / columns, availableHeight / rows);
    const gridWidthMm = cellMm * columns;
    const gridHeightMm = cellMm * rows;
    const gridLeftMm = (widthMm - gridWidthMm) / 2;
    const gridTopMm = (heightMm - gridHeightMm) / 2;
    return {
      widthMm, heightMm, margin, columns, rows, cellMm,
      widthPx: widthMm * PX_PER_MM, heightPx: heightMm * PX_PER_MM,
      cellPx: cellMm * PX_PER_MM, gridWidthPx: gridWidthMm * PX_PER_MM, gridHeightPx: gridHeightMm * PX_PER_MM,
      gridLeftPx: gridLeftMm * PX_PER_MM, gridTopPx: gridTopMm * PX_PER_MM
    };
  }

  function normalizeBlocks(blocks) {
    const result = [];
    let blankPending = false;
    (blocks || []).forEach((item) => {
      if (!item) return;
      if (item.type === "pageBreak") {
        blankPending = false;
        if (result.length && result[result.length - 1].type !== "pageBreak") result.push(block("pageBreak", ""));
        return;
      }
      const raw = String(item.text ?? "").replace(/\u00a0/g, " ");
      const leading = raw.match(/^[ \t]+/)?.[0] || "";
      const text = raw.replace(/\r/g, "").replace(/[ \t]+\n/g, "\n").trim();
      if (!text) {
        if (result.length && !blankPending) { result.push(block("body", "")); blankPending = true; }
        return;
      }
      blankPending = false;
      result.push({
        type: item.type || "body",
        text,
        align: item.align || "",
        bold: Boolean(item.bold),
        leadingSpaces: leading.replace(/\t/g, "    ").length,
        hadTab: leading.includes("\t") || raw.includes("\t"),
        sourceText: raw
      });
    });
    while (result.length && (!result[result.length - 1].text || result[result.length - 1].type === "pageBreak")) result.pop();
    return result;
  }

  function acceptContent(blocks, fileName, fileType) {
    const prepared = prepareBlocks(blocks);
    if (!prepared.some((item) => item.type !== "pageBreak" && item.graphemes.length)) {
      throw new Error("没有读取到可用文字");
    }
    state.blocks = prepared;
    state.fileName = fileName;
    state.fileType = fileType;
    state.currentPage = 0;
    state.selectedIssue = "";
    renderAll();
  }

  async function handleFile(file) {
    const extension = (file.name.split(".").pop() || "").toLowerCase();
    const supported = ["docx","txt","md","markdown","html","htm"];
    if (!supported.includes(extension)) {
      if (extension === "pdf" || extension === "doc") showToast(`${extension.toUpperCase()} 暂不能直接解析。请复制正文后使用左侧“粘贴正文”。`, "error");
      else showToast("暂不支持这种格式，请使用 DOCX、TXT、Markdown 或 HTML。", "error");
      return;
    }
    showLoading("正在解析文档", file.name);
    try {
      let blocks;
      if (extension === "docx") blocks = await parseDocx(file);
      else {
        const text = await file.text();
        if (extension === "html" || extension === "htm") blocks = blocksFromHtml(text);
        else if (extension === "md" || extension === "markdown") blocks = blocksFromMarkdown(text);
        else blocks = blocksFromPlainText(text);
      }
      acceptContent(blocks, file.name, extension.toUpperCase());
      const count = state.layout ? state.layout.pages.length : 0;
      showToast(`已生成 ${count} 页稿纸，并完成 ${state.issues.length} 项检查。`, "success");
    } catch (error) {
      console.error(error);
      showToast(`解析失败：${friendlyError(error)}`, "error");
    } finally { hideLoading(); }
  }

  function openTutorial(trigger) {
    tutorialReturnFocus = trigger || document.activeElement;
    renderTutorialPreview();
    elements.tutorialOverlay.classList.remove("is-hidden");
    elements.tutorialOverlay.setAttribute("aria-hidden", "false");
    document.body.classList.add("tutorial-open");
    window.setTimeout(() => elements.tutorialDialog.focus(), 0);
  }

  function closeTutorial() {
    if (elements.tutorialOverlay.classList.contains("is-hidden")) return;
    elements.tutorialOverlay.classList.add("is-hidden");
    elements.tutorialOverlay.setAttribute("aria-hidden", "true");
    document.body.classList.remove("tutorial-open");
    if (tutorialReturnFocus && typeof tutorialReturnFocus.focus === "function") tutorialReturnFocus.focus();
    tutorialReturnFocus = null;
  }

  function handleTutorialKeydown(event) {
    if (elements.tutorialOverlay.classList.contains("is-hidden")) return;
    if (event.key === "Escape") { event.preventDefault(); closeTutorial(); return; }
    if (event.key !== "Tab") return;
    const focusable = Array.from(elements.tutorialDialog.querySelectorAll('button,[href],input,select,textarea,[tabindex]:not([tabindex="-1"])'))
      .filter((node) => !node.disabled && node.offsetParent !== null);
    if (!focusable.length) return;
    const first = focusable[0], last = focusable[focusable.length - 1];
    if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
    else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
  }

  function renderTutorialPreview() {
    const previewBlocks = prepareBlocks(tutorialBlocks);
    const savedMerge = state.settings.mergePunctuation;
    state.settings.mergePunctuation = true;
    const layout = buildLayout(previewBlocks, "standard");
    state.settings.mergePunctuation = savedMerge;
    const metrics = getGridMetrics();
    const page = layout.pages[0];
    elements.tutorialMiniPreview.textContent = "";
    if (!page) return;
    const pageNode = renderGridPage(page, metrics, "standard", [], { blocks: previewBlocks });
    const availableWidth = Math.max(180, elements.tutorialMiniPreview.clientWidth - 18);
    const availableHeight = Math.max(220, elements.tutorialMiniPreview.clientHeight - 14);
    const zoom = Math.min(0.31, availableWidth / metrics.widthPx, availableHeight / metrics.heightPx);
    pageNode.style.zoom = String(zoom);
    elements.tutorialMiniPreview.appendChild(pageNode);
  }

  function loadTutorialSample() {
    state.blocks = prepareBlocks(tutorialBlocks);
    state.fileName = "新手教学示例.docx";
    state.fileType = "示例文档";
    state.currentPage = 0;
    state.selectedIssue = "";
    renderAll();
    closeTutorial();
    showToast("已载入新手教学示例。");
  }

  function clearDocument() {
    state.blocks = [];
    state.layout = null;
    state.issues = [];
    state.fileName = "未导入文档";
    state.fileType = "空白";
    state.currentPage = 0;
    state.selectedIssue = "";
    elements.fileCard.classList.add("is-hidden");
    elements.pasteInput.value = "";
    renderAll();
    showToast("已清空当前文档。");
  }

  const OPENING_MARKS = new Set(["（","〔","［","｛","【","《","〈","“","‘","「","『"]);
  const CLOSING_START_MARKS = new Set(["）","〕","］","｝","】","》","〉","”","’","」","』","，","。","！","？","；","：","、",",",".","!","?",";",":"]);
  const HALFWIDTH_PUNCTUATION = new Set([",",".","!","?",";",":","(",")","[","]"]); 
  const PUNCTUATION_GLYPHS = new Set(["，","。","！","？","；","：","、","（","）","【","】","《","》","〈","〉","“","”","‘","’","「","」","『","』","〔","〕","［","］","｛","｝","…","—",",",".","!","?",";",":","(",")","[","]","{","}"]);

  function tokenizeBlock(blockData, mode) {
    const source = blockData.graphemes || [];
    const suggested = mode !== "raw";
    const tokens = [];
    const sourceOffsets = [];
    let sourceCursor = 0;
    source.forEach((glyph) => { sourceOffsets.push(sourceCursor); sourceCursor += glyph.length; });
    let index = 0;
    while (index < source.length) {
      const current = source[index];
      const next = source[index + 1];
      const seqStart = blockData.startSeq + index;
      if (suggested && state.settings.mergePunctuation && isPunctuationGlyph(current) && !protectedSpanAt(blockData.text, sourceOffsets[index])) {
        let end = index;
        while (end < source.length && isPunctuationGlyph(source[end]) && !protectedSpanAt(blockData.text, sourceOffsets[end])) end += 1;
        const run = source.slice(index, end);
        const onlyEllipsis = run.length === 2 && run.every((glyph) => glyph === "…");
        const onlyDash = run.length === 2 && run.every((glyph) => glyph === "—");
        if (run.length >= 2 && !onlyEllipsis && !onlyDash) {
          const cluster = makeToken(run.join(""), 1, "punctuationCluster", seqStart, blockData.startSeq + end - 1, blockData);
          cluster.changed = true;
          tokens.push(cluster);
          index = end;
          continue;
        }
      }
      if (suggested && current === "…" && next === "…") {
        tokens.push(makeToken("……", 2, "ellipsis", seqStart, seqStart + 1, blockData));
        index += 2; continue;
      }
      if (suggested && current === "—" && next === "—") {
        tokens.push(makeToken("——", 2, "dash", seqStart, seqStart + 1, blockData));
        index += 2; continue;
      }
      const category = halfwidthCategory(current);
      if (state.settings.pairingMode !== "single" && category) {
        let text = current;
        let end = index;
        while (end + 1 < source.length && text.length < 2) {
          const nextCategory = halfwidthCategory(source[end + 1]);
          if (!nextCategory) break;
          if (state.settings.pairingMode === "same" && nextCategory !== category) break;
          text += source[end + 1]; end += 1;
        }
        tokens.push(makeToken(text, 1, text.length === 2 ? "asciiPair" : "ascii", seqStart, blockData.startSeq + end, blockData));
        index = end + 1; continue;
      }
      const kind = current === "…" ? "ellipsisHalf" : current === "—" ? "dashHalf" : current === "\t" ? "tab" : category ? "ascii" : "fullwidth";
      tokens.push(makeToken(current, 1, kind, seqStart, seqStart, blockData));
      index += 1;
    }
    return tokens;
  }

  function isPunctuationGlyph(glyph) {
    return PUNCTUATION_GLYPHS.has(glyph);
  }

  function makeToken(text, width, kind, seqStart, seqEnd, blockData) {
    return { text, width, kind, seqStart, seqEnd, blockIndex: blockData.blockIndex, occupants: [], changed: false, moved: false, sourceText: text };
  }

  function halfwidthCategory(glyph) {
    if (!glyph || /\s/.test(glyph)) return "";
    if (/[0-9]/.test(glyph)) return "digit";
    if (/[A-Za-z]/.test(glyph)) return "letter";
    return glyph.codePointAt(0) <= 0xff ? "symbol" : "";
  }
  function firstVisibleToken(line) { return line && line.tokens && line.tokens.length ? line.tokens[0] : null; }
  function lastVisibleToken(line) { return line && line.tokens && line.tokens.length ? line.tokens[line.tokens.length - 1] : null; }
  function tokenStartsClosing(token) { return token && CLOSING_START_MARKS.has(Array.from(token.text)[0]); }
  function tokenEndsOpening(token) { return token && OPENING_MARKS.has(Array.from(token.text).at(-1)); }

  function createLine(blockData, leading = 0, blank = false) {
    return { tokens: [], leading, used: leading, blockIndex: blockData?.blockIndex ?? -1, kind: blockData?.type || "body", blank, changed: false };
  }

  function isSalutation(blockData) {
    return blockData.type === "body" && /[：:]$/.test(blockData.text) && blockData.graphemes.length <= Math.max(8, Math.floor(state.settings.columns * 0.75));
  }

  function linesForBlock(blockData, mode) {
    if (!blockData || blockData.type === "pageBreak") return [{ forcedBreak: true }];
    if (!blockData.graphemes.length) return [createLine(blockData, 0, true)];
    const tokens = tokenizeBlock(blockData, mode);
    if (blockData.type === "title") return alignedLines(tokens, "center", blockData);
    if (blockData.type === "signature" || blockData.type === "date") return alignedLines(tokens, "right", blockData);
    return packBodyTokens(tokens, blockData, mode, isSalutation(blockData) ? 0 : 2);
  }

  function alignedLines(tokens, alignment, blockData) {
    const lines = [];
    let current = createLine(blockData);
    tokens.forEach((token) => {
      if (current.tokens.length && current.used + token.width > state.settings.columns) {
        lines.push(current); current = createLine(blockData);
      }
      current.tokens.push(token); current.used += token.width;
    });
    if ((current.tokens.length && !lines.includes(current)) || !lines.length) lines.push(current);
    lines.forEach((line) => {
      if (alignment === "center") line.leading = Math.max(0, Math.floor((state.settings.columns - line.used) / 2));
      else if (alignment === "right") line.leading = Math.max(0, state.settings.columns - line.used);
      line.used += line.leading;
    });
    return lines;
  }

  function packBodyTokens(tokens, blockData, mode, firstIndent) {
    const lines = [];
    let current = createLine(blockData, firstIndent);
    tokens.forEach((token) => {
      const columns = state.settings.columns;
      if (current.used + token.width > columns && current.tokens.length) {
        if (mode === "standard" && token.width === 2) {
          current = splitTwoCellTokenAcrossLines(token, blockData, current, lines);
          return;
        }
        if (token.width === 2 || tokenEndsOpening(token)) token.changed = true;
        lines.push(current); current = createLine(blockData, 0);
      }
      if (mode !== "raw" && state.settings.rulePunctuation && !current.tokens.length && tokenStartsClosing(token)) {
        if (handleClosingAtStart(current, token, lines, mode)) return;
      }
      const remaining = columns - current.used;
      if (mode !== "raw" && mode !== "standard" && state.settings.rulePunctuation && tokenEndsOpening(token) && current.tokens.length && remaining <= token.width) {
        token.changed = true;
        lines.push(current); current = createLine(blockData, 0);
      }
      if (mode !== "raw" && mode !== "standard" && state.settings.rulePunctuation && token.width === 2 && columns - current.used < 2) {
        if (current.tokens.length) { lines.push(current); current = createLine(blockData, 0); }
      }
      if (mode === "standard" && token.width === 2 && columns - current.used < 2 && current.tokens.length) {
        current = splitTwoCellTokenAcrossLines(token, blockData, current, lines);
        return;
      }
      current.tokens.push(token); current.used += token.width;
    });
    if ((current.tokens.length && !lines.includes(current)) || !lines.length) lines.push(current);
    return lines;
  }

  function splitTwoCellTokenAcrossLines(token, blockData, current, lines) {
    const glyphs = Array.from(token.text);
    if (glyphs.length !== 2) { lines.push(current); return; }
    const kind = token.kind === "ellipsis" || token.kind === "dash" ? `${token.kind}Half` : "fullwidth";
    const first = makeToken(glyphs[0], 1, kind, token.seqStart, token.seqStart, blockData);
    const second = makeToken(glyphs[1], 1, kind, token.seqEnd, token.seqEnd, blockData);
    first.changed = false; second.changed = false;
    current.tokens.push(first); current.used += 1; lines.push(current);
    const continuation = createLine(blockData, 0); continuation.tokens.push(second); continuation.used = 1; lines.push(continuation);
    return continuation;
  }

  function handleClosingAtStart(current, token, lines, mode) {
    if (!lines.length) return false;
    const previous = lines[lines.length - 1];
    const previousToken = lastVisibleToken(previous);
    if (!previousToken) return false;
    const effectiveMode = mode === "standard" ? "squeeze" : mode;
    if (effectiveMode === "squeeze") {
      previousToken.occupants.push({ text: token.text, seqStart: token.seqStart, seqEnd: token.seqEnd, kind: token.kind, changed: true });
      previous.changed = true; previousToken.changed = true; token.resolvedBy = "squeeze";
      return true;
    }
    if (effectiveMode === "shift") {
      if (previousToken.width + token.width > state.settings.columns) return false;
      previous.tokens.pop(); previous.used -= previousToken.width;
      previousToken.moved = true; previousToken.changed = true;
      current.tokens.push(previousToken, token); current.used += previousToken.width + token.width;
      current.changed = true; token.changed = true; token.resolvedBy = "shift";
      return true;
    }
    return false;
  }

  function buildLayout(blocks, mode) {
    const logicalLines = [];
    blocks.forEach((blockData) => {
      if (blockData.type === "pageBreak") { logicalLines.push({ forcedBreak: true }); return; }
      linesForBlock(blockData, mode).forEach((line) => logicalLines.push(line));
    });
    return paginateLines(logicalLines, state.settings.rows, state.settings.columns, mode);
  }

  function paginateLines(lines, rows, columns, mode) {
    const pages = [];
    let current = [];
    const flush = () => {
      if (!current.length) return;
      pages.push({
        index: pages.length,
        lines: current.map((line) => ({ ...line, tokens: line.tokens.map((token) => ({ ...token, occupants: [...(token.occupants || [])] })) })),
        rows, columns
      });
      current = [];
    };
    lines.forEach((line) => {
      if (line.forcedBreak) { flush(); return; }
      if (current.length >= rows) flush();
      current.push(line);
    });
    flush();
    if (!pages.length) pages.push({ index: 0, lines: [], rows, columns });
    return { pages, mode };
  }

  function getCellPlacements(page) {
    const placements = [];
    page.lines.forEach((line, row) => {
      let col = line.leading;
      line.tokens.forEach((token) => {
        placements.push({ row, col, token, width: token.width });
        col += token.width;
      });
    });
    return placements;
  }

  function findSeqPosition(layout, seqStart) {
    if (!layout || !Number.isFinite(seqStart)) return null;
    for (const page of layout.pages) {
      const placements = getCellPlacements(page);
      for (const placement of placements) {
        const token = placement.token;
        const inToken = seqStart >= token.seqStart && seqStart <= token.seqEnd;
        const inOccupant = (token.occupants || []).some((item) => seqStart >= item.seqStart && seqStart <= item.seqEnd);
        if (inToken || inOccupant) {
          return { page: page.index, row: placement.row, col: placement.col, width: placement.width, changed: Boolean(token.changed || (token.occupants || []).some((item) => item.changed)) };
        }
      }
    }
    return null;
  }

  function scanIssues(rawLayout) {
    const issues = [];
    const push = (rule, severity, message, suggested, blockIndex, seqStart, seqEnd = seqStart, category = "punct") => {
      issues.push({ rule, severity, message, suggested, blockIndex, seqStart, seqEnd, category, solutions: {} });
    };
    const rules = state.settings;
    state.blocks.forEach((blockData) => {
      if (blockData.type === "pageBreak" || !blockData.graphemes.length) return;
      const text = blockData.text;
      const graphemes = blockData.graphemes;
      const offsets = [];
      let cursor = 0;
      graphemes.forEach((glyph) => { offsets.push(cursor); cursor += glyph.length; });

      if (state.settings.ruleSpaces) {
        if (blockData.hadTab) push("space-tab","warning","正文中存在制表符，稿纸落格时会产生难以控制的大空格。","删除制表符，正文首行缩进由两格空白格表示。",blockData.blockIndex,blockData.startSeq,blockData.startSeq,"spaces");
        const repeated = text.match(/ {2,}/);
        if (repeated) {
          const byteIndex = text.indexOf(repeated);
          const seq = offsets.findIndex((offset) => offset >= byteIndex);
          push("space-repeat","warning","存在连续半角空格，可能出现占格不一致。","保留必要的字间空格；段落缩进请让工具自动留两格。",blockData.blockIndex,blockData.startSeq + Math.max(0,seq),blockData.startSeq + Math.max(0,seq),"spaces");
        }
        if (blockData.leadingSpaces > 0 && blockData.type === "body") {
          push("space-leading","warning","正文开头存在手工空格，可能与自动两格缩进重复。","删除手工空格，改由稿纸缩进规则控制。",blockData.blockIndex,blockData.startSeq,blockData.startSeq,"structure");
        }
      }

      if (state.settings.ruleHalfwidth && /[\u3400-\u9fff]/.test(text)) {
        const fullwidthMap = { ",":"，", ".":"。", "!":"！", "?":"？", ";":"；", ":":"：", "(":"（", ")":"）" };
        graphemes.forEach((glyph, index) => {
          if (!HALFWIDTH_PUNCTUATION.has(glyph)) return;
          if (protectedSpanAt(text, offsets[index])) return;
          if (glyph === "." && graphemes[index - 1] === ".") return;
          push("halfwidth-punctuation","warning",`中文内容中使用了半角“${glyph}”，照抄时容易写成不符合规范的标点。`,`建议改为全角“${fullwidthMap[glyph]}”。`,blockData.blockIndex,blockData.startSeq + index,blockData.startSeq + index,"halfwidth");
        });
        graphemes.forEach((glyph, index) => {
          if ((glyph === "\"" || glyph === "'") && !protectedSpanAt(text, offsets[index])) {
            push("straight-quote","warning","中文内容中使用了直引号，稿纸规范通常要求使用弯引号。","将直引号替换为“ ”或‘ ’。",blockData.blockIndex,blockData.startSeq + index,blockData.startSeq + index,"halfwidth");
          }
        });
      }
      if (state.settings.rulePairs) scanPairIssues(blockData, graphemes, offsets, issues);
      if (state.settings.ruleStructure && blockData.type === "title" && blockData.graphemes.length > state.settings.columns * 2) {
        push("title-long","warning","标题较长，会分成多行居中排列。","确认标题断行位置自然，且每行居中。",blockData.blockIndex,blockData.startSeq,blockData.startSeq + blockData.graphemes.length - 1,"structure");
      }
    });
    if (state.settings.ruleSpaces) {
      for (let index = 1; index < state.blocks.length; index += 1) {
        const current = state.blocks[index], previous = state.blocks[index - 1];
        if (!current.graphemes.length && !previous.graphemes.length && current.type !== "pageBreak" && previous.type !== "pageBreak") {
          push("blank-multiple","warning","连续多个空行会过多占用稿纸格。","最多保留一个空行。",current.blockIndex,current.startSeq,current.startSeq,"spaces");
        }
      }
    }
    if (rules.rulePunctuation) scanLineBreakIssues(rawLayout, issues);
    const unique = new Map();
    issues.forEach((issue) => {
      const key = `${issue.rule}|${issue.blockIndex}|${issue.seqStart}|${issue.seqEnd}`;
      if (!unique.has(key)) unique.set(key, issue);
    });
    return Array.from(unique.values());
  }

  function protectedSpanAt(text, index) {
    const patterns = [
      /https?:\/\/\S+/gi,
      /[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}/g,
      /\b\d{1,2}:\d{2}(?::\d{2})?\b/g,
      /\b\d+[.,:]\d+\b/g,
      /\b\d{1,3}(?:,\d{3})+\b/g
    ];
    return patterns.some((pattern) => {
      pattern.lastIndex = 0;
      let match;
      while ((match = pattern.exec(text))) {
        if (index >= match.index && index < match.index + match[0].length) return true;
      }
      return false;
    });
  }

  function scanPairIssues(blockData, graphemes, offsets, issues) {
    const pairs = [["（","）"],["【","】"],["《","》"],["“","”"],["‘","’"],["「","」"],["『","』"]];
    pairs.forEach(([open, close]) => {
      const stack = [];
      graphemes.forEach((glyph, index) => {
        if (glyph === open) stack.push(index);
        else if (glyph === close) {
          if (stack.length) stack.pop();
          else issues.push({ rule:"pair-close", severity:"error", message:`后符号“${close}”没有对应的前符号。`, suggested:`补写“${open}”，或删除多余的“${close}”。`, blockIndex:blockData.blockIndex, seqStart:blockData.startSeq + index, seqEnd:blockData.startSeq + index, category:"pairs", solutions:{} });
        }
      });
      stack.forEach((index) => issues.push({ rule:"pair-open", severity:"error", message:`前符号“${open}”没有对应的后符号。`, suggested:`补写“${close}”，或删除多余的“${open}”。`, blockIndex:blockData.blockIndex, seqStart:blockData.startSeq + index, seqEnd:blockData.startSeq + index, category:"pairs", solutions:{} }));
    });
  }

  function scanLineBreakIssues(rawLayout, issues) {
    const entries = [];
    rawLayout.pages.forEach((page) => page.lines.forEach((line, row) => entries.push({ page, row, line })));
    entries.forEach((entry, position) => {
      const { page, row, line } = entry;
      const first = firstVisibleToken(line);
      const last = lastVisibleToken(line);
      if (first && tokenStartsClosing(first)) {
        issues.push({ rule:"line-start", severity:"error", message:`标点“${lineStartGlyph(first)}”出现在行首。`, suggested:"当前行首没有可压缩的前格，请调整前文；稿纸保持连续并已标记。", blockIndex:first.blockIndex, seqStart:first.seqStart, seqEnd:first.seqStart, category:"punct", solutions:{} });
      }
      const next = entries[position + 1];
      const nextLine = next && next.line.blockIndex === line.blockIndex ? next.line : null;
      if (last && tokenEndsOpening(last) && nextLine) {
        issues.push({ rule:"line-end", severity:"error", message:`前引号或前括号“${Array.from(last.text).at(-1)}”出现在行末。`, suggested:"前引号或前括号不宜置于行末；当前稿纸不留空，请手写调整并已标记。", blockIndex:last.blockIndex, seqStart:last.seqEnd, seqEnd:last.seqEnd, category:"punct", solutions:{} });
      }
      if (last && (last.kind === "ellipsisHalf" || last.kind === "dashHalf") && nextLine) {
        const nextFirst = firstVisibleToken(nextLine);
        if (nextFirst && nextFirst.text === last.text) {
          issues.push({ rule:"two-cell-split", severity:"error", message:`“${last.text}${nextFirst.text}”被拆到两行。`, suggested:"省略号或破折号跨越两行；当前保持连续并已标记，建议整组移到下一行。", blockIndex:last.blockIndex, seqStart:last.seqStart, seqEnd:nextFirst.seqEnd, category:"punct", solutions:{} });
        }
      }
    });
  }

  function lineStartGlyph(token) { return Array.from(token.text)[0] || token.text; }

  function attachIssues(layout) {
    const issues = scanIssues(layout);
    issues.forEach((issue, index) => {
      issue.id = `issue-${index + 1}`;
      const position = findSeqPosition(layout, issue.seqStart) || { page:0,row:0,col:0,width:1,changed:false };
      issue.rawPage = position.page;
      issue.rawRow = position.row;
      issue.rawCol = position.col;
      issue.rawWidth = position.width;
      if (issue.rule === "two-cell-split" && issue.seqEnd > issue.seqStart) {
        const second = findSeqPosition(layout, issue.seqEnd);
        issue.positions = second ? [position, second] : [position];
      } else {
        issue.positions = [position];
      }
    });
    return issues;
  }

  function renderAll() {
    window.clearTimeout(state.renderTimer);
    const hasContent = state.blocks.length > 0;
    if (!hasContent) {
      state.layout = null; state.issues = [];
      elements.emptyState.classList.remove("is-hidden"); elements.singleLayout.classList.add("is-hidden");
      elements.documentTitle.textContent = "未导入文档";
      elements.previewPages.textContent = ""; elements.previewPageCount.textContent = "0 页";
      elements.issueList.innerHTML = '<div class="issue-empty">导入文档后会显示逐格问题。</div>';
      updateSummary(); updateNavigation(); return;
    }
    state.layout = buildLayout(state.blocks, "standard");
    state.issues = attachIssues(state.layout);
    elements.emptyState.classList.add("is-hidden"); elements.singleLayout.classList.remove("is-hidden");
    renderStandardLayout(); renderOutline(); renderIssueList(); updateDocumentMeta(); updateSummary(); updateNavigation();
    if (state.autoFit) requestAnimationFrame(() => fitToWindow(false));
  }

  function renderStandardLayout() {
    renderPane(elements.previewPages, state.layout.pages, state.issues, "standard");
    elements.previewPageCount.textContent = `${state.layout.pages.length} 页`;
  }

  function renderPane(target, pages, issues, mode = "standard") {
    target.textContent = "";
    const metrics = getGridMetrics();
    pages.forEach((page) => target.appendChild(renderGridPage(page, metrics, mode, issues)));
    target.style.zoom = String(state.zoom);
  }

  function renderGridPage(page, metrics, mode, issues = [], options = {}) {
    const renderBlocks = options.blocks || state.blocks;
    const selectedIssue = options.selectedIssue ?? state.selectedIssue;
    const pageNode = document.createElement("div");
    pageNode.className = "grid-page";
    pageNode.dataset.pageIndex = String(page.index); pageNode.dataset.mode = mode;
    pageNode.style.width = `${metrics.widthPx}px`; pageNode.style.height = `${metrics.heightPx}px`;
    pageNode.style.setProperty("--cell-line", state.settings.gridColor);
    pageNode.style.setProperty("--grid-font", state.settings.gridFont);

    const metaLeft = document.createElement("span");
    metaLeft.className = "grid-paper-meta grid-meta-left";
    metaLeft.textContent = `${metrics.columns} × ${metrics.rows} 稿纸`;
    metaLeft.style.top = `${Math.max(8, metrics.gridTopPx - 18)}px`;
    pageNode.appendChild(metaLeft);
    const metaRight = document.createElement("span");
    metaRight.className = "grid-paper-meta grid-meta-right";
    metaRight.textContent = `每页 ${metrics.columns * metrics.rows} 格`;
    metaRight.style.top = `${Math.max(8, metrics.gridTopPx - 18)}px`;
    pageNode.appendChild(metaRight);

    const surface = document.createElement("div");
    surface.className = "grid-surface";
    surface.style.left = `${metrics.gridLeftPx}px`; surface.style.top = `${metrics.gridTopPx}px`;
    surface.style.width = `${metrics.gridWidthPx}px`; surface.style.height = `${metrics.gridHeightPx}px`;
    pageNode.appendChild(surface);

    const issueMap = buildCellMap(issues, page.index, mode);
    for (let row = 0; row < metrics.rows; row += 1) {
      for (let col = 0; col < metrics.columns; col += 1) {
        const cell = document.createElement("div");
        cell.className = "grid-cell";
        if (row === 0) cell.classList.add("top-edge");
        if (col === 0) cell.classList.add("left-edge");
        const issueIds = issueMap.get(`${row}:${col}`) || [];
        if (issueIds.length) {
          const firstIssue = issues.find((item) => item.id === issueIds[0]);
          cell.classList.add(firstIssue?.severity === "warning" ? "issue-warning" : "issue-error");
          if (issueIds.includes(selectedIssue)) cell.classList.add("selected");
          cell.dataset.issueIds = issueIds.join(" ");
        }
        cell.style.left = `${col * metrics.cellPx}px`; cell.style.top = `${row * metrics.cellPx}px`;
        cell.style.width = `${metrics.cellPx}px`; cell.style.height = `${metrics.cellPx}px`;
        cell.dataset.row = String(row); cell.dataset.col = String(col); cell.dataset.page = String(page.index); cell.dataset.mode = mode;
        surface.appendChild(cell);
      }
    }

    getCellPlacements(page).forEach((placement) => {
      const token = placement.token;
      if (!token.text && !token.occupants.length) return;
      const tokenNode = document.createElement("div");
      tokenNode.className = "token";
      if (token.kind === "asciiPair") tokenNode.classList.add("ascii-pair");
      if (token.kind === "punctuationCluster") tokenNode.classList.add("punctuation-cluster");
      if (token.width === 2) tokenNode.classList.add("two-cell");
      if (token.moved) tokenNode.classList.add("changed-token");
      tokenNode.style.left = `${placement.col * metrics.cellPx + 1}px`;
      tokenNode.style.top = `${placement.row * metrics.cellPx + 1}px`;
      tokenNode.style.width = `${token.width * metrics.cellPx - 2}px`;
      tokenNode.style.height = `${metrics.cellPx - 2}px`;
      const clusterSize = Math.min(metrics.cellPx * 0.3, metrics.cellPx / Math.max(2, Math.ceil(Array.from(token.text).length / 2) * 1.18));
      const size = token.kind === "punctuationCluster" ? clusterSize : token.kind === "asciiPair" ? metrics.cellPx * 0.47 : token.width === 2 ? metrics.cellPx * 0.67 : metrics.cellPx * 0.68;
      tokenNode.style.fontSize = `${size}px`; tokenNode.style.fontFamily = state.settings.gridFont;
      tokenNode.style.fontWeight = token.blockIndex >= 0 && ["title","h1","h2"].includes(renderBlocks[token.blockIndex]?.type) ? "700" : "400";
      if (token.kind === "punctuationCluster") {
        const glyphs = Array.from(token.text);
        tokenNode.style.gridTemplateColumns = "repeat(2, 1fr)";
        tokenNode.style.gridTemplateRows = `repeat(${Math.ceil(glyphs.length / 2)}, 1fr)`;
        glyphs.forEach((glyph) => { const item = document.createElement("span"); item.className = "cluster-glyph"; item.textContent = glyph; tokenNode.appendChild(item); });
      } else {
        const main = document.createElement("span"); main.className = "glyph-main"; main.textContent = token.text; tokenNode.appendChild(main);
        (token.occupants || []).forEach((occupant) => {
          const sub = document.createElement("span"); sub.className = "glyph-sub"; sub.textContent = occupant.text; tokenNode.appendChild(sub);
        });
      }
      if (token.moved) { const arrow = document.createElement("span"); arrow.className = "correction-arrow"; arrow.textContent = "↳"; tokenNode.appendChild(arrow); }
      surface.appendChild(tokenNode);
    });

    if (state.settings.showPageNumber) {
      const number = document.createElement("span");
      number.className = "grid-paper-meta grid-meta-right";
      number.textContent = `第 ${page.index + 1} 页`; number.style.right = "18px"; number.style.bottom = "18px";
      pageNode.appendChild(number);
    }
    return pageNode;
  }

  function buildCellMap(issues, pageIndex, mode) {
    const map = new Map();
    issues.forEach((issue) => {
      if (issue.rawPage !== pageIndex && !(issue.positions || []).some((position) => position.page === pageIndex)) return;
      const positions = (issue.positions || [{ page:issue.rawPage, row:issue.rawRow, col:issue.rawCol, width:issue.rawWidth || 1 }]).filter((position) => position.page === pageIndex);
      positions.forEach((position) => {
        const width = position.width || 1;
        for (let offset = 0; offset < width; offset += 1) {
          const key = `${position.row}:${position.col + offset}`;
          const list = map.get(key) || []; list.push(issue.id); map.set(key, list);
        }
      });
    });
    return map;
  }

  function cellIsChanged(page, row, col) {
    const line = page.lines[row];
    if (!line) return false;
    let cursor = line.leading;
    return line.tokens.some((token) => {
      const start = cursor; cursor += token.width;
      return token.changed && col >= start && col < start + token.width;
    });
  }

  function renderOutline() {
    const items = state.blocks.filter((item) => ["title","h1","h2","signature","date"].includes(item.type) && item.text);
    elements.outlineCount.textContent = String(items.length);
    elements.outlineList.textContent = "";
    if (!items.length) {
      const empty = document.createElement("div"); empty.className = "outline-empty"; empty.textContent = "当前文档没有识别到标题或落款"; elements.outlineList.appendChild(empty); return;
    }
    items.forEach((item) => {
      const button = document.createElement("button"); button.type = "button"; button.className = "outline-item";
      const label = item.type === "title" ? "题" : item.type === "signature" ? "署" : item.type === "date" ? "日" : item.type === "h1" ? "1" : "2";
      const level = document.createElement("span"); level.textContent = label;
      const copy = document.createElement("span"); copy.textContent = item.text;
      button.append(level, copy);
      button.addEventListener("click", () => scrollToBlock(item.blockIndex));
      elements.outlineList.appendChild(button);
    });
  }

  function renderIssueList() {
    const filtered = state.issues.filter((issue) => state.issueFilter === "all" || issue.severity === state.issueFilter);
    elements.issueList.textContent = "";
    elements.issuePanelTitle.textContent = filtered.length ? `${filtered.length} 项需要核对` : "没有发现明显问题";
    if (!filtered.length) {
      const empty = document.createElement("div"); empty.className = "issue-empty"; empty.textContent = state.issues.length ? "当前筛选条件下没有问题" : "稿纸落格未发现明显错误"; elements.issueList.appendChild(empty); return;
    }
    let currentPage = -1;
    filtered.forEach((issue) => {
      if (issue.rawPage !== currentPage) {
        currentPage = issue.rawPage;
        const title = document.createElement("div"); title.className = "issue-group-title"; title.textContent = `第 ${currentPage + 1} 页`; elements.issueList.appendChild(title);
      }
      const button = document.createElement("button");
      button.type = "button"; button.className = `issue-item${issue.severity === "warning" ? " is-warning" : ""}${issue.id === state.selectedIssue ? " is-active" : ""}`;
      button.dataset.issueId = issue.id;
      const title = document.createElement("strong"); title.textContent = `${issue.rawRow + 1} 行 ${issue.rawCol + 1} 格 · ${issue.message}`;
      const location = document.createElement("span"); location.textContent = issue.suggested;
      button.append(title, location);
      elements.issueList.appendChild(button);
    });
  }

  function updateSummary() {
    const errors = state.issues.filter((issue) => issue.severity === "error").length;
    const warnings = state.issues.filter((issue) => issue.severity === "warning").length;
    elements.errorCount.textContent = String(errors); elements.warningCount.textContent = String(warnings);
    elements.summaryIssueCount.textContent = String(state.issues.length);
    if (!state.blocks.length) {
      elements.summaryTitle.textContent = "等待文档";
      elements.summaryText.textContent = "导入后会生成规范稿纸并标记问题。";
    } else if (!state.issues.length) {
      elements.summaryTitle.textContent = "未发现明显落格问题";
      elements.summaryText.textContent = "仍建议逐页检查标题、缩进和落款位置。";
    } else {
      elements.summaryTitle.textContent = `发现 ${errors} 个错误、${warnings} 个提醒`;
      elements.summaryText.textContent = "红格是确定错误，黄格是建议核对；稿纸不会为修正留出多余空格。";
    }
  }

  function updateDocumentMeta() {
    const title = state.blocks.find((item) => item.type === "title")?.text;
    elements.documentTitle.textContent = title || state.fileName.replace(/\.[^.]+$/, "");
    elements.fileName.textContent = state.fileName;
    const chars = state.blocks.reduce((sum, item) => sum + item.graphemes.length, 0);
    elements.fileMeta.textContent = `${chars.toLocaleString("zh-CN")} 字 · ${state.layout?.pages.length || 0} 页规范稿纸`;
    elements.fileCard.classList.remove("is-hidden");
    const ext = state.fileName.includes(".") ? state.fileName.split(".").pop().slice(0, 4).toUpperCase() : state.fileType.slice(0, 4);
    elements.fileIcon.textContent = ext || "TEXT";
  }

  function selectIssue(id, shouldScroll) {
    state.selectedIssue = id;
    document.querySelectorAll(".grid-cell.selected").forEach((cell) => cell.classList.remove("selected"));
    document.querySelectorAll(".issue-item.is-active").forEach((item) => item.classList.remove("is-active"));
    document.querySelectorAll(".grid-cell[data-issue-ids]").forEach((cell) => {
      if ((cell.dataset.issueIds || "").split(" ").includes(id)) cell.classList.add("selected");
    });
    document.querySelectorAll(`[data-issue-id="${id}"]`).forEach((item) => item.classList.add("is-active"));
    if (shouldScroll) scrollToIssue(id);
  }

  function scrollToIssue(id) {
    const issue = state.issues.find((item) => item.id === id);
    if (!issue) return;
    state.currentPage = issue.rawPage;
    updateNavigation();
    const page = elements.previewPages.querySelector(`[data-page-index="${issue.rawPage}"]`);
    if (page) page.scrollIntoView({ behavior:"smooth", block:"start" });
    window.setTimeout(() => {
      const cell = elements.previewPages.querySelector(`.grid-cell.selected[data-page="${issue.rawPage}"]`);
      if (cell) cell.scrollIntoView({ behavior:"smooth", block:"center", inline:"center" });
    }, 180);
  }

  function scrollToBlock(blockIndex) {
    for (let pageIndex = 0; pageIndex < state.layout.pages.length; pageIndex += 1) {
      const page = state.layout.pages[pageIndex];
      if (page.lines.some((line) => line.blockIndex === blockIndex)) { scrollToPage(pageIndex); return; }
    }
  }

  function scrollToPage(index) {
    if (!state.layout) return;
    const target = clamp(index, 0, Math.max(0, state.layout.pages.length - 1));
    state.currentPage = target;
    scrollPaneToPage(elements.previewScroll, elements.previewPages, target);
    updateNavigation();
  }

  function scrollPaneToPage(scroll, pagesRoot, index) {
    const page = pagesRoot.querySelector(`[data-page-index="${index}"]`);
    if (page) scroll.scrollTo({ top:page.offsetTop - 18, behavior:"smooth" });
  }

  function setZoom(value) {
    state.zoom = clamp(value, 0.16, 1.2);
    applyZoom();
  }

  function applyZoom() {
    elements.previewPages.style.zoom = String(state.zoom);
    elements.zoomLabel.textContent = `${Math.round(state.zoom * 100)}%`;
  }

  function fitToWindow(showMessage) {
    const metrics = getGridMetrics();
    const paneWidth = Math.max(220, elements.previewScroll.clientWidth - 56);
    const next = clamp(paneWidth / metrics.widthPx, 0.16, 0.9);
    setZoom(next);
    if (showMessage) showToast(`已适合窗口：${Math.round(next * 100)}%。`);
  }

  function updateNavigation() {
    if (!state.layout) {
      elements.currentPageLabel.textContent = "1"; elements.pageCountLabel.textContent = "1";
      elements.prevPageBtn.disabled = true; elements.nextPageBtn.disabled = true; return;
    }
    const count = state.layout.pages.length;
    state.currentPage = clamp(state.currentPage, 0, Math.max(0, count - 1));
    elements.currentPageLabel.textContent = String(state.currentPage + 1);
    elements.pageCountLabel.textContent = String(count);
    elements.prevPageBtn.disabled = state.currentPage <= 0;
    elements.nextPageBtn.disabled = state.currentPage >= count - 1;
  }

  async function printReference() {
    if (!state.layout) { showToast("请先导入文档。", "error"); return; }
    updatePrintStyle();
    await nextFrame();
    window.print();
  }

  function updatePrintStyle() {
    const metrics = getGridMetrics();
    elements.printPageStyle.textContent = `@page { size: ${metrics.widthMm}mm ${metrics.heightMm}mm; margin: 0; }`;
  }

  async function exportImages() {
    if (!state.layout) { showToast("请先导入文档。", "error"); return; }
    showLoading("正在生成图片", "正在绘制规范稿纸");
    try {
      await nextFrame();
      const metrics = getGridMetrics();
      const base = safeBaseName(state.fileName);
      const files = [];
      for (const page of state.layout.pages) {
        const canvas = renderGridPageToCanvas(page, metrics, 2);
        const blob = await canvasToBlob(canvas, "image/png");
        files.push({ name:`${base}-规范稿纸-第${String(page.index + 1).padStart(2, "0")}页.png`, data:new Uint8Array(await blob.arrayBuffer()) });
      }
      if (files.length === 1) downloadBlob(new Blob([files[0].data], { type:"image/png" }), files[0].name);
      else downloadBlob(createZip(files), `${base}-规范稿纸.zip`);
      showToast(files.length === 1 ? "规范稿纸 PNG 已导出。" : `已导出 ${files.length} 页规范稿纸。`, "success");
    } catch (error) {
      console.error(error); showToast(`导出失败：${friendlyError(error)}`, "error");
    } finally { hideLoading(); }
  }

  async function exportPdf() {
    if (!state.layout) { showToast("请先导入文档。", "error"); return; }
    showLoading("正在生成 PDF", "正在绘制规范稿纸");
    try {
      await nextFrame();
      const metrics = getGridMetrics();
      const images = [];
      for (const page of state.layout.pages) {
        const canvas = renderGridPageToCanvas(page, metrics, 2);
        const blob = await canvasToBlob(canvas, "image/jpeg", 0.93);
        images.push({ data:new Uint8Array(await blob.arrayBuffer()), width:canvas.width, height:canvas.height });
      }
      const pdf = buildImagePdf(images, metrics.widthMm, metrics.heightMm);
      downloadBlob(pdf, `${safeBaseName(state.fileName)}-规范稿纸.pdf`);
      showToast(`已导出 ${state.layout.pages.length} 页规范稿纸 PDF。`, "success");
    } catch (error) {
      console.error(error); showToast(`PDF 导出失败：${friendlyError(error)}`, "error");
    } finally { hideLoading(); }
  }

  function renderGridPageToCanvas(page, metrics, scale) {
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(metrics.widthPx * scale); canvas.height = Math.round(metrics.heightPx * scale);
    const context = canvas.getContext("2d");
    context.save(); context.scale(scale, scale);
    context.fillStyle = "#fffdf7"; context.fillRect(0, 0, metrics.widthPx, metrics.heightPx);
    context.strokeStyle = state.settings.gridColor; context.lineWidth = 0.8;
    const gx = metrics.gridLeftPx, gy = metrics.gridTopPx, cw = metrics.cellPx, ch = metrics.cellPx;
    const issueMap = buildCellMap(state.issues, page.index, "standard");
    for (let row = 0; row < metrics.rows; row += 1) {
      for (let col = 0; col < metrics.columns; col += 1) {
        const x = gx + col * cw, y = gy + row * ch;
        const ids = issueMap.get(`${row}:${col}`);
        if (ids?.length) {
          const issue = state.issues.find((item) => item.id === ids[0]);
          context.fillStyle = issue?.severity === "warning" ? "rgba(218,153,49,.16)" : "rgba(203,73,55,.17)";
          context.fillRect(x, y, cw, ch);
        }
        context.strokeRect(x, y, cw, ch);
      }
    }
    getCellPlacements(page).forEach((placement) => drawGridToken(context, placement, metrics, state.blocks[placement.token.blockIndex]));
    context.fillStyle = "#676a63"; context.font = '10px "Microsoft YaHei",sans-serif'; context.textAlign = "left";
    context.fillText(`${metrics.columns} × ${metrics.rows} 稿纸`, 18, Math.max(17, gy - 8));
    context.textAlign = "right"; context.fillText(`每页 ${metrics.columns * metrics.rows} 格`, metrics.widthPx - 18, Math.max(17, gy - 8));
    if (state.settings.showPageNumber) {
      context.textAlign = "right"; context.fillText(`第 ${page.index + 1} 页`, metrics.widthPx - 18, metrics.heightPx - 18);
    }
    context.restore();
    return canvas;
  }

  function drawGridToken(context, placement, metrics, blockData) {
    const token = placement.token;
    const x = metrics.gridLeftPx + placement.col * metrics.cellPx;
    const y = metrics.gridTopPx + placement.row * metrics.cellPx;
    const width = token.width * metrics.cellPx;
    const type = blockData?.type || "body";
    context.save();
    context.fillStyle = "#151714";
    if (token.kind === "punctuationCluster") {
      const glyphs = Array.from(token.text);
      const columns = 2, rows = Math.ceil(glyphs.length / columns);
      const fontSize = Math.min(metrics.cellPx * 0.3, metrics.cellPx / (rows * 1.18));
      context.font = `400 ${fontSize}px ${state.settings.gridFont}`;
      context.textAlign = "center"; context.textBaseline = "middle";
      glyphs.forEach((glyph, index) => {
        const col = index % columns, row = Math.floor(index / columns);
        context.fillText(glyph, x + (col + 0.5) * metrics.cellPx / columns, y + (row + 0.5) * metrics.cellPx / rows);
      });
    } else {
      const fontSize = token.kind === "asciiPair" ? metrics.cellPx * 0.47 : token.width === 2 ? metrics.cellPx * 0.67 : metrics.cellPx * 0.68;
      context.font = `${["title","h1","h2"].includes(type) ? "700" : "400"} ${fontSize}px ${state.settings.gridFont}`;
      context.textAlign = "center"; context.textBaseline = "middle";
      context.fillText(token.text, x + width / 2, y + metrics.cellPx * 0.52);
      (token.occupants || []).forEach((occupant) => {
        context.font = `400 ${metrics.cellPx * 0.34}px ${state.settings.gridFont}`;
        context.fillStyle = "#8f2e24";
        context.textAlign = "right"; context.textBaseline = "bottom";
        context.fillText(occupant.text, x + width - 2, y + metrics.cellPx - 2);
      });
    }
    context.restore();
  }

  function blocksFromPlainText(text) {
    const normalized = String(text || "").replace(/\r\n?/g, "\n").replace(/\u00a0/g, " ").replace(/^\uFEFF/, "");
    const groups = normalized.split(/\n\s*\n+/).map((group) => group.trim()).filter(Boolean);
    const sourceLines = [];
    groups.forEach((group) => {
      const lines = group.split("\n").map((line) => line.trim()).filter(Boolean);
      let buffer = [];
      const flush = () => { if (buffer.length) sourceLines.push(buffer.join(" ")); buffer = []; };
      lines.forEach((line) => { if (isStandaloneLine(line)) { flush(); sourceLines.push(line); } else buffer.push(line); });
      flush();
    });
    return sourceLines.map((line, index) => {
      let type = "body";
      if (index === 0 && line.length <= 60) type = "title";
      else if (looksLikeDate(line)) type = "date";
      else if (/^(申请人|签名|姓名|单位|联系电话)[:：]/.test(line)) type = "signature";
      else if (/^[一二三四五六七八九十]+[、.]/.test(line) && line.length < 40) type = "h1";
      else if (/^（[一二三四五六七八九十]+）/.test(line) && line.length < 40) type = "h2";
      return block(type, line);
    });
  }

  function isStandaloneLine(line) {
    return line.length <= 26 || looksLikeDate(line) || /^(申请人|签名|姓名|单位|联系电话)[:：]/.test(line) || /^([一二三四五六七八九十]+[、.]|（[一二三四五六七八九十]+）)/.test(line);
  }

  function blocksFromMarkdown(text) {
    const lines = String(text || "").replace(/\r\n?/g, "\n").split("\n");
    const blocks = [];
    let paragraph = [];
    const flush = () => { if (paragraph.length) blocks.push(block("body", paragraph.join(" "))); paragraph = []; };
    lines.forEach((raw) => {
      const line = raw.trim();
      if (!line) { flush(); return; }
      const heading = line.match(/^(#{1,3})\s+(.+)$/);
      if (heading) { flush(); const level = heading[1].length; blocks.push(block(level === 1 ? "title" : level === 2 ? "h1" : "h2", stripInlineMarkdown(heading[2]))); return; }
      const list = line.match(/^[-*+]\s+(.+)$/) || line.match(/^\d+[.)]\s+(.+)$/);
      if (list) { flush(); blocks.push(block("list", `• ${stripInlineMarkdown(list[1])}`)); return; }
      const quote = line.match(/^>\s?(.+)$/);
      if (quote) { flush(); blocks.push(block("quote", stripInlineMarkdown(quote[1]))); return; }
      paragraph.push(stripInlineMarkdown(line));
    });
    flush();
    if (blocks.length && blocks[0].type === "body" && blocks[0].text.length <= 60) blocks[0].type = "title";
    return blocks;
  }

  function blocksFromHtml(text) {
    const doc = new DOMParser().parseFromString(String(text || ""), "text/html");
    doc.querySelectorAll("script,style,noscript,template").forEach((node) => node.remove());
    const nodes = Array.from(doc.body.querySelectorAll("h1,h2,h3,h4,p,li,blockquote"));
    const blocks = [];
    nodes.forEach((node) => {
      const content = cleanWhitespace(node.textContent || "");
      if (!content) return;
      const tag = node.tagName.toLowerCase();
      if (tag === "h1") blocks.push(block("title", content));
      else if (tag === "h2") blocks.push(block("h1", content));
      else if (tag === "h3" || tag === "h4") blocks.push(block("h2", content));
      else if (tag === "li") blocks.push(block("list", `• ${content}`));
      else if (tag === "blockquote") blocks.push(block("quote", content));
      else blocks.push(block(looksLikeDate(content) ? "date" : "body", content));
    });
    if (!blocks.length && doc.body.textContent.trim()) return blocksFromPlainText(doc.body.textContent);
    if (blocks.length && blocks[0].type === "body" && blocks[0].text.length <= 60) blocks[0].type = "title";
    return blocks;
  }

  function stripInlineMarkdown(text) {
    return String(text || "").replace(/!\[([^\]]*)\]\([^)]+\)/g, "$1").replace(/\[([^\]]+)\]\([^)]+\)/g, "$1").replace(/[*_~`]+/g, "").replace(/\\([\\`*_[\]{}()#+\-.!>])/g, "$1").trim();
  }

  function cleanWhitespace(text) { return String(text || "").replace(/\s+/g, " ").trim(); }
  function looksLikeDate(text) { return /^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日$/.test(String(text || "").trim()); }

  async function parseDocx(file) {
    const buffer = await file.arrayBuffer();
    const xmlBytes = await readZipEntry(new Uint8Array(buffer), "word/document.xml");
    const xml = new TextDecoder("utf-8").decode(xmlBytes);
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.querySelector("parsererror")) throw new Error("DOCX 内部 XML 无法解析");
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    const blocks = [];
    let seenNonEmpty = false;
    paragraphs.forEach((paragraph) => {
      const pPr = directChild(paragraph, "pPr");
      const styleId = pPr ? getChildAttribute(pPr, "pStyle", "val") : "";
      const justification = pPr ? getChildAttribute(pPr, "jc", "val") : "";
      const runProps = findFirstRunProperties(paragraph);
      const boldValue = runProps ? getChildAttribute(runProps, "b", "val") : "";
      const text = extractWordParagraphText(paragraph);
      if (text.includes("\f")) {
        text.split("\f").forEach((piece, index, pieces) => {
          if (piece.trim()) { blocks.push(buildWordBlock(piece, styleId, justification, boldValue, !seenNonEmpty)); seenNonEmpty = true; }
          if (index < pieces.length - 1) blocks.push(block("pageBreak", ""));
        });
        return;
      }
      if (!text.trim()) { if (seenNonEmpty) blocks.push(block("body", "")); return; }
      blocks.push(buildWordBlock(text, styleId, justification, boldValue, !seenNonEmpty));
      seenNonEmpty = true;
    });
    return blocks;
  }

  function buildWordBlock(text, styleId, justification, boldValue, isFirst) {
    const value = String(styleId || "").toLowerCase();
    const type = /title|标题$/.test(value) ? "title" : /heading\s*1|heading1|^1$|一级/.test(value) ? "h1" : /heading\s*2|heading2|^2$|二级/.test(value) ? "h2" : /heading\s*3|heading3|^3$|三级/.test(value) ? "h3" : isFirst && String(text).trim().length <= 60 ? "title" : looksLikeDate(text) ? "date" : /^(申请人|签名|姓名|单位|联系电话)[:：]/.test(String(text).trim()) ? "signature" : "body";
    let align = "";
    if (["center","centre"].includes(justification)) align = "center";
    else if (["right","end"].includes(justification)) align = "right";
    else if (["both","distribute","justify"].includes(justification)) align = "justify";
    const bold = !["0","false","off"].includes(String(boldValue).toLowerCase()) && boldValue !== "";
    return block(type, cleanWhitespacePreserveBreaks(text), { align, bold });
  }

  function cleanWhitespacePreserveBreaks(text) { return String(text || "").replace(/\r/g, "").replace(/[ \t]+\n/g, "\n").replace(/^\s+|\s+$/g, ""); }
  function directChild(node, localName) { return Array.from(node.childNodes || []).find((child) => child.nodeType === 1 && child.localName === localName) || null; }
  function getChildAttribute(node, childName, attributeName) { const child = directChild(node, childName); if (!child) return ""; return child.getAttributeNS(W_NS, attributeName) || child.getAttribute(`w:${attributeName}`) || ""; }

  function findFirstRunProperties(paragraph) {
    const walker = paragraph.ownerDocument.createTreeWalker(paragraph, NodeFilter.SHOW_ELEMENT);
    let node = walker.nextNode();
    while (node) { if (node.localName === "rPr") return node; node = walker.nextNode(); }
    return null;
  }

  function extractWordParagraphText(paragraph) {
    const parts = [];
    const walk = (node) => {
      Array.from(node.childNodes || []).forEach((child) => {
        if (child.nodeType !== 1) return;
        if (child.localName === "t") { parts.push(child.textContent || ""); return; }
        if (child.localName === "tab") { parts.push("\t"); return; }
        if (child.localName === "br" || child.localName === "cr") { const type = child.getAttributeNS(W_NS, "type") || child.getAttribute("w:type") || ""; parts.push(type === "page" ? "\f" : "\n"); return; }
        if (!["instrText","delText","rPr","pPr"].includes(child.localName)) walk(child);
      });
    };
    walk(paragraph); return parts.join("");
  }

  async function readZipEntry(bytes, targetName) {
    const eocd = findEndOfCentralDirectory(bytes);
    if (eocd < 0) throw new Error("文件不是有效的 DOCX 压缩包");
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const count = view.getUint16(eocd + 10, true);
    let offset = view.getUint32(eocd + 16, true);
    const decoder = new TextDecoder("utf-8");
    for (let index = 0; index < count; index += 1) {
      if (view.getUint32(offset, true) !== 0x02014b50) break;
      const compression = view.getUint16(offset + 10, true);
      const compressedSize = view.getUint32(offset + 20, true);
      const nameLength = view.getUint16(offset + 28, true);
      const extraLength = view.getUint16(offset + 30, true);
      const commentLength = view.getUint16(offset + 32, true);
      const localOffset = view.getUint32(offset + 42, true);
      const name = decoder.decode(bytes.slice(offset + 46, offset + 46 + nameLength));
      if (name === targetName) {
        const localNameLength = view.getUint16(localOffset + 26, true);
        const localExtraLength = view.getUint16(localOffset + 28, true);
        const start = localOffset + 30 + localNameLength + localExtraLength;
        const compressed = bytes.slice(start, start + compressedSize);
        if (compression === 0) return compressed;
        if (compression !== 8) throw new Error(`不支持 DOCX 压缩方式 ${compression}`);
        if (!("DecompressionStream" in window)) throw new Error("当前浏览器不支持解压 DOCX，请使用新版 Edge、Chrome 或 Firefox");
        const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
        return new Uint8Array(await new Response(stream).arrayBuffer());
      }
      offset += 46 + nameLength + extraLength + commentLength;
    }
    throw new Error("DOCX 中未找到正文内容");
  }

  function findEndOfCentralDirectory(bytes) {
    const min = Math.max(0, bytes.length - 65557);
    for (let index = bytes.length - 22; index >= min; index -= 1) {
      if (bytes[index] === 0x50 && bytes[index + 1] === 0x4b && bytes[index + 2] === 0x05 && bytes[index + 3] === 0x06) return index;
    }
    return -1;
  }

  function createZip(files) {
    const localParts = [], centralParts = [];
    let localOffset = 0;
    const now = new Date();
    const dosTime = ((now.getHours() & 0x1f) << 11) | ((now.getMinutes() & 0x3f) << 5) | (Math.floor(now.getSeconds() / 2) & 0x1f);
    const dosDate = (((now.getFullYear() - 1980) & 0x7f) << 9) | (((now.getMonth() + 1) & 0x0f) << 5) | (now.getDate() & 0x1f);
    files.forEach((file) => {
      const nameBytes = new TextEncoder().encode(file.name);
      const crc = crc32(file.data);
      const local = new Uint8Array(30 + nameBytes.length), lv = new DataView(local.buffer);
      lv.setUint32(0,0x04034b50,true); lv.setUint16(4,20,true); lv.setUint16(6,0x0800,true); lv.setUint16(8,0,true);
      lv.setUint16(10,dosTime,true); lv.setUint16(12,dosDate,true); lv.setUint32(14,crc,true); lv.setUint32(18,file.data.length,true); lv.setUint32(22,file.data.length,true); lv.setUint16(26,nameBytes.length,true); lv.setUint16(28,0,true);
      local.set(nameBytes,30); localParts.push(local,file.data);
      const central = new Uint8Array(46 + nameBytes.length), cv = new DataView(central.buffer);
      cv.setUint32(0,0x02014b50,true); cv.setUint16(4,20,true); cv.setUint16(6,20,true); cv.setUint16(8,0x0800,true); cv.setUint16(10,0,true); cv.setUint16(12,dosTime,true); cv.setUint16(14,dosDate,true); cv.setUint32(16,crc,true); cv.setUint32(20,file.data.length,true); cv.setUint32(24,file.data.length,true); cv.setUint16(28,nameBytes.length,true); cv.setUint16(30,0,true); cv.setUint16(32,0,true); cv.setUint16(34,0,true); cv.setUint16(36,0,true); cv.setUint32(38,0,true); cv.setUint32(42,localOffset,true);
      central.set(nameBytes,46); centralParts.push(central); localOffset += local.length + file.data.length;
    });
    const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
    const eocd = new Uint8Array(22), ev = new DataView(eocd.buffer);
    ev.setUint32(0,0x06054b50,true); ev.setUint16(4,0,true); ev.setUint16(6,0,true); ev.setUint16(8,files.length,true); ev.setUint16(10,files.length,true); ev.setUint32(12,centralSize,true); ev.setUint32(16,localOffset,true); ev.setUint16(20,0,true);
    return new Blob([...localParts,...centralParts,eocd], { type:"application/zip" });
  }

  function crc32(bytes) {
    if (!crcTable) {
      crcTable = new Uint32Array(256);
      for (let index = 0; index < 256; index += 1) {
        let value = index;
        for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
        crcTable[index] = value >>> 0;
      }
    }
    let crc = 0xffffffff;
    for (let index = 0; index < bytes.length; index += 1) crc = crcTable[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }

  function buildImagePdf(images, widthMm, heightMm) {
    if (!images.length) throw new Error("没有可导出的页面");
    const pageWidth = widthMm * 72 / 25.4, pageHeight = heightMm * 72 / 25.4;
    const totalObjects = 2 + images.length * 3;
    const objects = new Array(totalObjects + 1), encoder = new TextEncoder();
    const stringObject = (text) => [encoder.encode(text)];
    const kids = images.map((_, index) => `${3 + index * 3} 0 R`).join(" ");
    objects[1] = stringObject("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
    objects[2] = stringObject(`2 0 obj\n<< /Type /Pages /Count ${images.length} /Kids [${kids}] >>\nendobj\n`);
    images.forEach((image, index) => {
      const pageId = 3 + index * 3, imageId = pageId + 1, contentId = pageId + 2;
      const content = `q\n${formatPdfNumber(pageWidth)} 0 0 ${formatPdfNumber(pageHeight)} 0 0 cm\n/Im0 Do\nQ\n`;
      const contentBytes = encoder.encode(content);
      objects[pageId] = stringObject(`${pageId} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${formatPdfNumber(pageWidth)} ${formatPdfNumber(pageHeight)}] /Resources << /XObject << /Im0 ${imageId} 0 R >> >> /Contents ${contentId} 0 R >>\nendobj\n`);
      objects[imageId] = [encoder.encode(`${imageId} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${image.data.length} >>\nstream\n`), image.data, encoder.encode("\nendstream\nendobj\n")];
      objects[contentId] = [encoder.encode(`${contentId} 0 obj\n<< /Length ${contentBytes.length} >>\nstream\n`), contentBytes, encoder.encode("endstream\nendobj\n")];
    });
    const parts = [], offsets = new Array(totalObjects + 1).fill(0);
    let byteLength = 0;
    const push = (part) => { const bytes = part instanceof Uint8Array ? part : new Uint8Array(part); parts.push(bytes); byteLength += bytes.length; };
    push(encoder.encode("%PDF-1.4\n%1234\n"));
    for (let id = 1; id <= totalObjects; id += 1) { offsets[id] = byteLength; if (!objects[id]) throw new Error("PDF 对象生成失败"); objects[id].forEach(push); }
    const xrefOffset = byteLength;
    let xref = `xref\n0 ${totalObjects + 1}\n0000000000 65535 f \n`;
    for (let id = 1; id <= totalObjects; id += 1) xref += `${String(offsets[id]).padStart(10,"0")} 00000 n \n`;
    xref += `trailer\n<< /Size ${totalObjects + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
    push(encoder.encode(xref));
    return new Blob(parts, { type:"application/pdf" });
  }

  function formatPdfNumber(value) { return Number(value.toFixed(4)).toString(); }
  async function canvasToBlob(canvas, type, quality) { return new Promise((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("浏览器无法创建图片文件")), type, quality)); }
  function downloadBlob(blob, fileName) { const url = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = url; anchor.download = fileName; document.body.appendChild(anchor); anchor.click(); anchor.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 2000); }
  function safeBaseName(fileName) { return String(fileName || "文档").replace(/\.[^.]+$/, "").replace(/[\\/:*?"<>|]+/g, "-").slice(0,80) || "文档"; }
  function showLoading(title, text) { elements.loadingTitle.textContent = title || "正在处理"; elements.loadingText.textContent = text || "请稍候…"; elements.loadingOverlay.classList.remove("is-hidden"); elements.loadingOverlay.setAttribute("aria-hidden","false"); }
  function hideLoading() { elements.loadingOverlay.classList.add("is-hidden"); elements.loadingOverlay.setAttribute("aria-hidden","true"); }
  function showToast(message, type = "") { const toast = document.createElement("div"); toast.className = `toast${type ? ` is-${type}` : ""}`; toast.textContent = message; elements.toastRegion.appendChild(toast); window.setTimeout(() => { toast.style.opacity = "0"; toast.style.transform = "translateY(6px)"; window.setTimeout(() => toast.remove(),180); },3300); }
  function friendlyError(error) { const message = error && error.message ? error.message : String(error); if (/DecompressionStream|deflate-raw/i.test(message)) return "浏览器不支持 DOCX 解压，请改用新版 Edge 或 Chrome"; if (/No such file|not found/i.test(message)) return "文档内部结构不完整"; return message || "未知错误"; }
  function clamp(value, min, max) { return Math.min(max, Math.max(min, value)); }
  function debounce(callback, wait) { let timer; return (...args) => { window.clearTimeout(timer); timer = window.setTimeout(() => callback(...args), wait); }; }
  function nextFrame() { return new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))); }
})();
