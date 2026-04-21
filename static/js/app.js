const state = {
  sessionId: "",
  sessions: [],
  slots: {},
  messages: [],
  generationLocked: false,
  isGeneratingPpt: false,
  isGeneratingDoc: false,
  isRevising: false,
  progressPercent: 0,
  progressLabel: "",
  progressTimer: null,
  templateSuggestions: [],
  templateCatalog: [],
  recommendedTemplate: "",
  selectedTemplate: "",
  templatePickerVisible: false,
  activePptSlideIndex: 0,
  currentPptSlides: [],
  isVoiceListening: false,
  speechRecognition: null,
};

const $ = (selector) => document.querySelector(selector);

const REQUIRED_SLOT_COUNT = 5;

function getFilledSlotCount(slots) {
  return Object.values(slots || {}).filter((value) => String(value || "").trim()).length;
}

function isSlotReady() {
  return getFilledSlotCount(state.slots) >= REQUIRED_SLOT_COUNT;
}

function isGeneratingBusy() {
  return Boolean(state.isGeneratingPpt || state.isGeneratingDoc);
}

async function requestJSON(url, options = {}) {
  const response = await fetch(url, options);
  const raw = await response.text();
  let data = null;
  try {
    data = raw ? JSON.parse(raw) : {};
  } catch {
    if (!response.ok) {
      const brief = (raw || "").replace(/\s+/g, " ").slice(0, 160);
      throw new Error(`请求失败（${response.status}）${brief ? `：${brief}` : ""}`);
    }
    throw new Error("服务返回了非 JSON 响应，请刷新页面后重试。");
  }
  if (!response.ok) throw new Error(data?.error || `请求失败（${response.status}）`);
  return data;
}

function switchView(view) {
  $("#chatView")?.classList.toggle("active", view === "chat");
  $("#resultView")?.classList.toggle("active", view === "result");
}

function updateDemandProgress(slots) {
  const values = Object.values(slots || {});
  const filled = getFilledSlotCount(slots || {});
  $("#sessionIdText").textContent = `${Math.round((filled / Math.max(values.length, 1)) * 100)}%`;
}

function renderSlots(slots) {
  const target = $("#slotBoard");
  if (!target) return;
  const labels = {
    course_theme: "课程主题",
    knowledge_points: "知识点",
    key_difficulties: "重难点",
    lesson_periods: "课时安排",
    style: "课件风格",
  };
  target.innerHTML = Object.entries(labels).map(([key, label]) => `
    <article class="slot-item"><span>${label}</span><strong>${slots[key] || "待补充"}</strong></article>
  `).join("");
  updateDemandProgress(slots);
}

function renderMessages() {
  const chat = $("#chatList");
  if (!chat) return;
  chat.innerHTML = state.messages.map((msg) => `
    <article class="chat-item ${msg.role === "user" ? "user" : "assistant"}">
      <div class="meta">${msg.role === "user" ? "教师" : "智能助理"}</div>
      <div>${msg.content || ""}</div>
    </article>
  `).join("");
  chat.scrollTop = chat.scrollHeight;
}

function appendChat(role, content) {
  state.messages.push({ role, content });
  renderMessages();
}

function resetUploads() {
  const target = $("#uploadResults");
  if (!target) return;
  target.className = "info-list empty";
  target.textContent = "还没有上传任何资料（建议优先上传高数相关内容）。";
  $("#fileInput").value = "";
  $("#chunkCount").textContent = "0";
}

function renderUploads(files, ragStats) {
  const target = $("#uploadResults");
  if (!target) return;
  target.className = "info-list";
  target.innerHTML = files.map((file) => `
    <article class="upload-item">
      <div><strong>${file.original_name}</strong></div>
      <div class="meta">${String(file.file_type || "").toUpperCase()} 解析结果</div>
      ${file.summary ? `<div><strong>摘要：</strong>${file.summary}</div>` : ""}
    </article>
  `).join("");
  $("#chunkCount").textContent = String(ragStats?.indexed_chunks || 0);
}

function resetPreview() {
  $("#retrievalList").className = "info-list empty";
  $("#retrievalList").textContent = "生成后会显示资料融入说明。";
  $("#downloadList").className = "info-list empty";
  $("#downloadList").textContent = "生成后会在这里提供下载入口。";
  $("#pptThumbPreview").className = "ppt-thumb-grid empty";
  $("#pptThumbPreview").textContent = "生成后会在左侧显示 PPT 页面的缩略图（可滚动）。";
  $("#pptFocusPreview").className = "ppt-focus-preview empty";
  $("#pptFocusPreview").textContent = "点击左侧任意缩略图后，这里会显示该页的放大预览。";
  $("#docPreview").className = "info-list empty";
  $("#docPreview").textContent = "生成后会显示教案结构预览。";
  $("#lessonSummary").className = "info-list empty";
  $("#lessonSummary").textContent = "生成后会显示摘要与优化建议。";
  $("#nextStepList").className = "next-step-list empty";
  $("#nextStepList").textContent = "生成后会在这里告诉你下一步该点哪里。";
  state.currentPptSlides = [];
  state.activePptSlideIndex = 0;
}

function renderNextSteps(mode = "initial", result = null) {
  const target = $("#nextStepList");
  if (!target) return;
  if (mode === "initial") {
    target.className = "next-step-list empty";
    target.textContent = "生成后会在这里告诉你下一步该点哪里。";
    return;
  }
  const hasDownloads = Boolean(result?.downloads?.pptx || result?.downloads?.docx);
  const revised = Boolean(result?.revision_applied);
  const cards = [
    {
      title: "先下载并快速检查",
      desc: "建议先打开生成文件看结构是否完整，再决定是否需要优化。",
      actions: hasDownloads ? [{ label: "查看下载区", action: "focus-download" }] : [],
    },
    {
      title: revised ? "继续细化优化" : "提交修改意见",
      desc: revised ? "你可以继续补充更细致的修改意见。" : "若不满意当前版本，直接输入意见并优化。",
      actions: [{ label: "去修改意见", action: "focus-revision" }],
    },
    {
      title: "开始下一份备课",
      desc: "当前版本确认后，可创建新会话开始下一节课。",
      actions: [{ label: "创建新会话", action: "new-session" }],
    },
  ];
  target.className = "next-step-list";
  target.innerHTML = cards.map((card) => `
    <article class="next-step-card">
      <strong>${card.title}</strong>
      <div class="meta">${card.desc}</div>
      <div class="next-step-actions">
        ${card.actions.map((a) => `<button class="next-step-btn" data-next-action="${a.action}">${a.label}</button>`).join("")}
      </div>
    </article>
  `).join("");
}

function changeActivePptSlide(delta) {
  const slides = state.currentPptSlides || [];
  if (!slides.length) return;
  const next = Math.max(0, Math.min((state.activePptSlideIndex || 0) + delta, slides.length - 1));
  if (next === state.activePptSlideIndex) return;
  state.activePptSlideIndex = next;
  renderPptThumbs(slides);
}

function setupPptKeyboardNavigation() {
  document.addEventListener("keydown", (event) => {
    const target = event.target;
    const tagName = (target?.tagName || "").toLowerCase();
    const isEditable = target?.isContentEditable || ["input", "textarea", "select"].includes(tagName);
    if (isEditable) return;
    if (event.key === "ArrowUp") {
      event.preventDefault();
      changeActivePptSlide(-1);
    }
    if (event.key === "ArrowDown") {
      event.preventDefault();
      changeActivePptSlide(1);
    }
  });
}

function renderPptFocus(slides, index = 0) {
  const target = $("#pptFocusPreview");
  if (!target) return;
  if (!slides.length) {
    target.className = "ppt-focus-preview empty";
    target.textContent = "暂无可显示的放大预览。";
    return;
  }
  const safeIndex = Math.max(0, Math.min(index, slides.length - 1));
  const slide = slides[safeIndex] || {};
  const focusVisual = slide.thumbnail_url
    ? `<img class="ppt-focus-image" src="${slide.thumbnail_url}" alt="PPT 第 ${safeIndex + 1} 页缩略图" loading="lazy" />`
    : `<div class="ppt-focus-text-fallback"><h4>${slide.title || `第 ${safeIndex + 1} 页`}</h4><ul>${(slide.bullets || []).slice(0, 5).map((bullet) => `<li>${bullet}</li>`).join("")}</ul></div>`;
  target.className = "ppt-focus-preview";
  target.innerHTML = `
    <div class="ppt-focus-canvas">${focusVisual}</div>
    <div class="ppt-focus-meta">
      <span class="preview-label ppt">PPT 放大预览</span>
      <span class="meta">第 ${safeIndex + 1} / ${slides.length} 页 · 布局：${slide.layout || "content"}</span>
    </div>
  `;
}

function renderPptThumbs(slides) {
  const target = $("#pptThumbPreview");
  if (!target) return;
  if (!slides.length) {
    target.className = "ppt-thumb-grid empty";
    target.textContent = "暂无可显示的 PPT 略缩图。";
    renderPptFocus([], 0);
    return;
  }
  const activeIndex = Math.max(0, Math.min(state.activePptSlideIndex || 0, slides.length - 1));
  state.activePptSlideIndex = activeIndex;
  target.className = "ppt-thumb-grid";
  target.innerHTML = slides.map((slide, index) => `
    <article class="ppt-thumb ${index === activeIndex ? "selected" : ""}" data-slide-index="${index}">
      <div class="ppt-thumb-canvas">
        ${slide.thumbnail_url
          ? `<img class="ppt-thumb-image" src="${slide.thumbnail_url}" alt="PPT 第 ${index + 1} 页缩略图" loading="lazy" />`
          : `<div class="ppt-thumb-text-fallback"><h4>${slide.title || `第 ${index + 1} 页`}</h4><ul>${(slide.bullets || []).slice(0, 3).map((bullet) => `<li>${bullet}</li>`).join("")}</ul></div>`}
      </div>
      <div class="ppt-thumb-meta">
        <span class="preview-label ppt">PPT</span>
        <span class="meta">第 ${index + 1} 页 · ${slide.layout || "content"}</span>
      </div>
    </article>
  `).join("");

  target.querySelectorAll(".ppt-thumb").forEach((item) => {
    item.addEventListener("click", () => {
      const index = Number(item.getAttribute("data-slide-index") || 0);
      state.activePptSlideIndex = Number.isNaN(index) ? 0 : index;
      renderPptThumbs(slides);
    });
  });

  const activeThumb = target.querySelector(`.ppt-thumb[data-slide-index="${activeIndex}"]`);
  activeThumb?.scrollIntoView({ block: "nearest", behavior: "smooth" });
  renderPptFocus(slides, activeIndex);
}

function renderDocPreview(docPreview) {
  const target = $("#docPreview");
  if (!target) return;
  const sections = docPreview?.sections || [];
  if (!sections.length) {
    target.className = "info-list empty";
    target.textContent = "暂无可显示的教案预览内容。";
    return;
  }
  target.className = "info-list";
  target.innerHTML = sections.map((section) => `
    <article class="lesson-card">
      <div><span class="preview-label docx">DOCX</span></div>
      <div><strong>${section.title || "教案章节"}</strong></div>
      <ul>${(section.items || []).slice(0, 4).map((item) => `<li>${item}</li>`).join("")}</ul>
    </article>
  `).join("");
}

function renderPackage(result) {
  $("#apiMode").textContent = result.api_mode === "online" ? "DeepSeek 生成" : "本地兜底生成";
  hydrateTemplateMeta(result);
  const pptSlides = result?.ppt_preview?.slides || result?.package?.slides || [];
  state.currentPptSlides = Array.isArray(pptSlides) ? pptSlides : [];
  state.activePptSlideIndex = 0;

  $("#retrievalList").className = "info-list";
  $("#retrievalList").innerHTML = (result.reference_digest || []).map((item) => `
    <article class="lesson-card">
      <div><strong>${item.title}</strong></div>
      ${item.score ? `<div class="meta">相关度 ${item.score}</div>` : ""}
      <div>${item.summary || "该资料已用于完善课件内容。"}</div>
    </article>
  `).join("") || "<div class='empty'>暂无可展示的资料融入说明。</div>";

  const downloads = [];
  if (result.downloads?.pptx) {
    const pptName = encodeURIComponent(result.downloads.pptx);
    downloads.push(`<article class=\"lesson-card\"><a class=\"file-link\" href=\"/api/download/${pptName}\" target=\"_blank\">下载 PPTX 课件</a></article>`);
  }
  if (result.downloads?.docx) {
    const docName = encodeURIComponent(result.downloads.docx);
    downloads.push(`<article class=\"lesson-card\"><a class=\"file-link\" href=\"/api/download/${docName}\" target=\"_blank\">下载 DOCX 教案</a></article>`);
  }
  $("#downloadList").className = "info-list";
  $("#downloadList").innerHTML = downloads.join("");

  renderPptThumbs(pptSlides);
  renderDocPreview(result.doc_preview);

  $("#lessonSummary").className = "info-list";
  $("#lessonSummary").innerHTML = `
    <article class=\"lesson-card\"><strong>本次产出</strong><div>${result.package?.title || ""}</div></article>
    <article class=\"lesson-card\"><strong>可继续优化方向</strong><div>可补充：课堂互动、学情分层、动画演示、作业分层。</div></article>
    ${result.revision_applied ? `<article class=\"lesson-card\"><strong>已应用修改意见</strong><div>${result.revision_applied}</div></article>` : ""}
  `;
  renderNextSteps("generated", result);
}

function templatePreviewColor(name) {
  const map = {
    "教育蓝": "linear-gradient(135deg, #dbeafe, #bfdbfe)",
    "学术风": "linear-gradient(135deg, #f1f5f9, #e2e8f0)",
    "极简风": "linear-gradient(135deg, #f8fafc, #f1f5f9)",
  };
  return map[name] || "linear-gradient(135deg, #e2e8f0, #cbd5e1)";
}

function hydrateTemplateMeta(data) {
  state.templateSuggestions = data.template_suggestions || state.templateSuggestions || [];
  state.templateCatalog = data.template_catalog || state.templateCatalog || [];
  state.recommendedTemplate = data.recommended_template || state.recommendedTemplate || "";
  state.selectedTemplate = data.selected_template || state.selectedTemplate || state.recommendedTemplate || "";
  renderTemplateSelector();
  renderTemplateCards();
}

function renderTemplateSelector() {
  const toggleBtn = $("#toggleTemplatePickerBtn");
  const selectedText = $("#selectedTemplateText");
  const enabled = isSlotReady();
  if (toggleBtn) {
    toggleBtn.disabled = !enabled;
    toggleBtn.textContent = enabled ? "选择模板" : "补全需求后可选模板";
  }
  if (selectedText) {
    selectedText.textContent = state.selectedTemplate
      ? `已选模板：${state.selectedTemplate}`
      : (state.recommendedTemplate ? `推荐模板：${state.recommendedTemplate}` : "尚未选择模板（将自动使用推荐模板）");
  }
}

function renderTemplateCards() {
  const container = $("#templateCardGrid");
  if (!container) return;
  const cards = (state.templateCatalog || []).slice(0, 12);
  if (!cards.length) {
    container.innerHTML = "<div class='meta'>请先补全教学需求，系统会推荐模板。</div>";
    return;
  }
  container.innerHTML = cards.map((item) => `
    <article class="template-card ${item.name === state.selectedTemplate ? "selected" : ""}" data-template-name="${item.name}">
      <div class="template-mini" style="${item.thumbnail_url ? `background-image:url('${item.thumbnail_url}');background-size:cover;background-position:center;` : `background:${templatePreviewColor(item.name)};`}"></div>
      <div class="template-title">${item.name}${item.recommended ? " · 推荐" : ""}</div>
      <div class="template-scene">${item.scene || "通用课堂"}</div>
    </article>
  `).join("");

  container.querySelectorAll(".template-card").forEach((card) => {
    card.addEventListener("click", () => {
      state.selectedTemplate = card.dataset.templateName || "";
      renderTemplateSelector();
      renderTemplateCards();
      $("#templatePickerHint").textContent = `已选择模板：${state.selectedTemplate}。你也可以继续在对话中补充内容。`;
      closeTemplatePicker();
    });
  });
}

function updateTemplatePickerVisibility(readyToGenerate = false) {
  const modal = $("#templatePickerModal");
  if (!modal) return;
  if (readyToGenerate && !state.templatePickerVisible) {
    modal.classList.remove("hidden");
    modal.setAttribute("aria-hidden", "false");
    state.templatePickerVisible = true;
  }
}

function openTemplatePicker() {
  const modal = $("#templatePickerModal");
  if (!modal) return;
  modal.classList.remove("hidden");
  modal.setAttribute("aria-hidden", "false");
  state.templatePickerVisible = true;
}

function closeTemplatePicker() {
  const modal = $("#templatePickerModal");
  if (!modal) return;
  modal.classList.add("hidden");
  modal.setAttribute("aria-hidden", "true");
}

function updateBackToResultBtn() {
  const wrap = $("#backToResultWrap");
  if (!wrap) return;
  if (state.generationLocked) {
    wrap.classList.remove("hidden");
  } else {
    wrap.classList.add("hidden");
  }
}

function renderProgressBar() {
  const fill = document.querySelector(".ppt-loading-fill");
  const text = document.querySelector(".ppt-loading-text");
  if (!fill || !text) return;
  const pct = Math.max(0, Math.min(100, state.progressPercent || 0));
  fill.style.width = `${pct}%`;
  text.innerHTML = `<span>${state.progressLabel || "准备中..."}</span><span>${pct}%</span>`;
}

function startProgressPolling() {
  stopProgressPolling();
  state.progressPercent = 0;
  state.progressLabel = "准备中...";
  renderProgressBar();
  state.progressTimer = setInterval(async () => {
    if (!state.sessionId) return;
    try {
      const data = await requestJSON(`/api/progress/${encodeURIComponent(state.sessionId)}`);
      state.progressPercent = data.percent || 0;
      state.progressLabel = data.label || "";
      renderProgressBar();
      if (!data.generating && state.progressPercent >= 100) {
        stopProgressPolling();
      }
    } catch {
      stopProgressPolling();
    }
  }, 1200);
}

function stopProgressPolling() {
  if (state.progressTimer) {
    clearInterval(state.progressTimer);
    state.progressTimer = null;
  }
}

function updateButtonStates() {
  const pptBtn = $("#generatePptBtn");
  const docBtn = $("#generateDocBtn");
  const reviseBtn = $("#reviseBtn");
  const sendBtn = $("#sendBtn");
  const attachBtn = $("#attachBtn");
  const voiceBtn = $("#voiceInputBtn");
  const chatInput = $("#chatInput");
  const fileInput = $("#fileInput");
  const guardHint = $("#generationGuardHint");
  const pptLoadingBar = $("#pptLoadingBar");
  const composer = document.querySelector(".composer-shell");
  if (!pptBtn || !docBtn || !reviseBtn) return;
  const slotReady = isSlotReady();
  const generatingBusy = isGeneratingBusy();
  const actionBusy = generatingBusy || state.isRevising;

  pptBtn.disabled = !slotReady || actionBusy;
  docBtn.disabled = !slotReady || actionBusy;
  reviseBtn.disabled = !state.generationLocked || state.isGeneratingPpt || state.isGeneratingDoc || state.isRevising;

  if (sendBtn) sendBtn.disabled = actionBusy;
  if (attachBtn) attachBtn.disabled = actionBusy;
  if (voiceBtn) voiceBtn.disabled = actionBusy;
  if (chatInput) chatInput.disabled = actionBusy;
  if (fileInput) fileInput.disabled = actionBusy;

  if (composer) composer.classList.toggle("locked", actionBusy);

  if (guardHint) {
    guardHint.classList.remove("warn", "locked");
    if (generatingBusy) {
      guardHint.textContent = "正在生成文件：暂时不能继续对话、上传或发起另一个生成。";
      guardHint.classList.add("locked");
    } else if (!slotReady) {
      guardHint.textContent = "请先补全课程主题、知识点、重难点、课时安排、课件风格后再生成。";
      guardHint.classList.add("warn");
    } else {
      guardHint.textContent = "信息已补全，可生成 PPT 或教案。";
    }
  }

  if (pptLoadingBar) {
    const active = Boolean(state.isGeneratingPpt);
    pptLoadingBar.classList.toggle("active", active);
    pptLoadingBar.setAttribute("aria-hidden", active ? "false" : "true");
  }

  pptBtn.textContent = state.isGeneratingPpt ? "生成中" : "生成 PPT";
  docBtn.textContent = state.isGeneratingDoc ? "生成中" : "生成教案";
  reviseBtn.textContent = state.isRevising ? "优化中" : "根据意见重新优化";
  renderTemplateSelector();
  updateBackToResultBtn();
}

function renderSessionHistory() {
  const list = $("#sessionHistoryList");
  if (!list) return;
  if (!state.sessions.length) {
    list.innerHTML = "<div class='meta'>暂无历史会话</div>";
    return;
  }
  list.innerHTML = state.sessions.map((item) => `
    <button class="session-item ${item.session_id === state.sessionId ? "active" : ""}" data-session-id="${item.session_id}">
      <div class="session-title">${item.title || "新对话"}</div>
      <div class="session-meta">${item.has_result ? "已生成结果" : "对话中"} · ${item.message_count || 0} 条消息</div>
    </button>
  `).join("");
}

async function refreshSessionHistory() {
  const data = await requestJSON("/api/sessions");
  state.sessions = data.sessions || [];
  renderSessionHistory();
}

function applySessionSnapshot(snapshot) {
  state.sessionId = snapshot.session_id;
  state.messages = snapshot.messages || [];
  state.slots = snapshot.slots || { course_theme: "", knowledge_points: "", key_difficulties: "", lesson_periods: "", style: "" };
  state.selectedTemplate = snapshot.selected_template || "";
  state.generationLocked = Boolean(snapshot.has_result);
  hydrateTemplateMeta(snapshot);
  renderSlots(state.slots);
  renderMessages();
  renderUploads(snapshot.documents || [], { indexed_chunks: $("#chunkCount")?.textContent || 0 });
  if (snapshot.has_result && snapshot.result) {
    renderPackage(snapshot.result);
    switchView("result");
  } else {
    resetPreview();
    switchView("chat");
  }
  updateButtonStates();
}

async function loadSession(sessionId) {
  const snapshot = await requestJSON(`/api/session/${encodeURIComponent(sessionId)}`);
  applySessionSnapshot(snapshot);
  await refreshSessionHistory();
}

async function createSession() {
  const data = await requestJSON("/api/session/new", { method: "POST" });
  state.sessionId = data.session_id;
  state.messages = [];
  state.slots = { course_theme: "", knowledge_points: "", key_difficulties: "", lesson_periods: "", style: "" };
  state.generationLocked = false;
  state.isGeneratingPpt = false;
  state.isGeneratingDoc = false;
  state.isRevising = false;
  state.progressPercent = 0;
  state.progressLabel = "";
  stopProgressPolling();
  state.templateSuggestions = [];
  state.templateCatalog = [];
  state.recommendedTemplate = "";
  state.selectedTemplate = "";
  state.templatePickerVisible = false;
  $("#sessionState").textContent = "会话已创建";
  $("#apiMode").textContent = "等待 DeepSeek";
  renderSlots(state.slots);
  resetUploads();
  resetPreview();
  appendChat("assistant", "新对话已创建。请先描述课程主题、知识点、重难点与课时安排。");
  closeTemplatePicker();
  updateButtonStates();
  switchView("chat");
  await refreshSessionHistory();
}

async function sendMessage() {
  if (!state.sessionId) await createSession();
  if (isGeneratingBusy()) {
    alert("正在生成文件，请稍候再继续对话。");
    return;
  }
  const message = $("#chatInput").value.trim();
  if (!message) return;
  appendChat("user", message);
  $("#chatInput").value = "";
  const data = await requestJSON("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ session_id: state.sessionId, message }),
  });
  state.slots = data.slots || state.slots;
  state.messages = data.messages || state.messages;
  renderSlots(state.slots);
  renderMessages();
  hydrateTemplateMeta(data);
  updateTemplatePickerVisibility(data.ready_to_generate);
  $("#apiMode").textContent = data.api_mode === "online" ? "DeepSeek 需求分析" : "本地规则分析";
  $("#sessionState").textContent = data.ready_to_generate ? "信息已齐，可开始生成" : "等待补充需求";
  updateButtonStates();
  await refreshSessionHistory();
}

async function uploadFiles() {
  if (!state.sessionId) await createSession();
  if (isGeneratingBusy()) {
    alert("正在生成文件，暂时不能上传资料。");
    return;
  }
  const files = $("#fileInput").files;
  if (!files || !files.length) return;
  const formData = new FormData();
  formData.append("session_id", state.sessionId);
  Array.from(files).forEach((file) => formData.append("files", file));
  const response = await fetch("/api/upload", { method: "POST", body: formData });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "上传失败");
  $("#apiMode").textContent = data.parse_mode === "deepseek" ? "DeepSeek + 本地解析" : "本地解析";
  renderUploads(data.all_files || data.files || [], data.rag_stats || {});
  $("#fileInput").value = "";
  await refreshSessionHistory();
}

function updateVoiceStatus(text) {
  const status = $("#voiceStatusText");
  if (status) status.textContent = `语音输入状态：${text}`;
}

function setupVoiceInput() {
  const btn = $("#voiceInputBtn");
  if (!btn) return;
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    btn.disabled = true;
    btn.textContent = "🎙️ 不支持";
    updateVoiceStatus("当前浏览器不支持（建议 Chrome / Edge）");
    return;
  }
  const recognition = new SpeechRecognition();
  recognition.lang = "zh-CN";
  recognition.interimResults = true;
  recognition.continuous = false;
  let baseText = "";
  let finalTranscript = "";
  recognition.onstart = () => {
    state.isVoiceListening = true;
    btn.textContent = "🛑";
    baseText = $("#chatInput")?.value || "";
    finalTranscript = "";
    updateVoiceStatus("正在识别...");
  };
  recognition.onresult = (event) => {
    let interimTranscript = "";
    for (let i = event.resultIndex; i < event.results.length; i += 1) {
      const text = event.results[i][0]?.transcript || "";
      if (event.results[i].isFinal) finalTranscript += text;
      else interimTranscript += text;
    }
    const input = $("#chatInput");
    if (!input) return;
    input.value = `${baseText}${baseText ? " " : ""}${finalTranscript}${interimTranscript}`.trim();
  };
  recognition.onerror = (event) => {
    state.isVoiceListening = false;
    btn.textContent = "🎙️";
    updateVoiceStatus(`识别失败：${event.error || "未知错误"}`);
  };
  recognition.onend = () => {
    state.isVoiceListening = false;
    btn.textContent = "🎙️";
    updateVoiceStatus("识别结束");
  };
  state.speechRecognition = recognition;
  btn.addEventListener("click", () => {
    if (!state.speechRecognition) return;
    if (state.isVoiceListening) state.speechRecognition.stop();
    else state.speechRecognition.start();
  });
}

async function generatePpt() {
  if (!state.sessionId) await createSession();
  if (!isSlotReady()) {
    alert("请先补全课程主题、知识点、重难点、课时安排、课件风格后再生成。")
    return;
  }
  if (state.isGeneratingPpt || state.isGeneratingDoc || state.isRevising) return;
  state.isGeneratingPpt = true;
  updateButtonStates();
  startProgressPolling();
  $("#sessionState").textContent = "正在生成 PPT...";
  try {
    const result = await requestJSON("/api/generate/ppt", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: state.sessionId, selected_template: state.selectedTemplate }),
    });
    renderPackage(result);
    state.generationLocked = true;
    $("#sessionState").textContent = "PPT 生成完成";
    switchView("result");
    await refreshSessionHistory();
  } finally {
    state.isGeneratingPpt = false;
    stopProgressPolling();
    state.progressPercent = 100;
    state.progressLabel = "生成完成";
    renderProgressBar();
    updateButtonStates();
  }
}

async function generateDoc() {
  if (!state.sessionId) await createSession();
  if (!isSlotReady()) {
    alert("请先补全课程主题、知识点、重难点、课时安排、课件风格后再生成。")
    return;
  }
  if (state.isGeneratingDoc || state.isGeneratingPpt || state.isRevising) return;
  state.isGeneratingDoc = true;
  updateButtonStates();
  startProgressPolling();
  $("#sessionState").textContent = "正在生成教案...";
  try {
    const result = await requestJSON("/api/generate/docx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: state.sessionId, selected_template: state.selectedTemplate }),
    });
    renderPackage(result);
    state.generationLocked = true;
    $("#sessionState").textContent = "教案生成完成";
    switchView("result");
    await refreshSessionHistory();
  } finally {
    state.isGeneratingDoc = false;
    stopProgressPolling();
    state.progressPercent = 100;
    state.progressLabel = "生成完成";
    renderProgressBar();
    updateButtonStates();
  }
}

async function revisePackage() {
  if (!state.sessionId) await createSession();
  const revision = $("#revisionInput").value.trim();
  if (!revision) {
    alert("请先输入修改意见。");
    return;
  }
  if (state.isRevising) return;
  state.isRevising = true;
  updateButtonStates();
  startProgressPolling();
  $("#sessionState").textContent = "正在根据修改意见优化...";
  try {
    const result = await requestJSON("/api/revise", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: state.sessionId, revision, selected_template: state.selectedTemplate }),
    });
    renderPackage(result);
    $("#revisionInput").value = "";
    state.generationLocked = true;
    $("#sessionState").textContent = "已根据意见重新优化";
    switchView("result");
    await refreshSessionHistory();
  } finally {
    state.isRevising = false;
    stopProgressPolling();
    state.progressPercent = 100;
    state.progressLabel = "优化完成";
    renderProgressBar();
    updateButtonStates();
  }
}

function handleError(error) {
  alert(error.message || "发生未知错误");
}

function bootstrapTemplateModal() {
  $("#toggleTemplatePickerBtn")?.addEventListener("click", () => {
    if (!isSlotReady()) return;
    openTemplatePicker();
  });
  $("#closeTemplatePickerBtn")?.addEventListener("click", () => closeTemplatePicker());
  $("#closeTemplatePickerMask")?.addEventListener("click", () => closeTemplatePicker());
  $("#skipTemplateBtn")?.addEventListener("click", () => {
    $("#templatePickerHint").textContent = "已跳过模板选择。你可以继续补充需求，或直接开始生成。";
    closeTemplatePicker();
  });
}

function bindEvents() {
  $("#sidebarNewChatBtn")?.addEventListener("click", () => createSession().catch(handleError));
  $("#newSessionBtn")?.addEventListener("click", () => createSession().catch(handleError));
  $("#backToChatBtn")?.addEventListener("click", () => switchView("chat"));
  $("#backToResultBtn")?.addEventListener("click", () => switchView("result"));
  $("#sendBtn")?.addEventListener("click", () => sendMessage().catch(handleError));
  $("#chatInput")?.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      sendMessage().catch(handleError);
    }
  });
  $("#attachBtn")?.addEventListener("click", () => $("#fileInput")?.click());
  $("#fileInput")?.addEventListener("change", () => uploadFiles().catch(handleError));
  $("#generatePptBtn")?.addEventListener("click", () => generatePpt().catch(handleError));
  $("#generateDocBtn")?.addEventListener("click", () => generateDoc().catch(handleError));
  $("#reviseBtn")?.addEventListener("click", () => revisePackage().catch(handleError));

  $("#sessionHistoryList")?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-session-id]");
    if (!button) return;
    loadSession(button.getAttribute("data-session-id")).catch(handleError);
  });

  $("#nextStepList")?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-next-action]");
    if (!button) return;
    const action = button.getAttribute("data-next-action");
    if (action === "focus-download") {
      $("#downloadList")?.scrollIntoView({ behavior: "smooth", block: "start" });
      return;
    }
    if (action === "focus-revision") {
      switchView("result");
      $("#revisionFloatingPanel")?.scrollIntoView({ behavior: "smooth", block: "end" });
      window.setTimeout(() => $("#revisionInput")?.focus(), 180);
      return;
    }
    if (action === "new-session") {
      createSession().catch(handleError);
    }
  });
}

async function bootstrap() {
  renderSlots({ course_theme: "", knowledge_points: "", key_difficulties: "", lesson_periods: "", style: "" });
  resetUploads();
  resetPreview();
  renderTemplateSelector();
  renderTemplateCards();
  updateButtonStates();
  setupVoiceInput();
  setupPptKeyboardNavigation();
  bootstrapTemplateModal();
  bindEvents();
  await refreshSessionHistory();
  await createSession();
}

bootstrap().catch(handleError);
