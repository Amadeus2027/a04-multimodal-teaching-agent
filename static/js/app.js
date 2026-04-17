const state = {
  sessionId: "",
  slots: {},
  generationLocked: false,
  isGeneratingPpt: false,
  isGeneratingDoc: false,
  isRevising: false,
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

function renderSlots(slots) {
  const labels = {
    course_theme: "课程主题",
    knowledge_points: "知识点",
    key_difficulties: "重难点",
    lesson_periods: "课时安排",
    style: "课件风格",
  };
  $("#slotBoard").innerHTML = Object.entries(labels).map(([key, label]) => `<article class="slot-item"><span>${label}</span><strong>${slots[key] || "待补充"}</strong></article>`).join("");
  updateDemandProgress(slots);
}

function updateDemandProgress(slots) {
  const values = Object.values(slots || {});
  const filled = values.filter((value) => String(value || "").trim()).length;
  $("#sessionIdText").textContent = `${Math.round((filled / Math.max(values.length, 1)) * 100)}%`;
}

function appendChat(role, content) {
  const item = document.createElement("article");
  item.className = `chat-item ${role}`;
  item.innerHTML = `<div class="meta">${role === "user" ? "教师" : "智能助理"}</div><div>${content}</div>`;
  $("#chatList").appendChild(item);
  $("#chatList").scrollTop = $("#chatList").scrollHeight;
}

function resetPreview() {
  $("#retrievalList").className = "info-list empty";
  $("#retrievalList").textContent = "生成后会显示高数相关资料如何被融入课件/教案（Demo）。";
  $("#downloadList").className = "info-list empty";
  $("#downloadList").textContent = "生成后会在这里提供下载入口（Demo）。";
  $("#pptThumbPreview").className = "ppt-thumb-grid empty";
  $("#pptThumbPreview").textContent = "生成后会在左侧显示 PPT 页面的缩略图（可滚动）。";
  $("#pptFocusPreview").className = "ppt-focus-preview empty";
  $("#pptFocusPreview").textContent = "点击左侧任意缩略图后，这里会显示该页的放大预览与详细标注。";
  state.currentPptSlides = [];
  state.activePptSlideIndex = 0;
  $("#docPreview").className = "info-list empty";
  $("#docPreview").textContent = "生成后会显示教案结构预览：教学目标、方法、课堂活动、课后作业。";
  $("#lessonSummary").className = "info-list empty";
  $("#lessonSummary").textContent = "生成后会显示精简摘要与建议（偏高数场景）。";
  $("#nextStepList").className = "next-step-list empty";
  $("#nextStepList").textContent = "生成后会在这里告诉你下一步该点哪里（Demo 引导）。";
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
      return;
    }
    if (event.key === "ArrowDown") {
      event.preventDefault();
      changeActivePptSlide(1);
    }
  });
}

function renderNextSteps(mode = "initial", result = null) {
  const target = $("#nextStepList");
  if (!target) return;

  if (mode === "initial") {
    target.className = "next-step-list empty";
    target.textContent = "生成后会在这里告诉你下一步该点哪里（Demo 引导）。";
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
      desc: revised
        ? "你可以继续补充更细致的修改意见（如课堂活动时长、分层练习难度）。"
        : "若不满意当前版本，直接在下方“修改意见”里输入要求并一键优化。",
      actions: [{ label: "去修改意见", action: "focus-revision" }],
    },
    {
      title: "开始下一份备课",
      desc: "当前版本确认后，可创建新会话开始下一节课，避免覆盖当前成果。",
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

function updateVoiceStatus(text) {
  const status = $("#voiceStatusText");
  if (status) {
    status.textContent = `语音输入状态：${text}`;
  }
}

function resetUploads() {
  $("#uploadResults").className = "info-list empty";
  $("#uploadResults").textContent = "还没有上传任何资料（建议优先上传高数相关内容）。";
  $("#fileInput").value = "";
  $("#fileCount").textContent = "0";
  $("#chunkCount").textContent = "0";
}

function renderUploads(files, ragStats) {
  $("#uploadResults").className = "info-list";
  $("#uploadResults").innerHTML = files.map((file) => `
    <article class="upload-item">
      <div><strong>${file.original_name}</strong></div>
      <div class="meta">${file.file_type.toUpperCase()} 解析结果</div>
      ${file.summary ? `<div><strong>摘要：</strong>${file.summary}</div>` : ""}
      ${Array.isArray(file.key_points) && file.key_points.length ? `<div><strong>要点：</strong>${file.key_points.slice(0, 4).join("；")}</div>` : ""}
      ${Array.isArray(file.knowledge_structure) && file.knowledge_structure.length ? `<div><strong>知识结构：</strong>${file.knowledge_structure.slice(0, 3).join("；")}</div>` : ""}
      ${Array.isArray(file.cases) && file.cases.length ? `<div><strong>案例素材：</strong>${file.cases.slice(0, 3).join("；")}</div>` : ""}
      ${file.content_style ? `<div><strong>版式风格：</strong>${file.content_style}</div>` : ""}
    </article>
  `).join("");
  $("#fileCount").textContent = String(files.length);
  $("#chunkCount").textContent = String(ragStats?.indexed_chunks || 0);
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
    <div class="ppt-focus-canvas">
      ${focusVisual}
    </div>
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
      ${Array.isArray(item.highlights) && item.highlights.length ? `<div class="meta">用于：${item.highlights.join("、")}</div>` : ""}
    </article>
  `).join("") || "<div class='empty'>暂无可展示的资料融入说明。</div>";

  const downloads = [];
  if (result.downloads.pptx) {
    const pptName = encodeURIComponent(result.downloads.pptx);
    downloads.push(`<article class="lesson-card"><a class="file-link" href="/api/download/${pptName}" target="_blank">下载 PPTX 课件</a></article>`);
  }
  if (result.downloads.docx) {
    const docName = encodeURIComponent(result.downloads.docx);
    downloads.push(`<article class="lesson-card"><a class="file-link" href="/api/download/${docName}" target="_blank">下载 DOCX 教案</a></article>`);
  }
  $("#downloadList").className = "info-list";
  $("#downloadList").innerHTML = downloads.join("");

  renderPptThumbs(pptSlides);

  renderDocPreview(result.doc_preview);

  $("#lessonSummary").className = "info-list";
  $("#lessonSummary").innerHTML = `
    <article class="lesson-card"><strong>本次产出</strong><div>${result.package.title}</div></article>
    <article class="lesson-card"><strong>可继续优化方向</strong><div>可补充：课堂互动、学情分层、动画演示、作业分层。</div></article>
    ${result.revision_applied ? `<article class="lesson-card"><strong>已应用修改意见</strong><div>${result.revision_applied}</div></article>` : ""}
  `;

  renderNextSteps("generated", result);
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
  const enabled = Object.values(state.slots || {}).filter((value) => String(value || "").trim()).length >= 5;

  if (toggleBtn) {
    toggleBtn.disabled = !enabled;
    toggleBtn.textContent = enabled ? "选择模板（可收起）" : "补全需求后可选模板";
  }

  if (selectedText) {
    selectedText.textContent = state.selectedTemplate
      ? `已选模板：${state.selectedTemplate}`
      : (state.recommendedTemplate ? `推荐模板：${state.recommendedTemplate}（可手动改选）` : "尚未选择模板（将自动使用推荐模板）");
  }
}

function templatePreviewColor(name) {
  const map = {
    "教育蓝": "linear-gradient(135deg, #dbeafe, #bfdbfe)",
    "学术风": "linear-gradient(135deg, #f1f5f9, #e2e8f0)",
    "极简风": "linear-gradient(135deg, #f8fafc, #f1f5f9)",
    "科技风": "linear-gradient(135deg, #1e3a8a, #0ea5e9)",
    "卡通风": "linear-gradient(135deg, #ffedd5, #fdba74)",
    "可爱风": "linear-gradient(135deg, #fce7f3, #f9a8d4)",
    "插画风": "linear-gradient(135deg, #fff7ed, #fed7aa)",
    "清新风": "linear-gradient(135deg, #dcfce7, #bbf7d0)",
    "中国风": "linear-gradient(135deg, #fef2f2, #fecaca)",
    "新中式": "linear-gradient(135deg, #f5f5f4, #e7e5e4)",
    "商务风": "linear-gradient(135deg, #dbeafe, #93c5fd)",
    "杂志风": "linear-gradient(135deg, #f5f5f4, #e7e5e4)",
    "几何风": "linear-gradient(135deg, #ede9fe, #c4b5fd)",
    "莫兰迪风": "linear-gradient(135deg, #f3f4f6, #d1d5db)",
    "粉彩风": "linear-gradient(135deg, #fdf2f8, #fbcfe8)",
    "暖阳风": "linear-gradient(135deg, #fef3c7, #fde68a)",
    "自然风": "linear-gradient(135deg, #ecfccb, #d9f99d)",
    "森林风": "linear-gradient(135deg, #dcfce7, #86efac)",
    "海洋风": "linear-gradient(135deg, #e0f2fe, #7dd3fc)",
    "复古风": "linear-gradient(135deg, #fef3c7, #fcd34d)",
    "手账风": "linear-gradient(135deg, #ffedd5, #fed7aa)",
    "未来感": "linear-gradient(135deg, #dbeafe, #818cf8)",
    "玻璃拟态": "linear-gradient(135deg, #e0f2fe, #bfdbfe)",
    "扁平风": "linear-gradient(135deg, #f1f5f9, #cbd5e1)",
    "渐变风": "linear-gradient(135deg, #c4b5fd, #f9a8d4)",
  };
  return map[name] || "linear-gradient(135deg, #e2e8f0, #cbd5e1)";
}

function renderTemplateCards() {
  const container = $("#templateCardGrid");
  if (!container) {
    return;
  }
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

function updateButtonStates() {
  $("#generatePptBtn").disabled = state.generationLocked || state.isGeneratingPpt || state.isGeneratingDoc || state.isRevising;
  $("#generateDocBtn").disabled = state.generationLocked || state.isGeneratingPpt || state.isGeneratingDoc || state.isRevising;
  $("#reviseBtn").disabled = !state.generationLocked || state.isGeneratingPpt || state.isGeneratingDoc || state.isRevising;

  if (state.isGeneratingPpt || state.isGeneratingDoc) {
    $("#generatePptBtn").textContent = "生成中";
    $("#generateDocBtn").textContent = "生成中";
  } else if (state.generationLocked) {
    $("#generatePptBtn").textContent = "已生成";
    $("#generateDocBtn").textContent = "已生成";
  } else {
    $("#generatePptBtn").textContent = "生成 PPT";
    $("#generateDocBtn").textContent = "生成教案";
  }

  $("#reviseBtn").textContent = state.isRevising ? "优化中" : "根据意见重新优化";
  renderTemplateSelector();
}

async function createSession() {
  const data = await requestJSON("/api/session/new", { method: "POST" });
  state.sessionId = data.session_id;
  state.slots = { course_theme: "", knowledge_points: "", key_difficulties: "", lesson_periods: "", style: "" };
  state.generationLocked = false;
  state.isGeneratingPpt = false;
  state.isGeneratingDoc = false;
  state.isRevising = false;
  state.templateSuggestions = [];
  state.templateCatalog = [];
  state.recommendedTemplate = "";
  state.selectedTemplate = "";
  state.templatePickerVisible = false;
  $("#sessionState").textContent = "A04 Demo 会话已创建";
  $("#apiMode").textContent = "等待 DeepSeek";
  $("#chatList").innerHTML = "";
  renderSlots(state.slots);
  resetUploads();
  resetPreview();
  appendChat("assistant", "A04 赛题 Demo 会话已创建。当前优先支持高数备课需求；当需求完整后，系统会自动给出推荐模板，你也可以手动选择。");
  closeTemplatePicker();
  updateButtonStates();
}

async function sendMessage() {
  if (!state.sessionId) await createSession();
  const message = $("#chatInput").value.trim();
  if (!message) return;
  appendChat("user", message);
  $("#chatInput").value = "";
  const data = await requestJSON("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ session_id: state.sessionId, message }),
  });
  state.slots = data.slots;
  renderSlots(state.slots);
  hydrateTemplateMeta(data);
  updateTemplatePickerVisibility(data.ready_to_generate);
  $("#apiMode").textContent = data.api_mode === "online" ? "DeepSeek 需求分析" : "本地规则分析";
  $("#sessionState").textContent = data.ready_to_generate ? "信息已齐，可选择模板并开始生成" : "等待补充需求";
  appendChat("assistant", data.assistant_reply);
  updateButtonStates();
}

async function uploadFiles() {
  if (!state.sessionId) await createSession();
  const files = $("#fileInput").files;
  if (!files.length) {
    alert("请先选择 PDF、图片、Word 或视频文件。");
    return;
  }
  const formData = new FormData();
  formData.append("session_id", state.sessionId);
  Array.from(files).forEach((file) => formData.append("files", file));
  const response = await fetch("/api/upload", { method: "POST", body: formData });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "上传失败");
  $("#apiMode").textContent = data.parse_mode === "deepseek" ? "DeepSeek + 本地解析" : "本地解析";
  renderUploads(data.all_files || data.files || [], data.rag_stats);
}

function setupVoiceInput() {
  const btn = $("#voiceInputBtn");
  if (!btn) return;
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    btn.disabled = true;
    btn.textContent = "🎙️ 当前浏览器不支持语音输入";
    updateVoiceStatus("当前浏览器不支持（建议 Chrome / Edge）");
    return;
  }

  if (window.location.protocol !== "https:" && window.location.hostname !== "localhost" && window.location.hostname !== "127.0.0.1") {
    updateVoiceStatus("当前站点非安全上下文，语音识别可能被浏览器拦截");
  } else {
    updateVoiceStatus("可用，点击“语音输入”开始");
  }

  const recognition = new SpeechRecognition();
  recognition.lang = "zh-CN";
  recognition.interimResults = true;
  recognition.continuous = false;
  recognition.maxAlternatives = 1;
  let baseText = "";
  let finalTranscript = "";

  recognition.onstart = () => {
    state.isVoiceListening = true;
    btn.textContent = "🛑 停止语音输入";
    baseText = $("#chatInput")?.value || "";
    finalTranscript = "";
    updateVoiceStatus("正在识别，请开始说话...");
  };

  recognition.onresult = (event) => {
    let interimTranscript = "";
    for (let i = event.resultIndex; i < event.results.length; i += 1) {
      const text = event.results[i][0]?.transcript || "";
      if (event.results[i].isFinal) {
        finalTranscript += text;
      } else {
        interimTranscript += text;
      }
    }
    const input = $("#chatInput");
    if (!input) return;
    const merged = `${baseText}${baseText ? " " : ""}${finalTranscript}${interimTranscript}`.trim();
    input.value = merged;
  };

  recognition.onerror = (event) => {
    state.isVoiceListening = false;
    btn.textContent = "🎙️ 语音输入";
    const errorMap = {
      "not-allowed": "麦克风权限被拒绝，请在浏览器地址栏开启麦克风权限",
      "service-not-allowed": "浏览器策略禁止语音服务",
      "no-speech": "未检测到语音，请重试",
      "audio-capture": "未检测到麦克风设备",
      "network": "语音识别网络异常，请检查网络",
      "aborted": "语音识别已停止",
    };
    updateVoiceStatus(errorMap[event.error] || `语音识别失败：${event.error || "未知错误"}`);
  };

  recognition.onend = () => {
    state.isVoiceListening = false;
    btn.textContent = "🎙️ 语音输入";
    const inputText = ($("#chatInput")?.value || "").trim();
    if (inputText) {
      updateVoiceStatus("识别完成，内容已填入输入框");
    } else {
      updateVoiceStatus("识别结束，但未获得有效文本");
    }
  };

  state.speechRecognition = recognition;

  btn.addEventListener("click", () => {
    if (!state.speechRecognition) return;
    if (state.isVoiceListening) {
      state.speechRecognition.stop();
      return;
    }
    try {
      state.speechRecognition.start();
    } catch (error) {
      updateVoiceStatus(`启动失败：${error?.message || "请稍后重试"}`);
    }
  });
}

async function generatePpt() {
  if (!state.sessionId) await createSession();
  if (state.isGeneratingPpt || state.generationLocked) return;
  state.isGeneratingPpt = true;
  updateButtonStates();
  $("#sessionState").textContent = "正在生成 PPT，请稍候…";
  try {
    const result = await requestJSON("/api/generate/ppt", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: state.sessionId, selected_template: state.selectedTemplate }),
    });
    renderPackage(result);
    state.generationLocked = true;
    $("#sessionState").textContent = "PPT 生成完成，可继续提交修改意见";
    $("#downloadList").scrollIntoView({ behavior: "smooth", block: "start" });
  } finally {
    state.isGeneratingPpt = false;
    updateButtonStates();
  }
}

async function generateDoc() {
  if (!state.sessionId) await createSession();
  if (state.isGeneratingDoc || state.generationLocked) return;
  state.isGeneratingDoc = true;
  updateButtonStates();
  $("#sessionState").textContent = "正在生成教案，请稍候…";
  try {
    const result = await requestJSON("/api/generate/docx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: state.sessionId, selected_template: state.selectedTemplate }),
    });
    renderPackage(result);
    state.generationLocked = true;
    $("#sessionState").textContent = "教案生成完成，可继续提交修改意见";
    $("#downloadList").scrollIntoView({ behavior: "smooth", block: "start" });
  } finally {
    state.isGeneratingDoc = false;
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
  $("#sessionState").textContent = "正在根据修改意见优化，请稍候…";
  try {
    const result = await requestJSON("/api/revise", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: state.sessionId, revision, selected_template: state.selectedTemplate }),
    });
    renderPackage(result);
    $("#revisionInput").value = "";
    $("#sessionState").textContent = "已根据意见重新优化";
  } finally {
    state.isRevising = false;
    updateButtonStates();
  }
}

function handleError(error) {
  alert(error.message || "发生未知错误");
}

function bootstrap() {
  renderSlots({ course_theme: "", knowledge_points: "", key_difficulties: "", lesson_periods: "", style: "" });
  renderTemplateSelector();
  renderTemplateCards();
  updateButtonStates();
  setupVoiceInput();
  setupPptKeyboardNavigation();

  $("#toggleTemplatePickerBtn").addEventListener("click", () => {
    if (Object.values(state.slots || {}).filter((value) => String(value || "").trim()).length < 5) {
      return;
    }
    openTemplatePicker();
  });
  $("#closeTemplatePickerBtn").addEventListener("click", () => closeTemplatePicker());
  $("#closeTemplatePickerMask").addEventListener("click", () => closeTemplatePicker());
  $("#skipTemplateBtn").addEventListener("click", () => {
    $("#templatePickerHint").textContent = "已跳过模板选择。你可以继续补充需求，或直接开始生成。";
    closeTemplatePicker();
  });
  $("#newSessionBtn").addEventListener("click", () => createSession().catch(handleError));
  $("#sendBtn").addEventListener("click", () => sendMessage().catch(handleError));
  $("#chatInput").addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      sendMessage().catch(handleError);
    }
  });
  $("#uploadBtn").addEventListener("click", () => uploadFiles().catch(handleError));
  $("#generatePptBtn").addEventListener("click", () => generatePpt().catch(handleError));
  $("#generateDocBtn").addEventListener("click", () => generateDoc().catch(handleError));
  $("#reviseBtn").addEventListener("click", () => revisePackage().catch(handleError));

  $("#nextStepList").addEventListener("click", (event) => {
    const button = event.target.closest("[data-next-action]");
    if (!button) return;
    const action = button.getAttribute("data-next-action");
    if (action === "focus-download") {
      $("#downloadList")?.scrollIntoView({ behavior: "smooth", block: "start" });
      return;
    }
    if (action === "focus-revision") {
      $("#revisionInput")?.scrollIntoView({ behavior: "smooth", block: "center" });
      $("#revisionInput")?.focus();
      return;
    }
    if (action === "new-session") {
      createSession().catch(handleError);
    }
  });
}

bootstrap();
