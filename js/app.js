/* ============================================
   Genspark free Export – App Logic
   ============================================ */

(function () {
  "use strict";

  // ===== State =====
  let slideCounter = 0;

  // ===== DOM =====
  const slidesContainer = document.getElementById("slides-container");
  const btnAddSlide = document.getElementById("btn-add-slide");
  const btnConvert = document.getElementById("btn-convert");
  const btnConvertText = document.getElementById("btn-convert-text");
  const btnConvertSpinner = document.getElementById("btn-convert-spinner");
  const progressOverlay = document.getElementById("progress-overlay");
  const progressTitle = document.getElementById("progress-title");
  const progressDetail = document.getElementById("progress-detail");
  const progressBar = document.getElementById("progress-bar");
  const renderArea = document.getElementById("render-area");

  // ===== Init =====
  addSlide();

  btnAddSlide.addEventListener("click", () => addSlide());
  btnConvert.addEventListener("click", () => convertToPptx());

  // ===== Add Slide =====
  function addSlide() {
    slideCounter++;
    const id = slideCounter;

    const card = document.createElement("div");
    card.className = "slide-card";
    card.id = `slide-card-${id}`;
    card.dataset.slideId = id;

    // Template without preview pane
    card.innerHTML = `
      <div class="slide-card-header">
        <div class="slide-card-title">
          <div class="slide-number">${id}</div>
          <span class="slide-label">Slide ${id}</span>
        </div>
        <div class="slide-card-actions">
          <button class="btn-icon btn-delete" title="Remove slide" data-id="${id}">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="16" height="16">
              <polyline points="3 6 5 6 21 6"></polyline>
              <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            </svg>
          </button>
        </div>
      </div>
      <div class="slide-card-body">
        <div class="editor-pane">

          <textarea
            class="code-editor"
            id="editor-${id}"
            placeholder="<!-- Paste your HTML slide code here -->\n<div style=&quot;padding: 60px; background: #020230; color: white; width: 1280px; height: 720px;&quot;>\n  <h1>New Slide</h1>\n</div>"
            spellcheck="false"
            data-id="${id}"
          ></textarea>
        </div>
      </div>
    `;

    slidesContainer.appendChild(card);

    const editor = document.getElementById(`editor-${id}`);
    const deleteBtn = card.querySelector(".btn-delete");

    // Tab key support in editor
    editor.addEventListener("keydown", (e) => {
      if (e.key === "Tab") {
        e.preventDefault();
        const start = editor.selectionStart;
        const end = editor.selectionEnd;
        editor.value = editor.value.substring(0, start) + "  " + editor.value.substring(end);
        editor.selectionStart = editor.selectionEnd = start + 2;
      }
    });

    deleteBtn.addEventListener("click", () => {
      if (slidesContainer.children.length <= 1) {
        shakeElement(card);
        return;
      }
      card.style.transition = "all 0.3s ease";
      card.style.opacity = "0";
      card.style.transform = "translateY(-8px) scale(0.98)";
      setTimeout(() => {
        card.remove();
        renumberSlides();
      }, 300);
    });

    // Scroll to new slide
    setTimeout(() => {
      card.scrollIntoView({ behavior: "smooth", block: "center" });
    }, 100);
  }

  // ===== Renumber slides =====
  function renumberSlides() {
    const cards = slidesContainer.querySelectorAll(".slide-card");
    cards.forEach((card, i) => {
      const num = i + 1;
      card.querySelector(".slide-number").textContent = num;
      card.querySelector(".slide-label").textContent = `Slide ${num}`;
    });
  }

  // ===== Shake animation =====
  function shakeElement(el) {
    el.style.animation = "none";
    el.offsetHeight; // force reflow
    el.style.animation = "shake 0.4s ease";
    setTimeout(() => { el.style.animation = ""; }, 400);
    if (!document.getElementById("shake-style")) {
      const style = document.createElement("style");
      style.id = "shake-style";
      style.textContent = `@keyframes shake {
        0%, 100% { transform: translateX(0); }
        20% { transform: translateX(-6px); }
        40% { transform: translateX(6px); }
        60% { transform: translateX(-4px); }
        80% { transform: translateX(4px); }
      }`;
      document.head.appendChild(style);
    }
  }

  // ===== Convert to PPTX =====
  async function convertToPptx() {
    const editors = slidesContainer.querySelectorAll(".code-editor");
    const slides = [];

    editors.forEach((editor) => {
      const html = editor.value.trim();
      if (html) slides.push(html);
    });

    if (slides.length === 0) {
      alert("Please enter some HTML code first.");
      return;
    }

    setConverting(true);
    showProgress(0, slides.length);

    try {
      const pres = new PptxGenJS();
      pres.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
      pres.layout = "WIDE";

      for (let i = 0; i < slides.length; i++) {
        updateProgress(i + 1, slides.length, "Rendering slide…");

        // Render HTML into hidden container
        const imgDataUrl = await renderHtmlToImage(slides[i]);

        // Create slide with full-bleed image
        const slide = pres.addSlide();
        slide.addImage({
          data: imgDataUrl,
          x: 0, y: 0,
          w: 13.333, h: 7.5,
          sizing: { type: "contain", w: 13.333, h: 7.5 },
        });

        updateProgress(i + 1, slides.length, "Slide added ✓");
        await sleep(200);
      }

      updateProgress(slides.length, slides.length, "Generating PPTX…");
      await sleep(400);

      // Download
      await pres.writeFile({ fileName: "genspark-export.pptx" });

      updateProgress(slides.length, slides.length, "Download started ✓");
      await sleep(800);
    } catch (err) {
      console.error("Conversion failed:", err);
      alert("Error during conversion: " + err.message);
    } finally {
      setConverting(false);
    }
  }

  // ===== Render HTML to image =====
  function renderHtmlToImage(htmlContent) {
    return new Promise((resolve, reject) => {
      const container = document.createElement("div");
      container.style.cssText =
        "position:absolute;left:0;top:-9999px;width:1280px;height:720px;overflow:hidden;background:#ffffff;visibility:hidden;";

      const iframe = document.createElement("iframe");
      iframe.style.cssText = "width:1280px;height:720px;border:none;";
      // Need allow-scripts for some CSS/Font features occasionally, but keeping it secure
      iframe.sandbox = "allow-same-origin allow-scripts";
      container.appendChild(iframe);
      document.body.appendChild(container);

      iframe.onload = async () => {
        try {
          const doc = iframe.contentDocument || iframe.contentWindow.document;
          doc.open();

          // Check if user provided a full HTML document or just a snippet
          if (htmlContent.toLowerCase().includes("<html") || htmlContent.toLowerCase().includes("<!doctype")) {
            doc.write(htmlContent);
          } else {
            doc.write(`<!DOCTYPE html>
              <html>
              <head><meta charset="utf-8">
              <style>html,body{margin:0;padding:0;width:1280px;height:720px;overflow:hidden;}</style>
              </head>
              <body>${htmlContent}</body>
              </html>`);
          }
          doc.close();

          // Wait for fonts and all internal resources (like FontAwesome)
          await iframe.contentWindow.document.fonts.ready;

          // Small extra safety delay for rendering engine
          await sleep(1000);

          html2canvas(doc.body, {
            width: 1280,
            height: 720,
            scale: 2, // Higher resolution for better quality in PPTX
            useCORS: true,
            allowTaint: true,
            backgroundColor: null,
            logging: false,
          })
            .then((canvas) => {
              const dataUrl = canvas.toDataURL("image/png");
              container.remove();
              resolve(dataUrl);
            })
            .catch((err) => {
              container.remove();
              reject(err);
            });
        } catch (err) {
          container.remove();
          reject(err);
        }
      };

      iframe.src = "about:blank";
    });
  }

  // ===== UI Helpers =====
  function setConverting(converting) {
    if (converting) {
      btnConvert.disabled = true;
      btnConvertText.classList.add("hidden");
      btnConvertSpinner.classList.remove("hidden");
      progressOverlay.classList.remove("hidden");
    } else {
      btnConvert.disabled = false;
      btnConvertText.classList.remove("hidden");
      btnConvertSpinner.classList.add("hidden");
      progressOverlay.classList.add("hidden");
    }
  }

  function showProgress(current, total) {
    progressBar.style.width = "0%";
    progressDetail.textContent = `Preparing slide ${current} of ${total}`;
  }

  function updateProgress(current, total, detail) {
    const pct = Math.round((current / total) * 100);
    progressBar.style.width = pct + "%";
    progressTitle.textContent = `Converting… (${current}/${total})`;
    progressDetail.textContent = detail || `Slide ${current} of ${total}`;
  }

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }
})();
