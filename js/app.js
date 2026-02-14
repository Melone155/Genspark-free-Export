/* ============================================
   Genspark free Export – App Logic
   ============================================ */

(function () {
  "use strict";

  let slideCounter = 0;

  const slidesContainer = document.getElementById("slides-container");
  const btnAddSlide = document.getElementById("btn-add-slide");
  const btnConvert = document.getElementById("btn-convert");
  const btnConvertText = document.getElementById("btn-convert-text");
  const btnConvertSpinner = document.getElementById("btn-convert-spinner");
  const progressOverlay = document.getElementById("progress-overlay");
  const progressTitle = document.getElementById("progress-title");
  const progressDetail = document.getElementById("progress-detail");
  const progressBar = document.getElementById("progress-bar");

  addSlide();

  btnAddSlide.addEventListener("click", () => addSlide());
  btnConvert.addEventListener("click", () => convertToPptx());

  function addSlide() {
    slideCounter++;
    const id = slideCounter;
    const card = document.createElement("div");
    card.className = "slide-card";
    card.id = `slide-card-${id}`;
    card.innerHTML = `
      <div class="slide-card-header">
        <div class="slide-card-title">
          <div class="slide-number">${id}</div>
          <span class="slide-label">Slide ${id}</span>
        </div>
        <div class="slide-card-actions">
          <button class="btn-icon btn-delete" title="Remove slide">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
          </button>
        </div>
      </div>
      <div class="slide-card-body">
        <textarea class="code-editor" id="editor-${id}" placeholder="Paste HTML code here..."></textarea>
      </div>
    `;
    slidesContainer.appendChild(card);
    const deleteBtn = card.querySelector(".btn-delete");
    deleteBtn.addEventListener("click", () => {
      if (slidesContainer.children.length > 1) {
        card.remove();
        renumberSlides();
      }
    });
  }

  function renumberSlides() {
    slidesContainer.querySelectorAll(".slide-card").forEach((card, i) => {
      const num = i + 1;
      card.querySelector(".slide-number").textContent = num;
      card.querySelector(".slide-label").textContent = `Slide ${num}`;
    });
  }

  async function convertToPptx() {
    const editors = Array.from(slidesContainer.querySelectorAll(".code-editor"));
    const slides = editors.map(e => e.value.trim()).filter(h => h);

    if (slides.length === 0) return alert("Please enter HTML code.");

    setConverting(true);
    const pres = new PptxGenJS();
    pres.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
    pres.layout = "WIDE";

    try {
      for (let i = 0; i < slides.length; i++) {
        updateProgress(i + 1, slides.length, `Capturing Slide ${i + 1}...`);
        const imgDataUrl = await renderHtmlToImage(slides[i]);
        const slide = pres.addSlide();

        // Ensure image fits perfectly on the WIDE slide (13.333 x 7.5 inches)
        slide.addImage({
          data: imgDataUrl,
          x: 0,
          y: 0,
          w: 13.333,
          h: 7.5,
          sizing: { type: "contain", w: 13.333, h: 7.5 }
        });
      }
      await pres.writeFile({ fileName: "genspark-export.pptx" });
    } catch (err) {
      console.error("CAPTURE ERROR:", err);
      alert("Error: " + err.message);
    } finally {
      setConverting(false);
    }
  }

  function renderHtmlToImage(htmlContent) {
    return new Promise((resolve, reject) => {
      const wrapper = document.createElement("div");
      wrapper.style.cssText = "position:absolute;left:-9999px;top:0;width:1280px;height:720px;overflow:hidden;background:#020230;";
      const iframe = document.createElement("iframe");
      iframe.style.cssText = "width:1280px;height:720px;border:none;";
      wrapper.appendChild(iframe);
      document.body.appendChild(wrapper);

      iframe.onload = async () => {
        const doc = iframe.contentDocument;
        doc.open();
        if (htmlContent.toLowerCase().includes("<html")) doc.write(htmlContent);
        else doc.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{margin:0;padding:0;width:1280px;height:720px;}</style></head><body>${htmlContent}</body></html>`);
        doc.close();

        try {
          // Robust check for library
          const h2i = window.htmlToImage || window.html_to_image;
          if (!h2i) throw new Error("Library 'html-to-image' not found. Please check your internet connection.");

          // Wait for fonts & safety delay
          await iframe.contentWindow.document.fonts.ready;
          await new Promise(r => setTimeout(r, 2000));

          // Capture
          const dataUrl = await h2i.toPng(doc.body, {
            width: 1280,
            height: 720,
            pixelRatio: 2
          });

          wrapper.remove();
          resolve(dataUrl);
        } catch (err) {
          wrapper.remove();
          reject(err);
        }
      };
      iframe.src = "about:blank";
    });
  }

  function setConverting(on) {
    btnConvert.disabled = on;
    btnConvertSpinner.classList.toggle("hidden", !on);
    btnConvertText.classList.toggle("hidden", on);
    progressOverlay.classList.toggle("hidden", !on);
  }

  function updateProgress(cur, total, msg) {
    progressBar.style.width = (cur / total * 100) + "%";
    progressDetail.textContent = msg;
  }
})();
