// Magic Rectangle task pane
const POLL_MS = 300;
const TAG_KEY_LOCKED = "m365copilot.cornerRadius.lockedPx";
const TAG_KEY_DPI    = "m365copilot.cornerRadius.dpi";
let pollTimer = null;
let detectedDpi = 96; // fallback default
function pxToPt(px, dpi) { return (px * 72) / dpi; }

async function detectDpiFromActiveSlide() {
  return await PowerPoint.run(async (context) => {
    const pres = context.presentation;
    pres.load("pageSetup");
    await context.sync();
    const ps = pres.pageSetup;
    ps.load("slideWidth,slideHeight");
    await context.sync();
    const slideWidthPts = ps.slideWidth;
    const slideWidthInches = slideWidthPts / 72.0;
    const active = pres.getActiveSlideOrNullObject();
    await context.sync();
    const pngB64 = active.getImageAsBase64();
    await context.sync();
    const img = new Image();
    img.src = "data:image/png;base64," + pngB64.value;
    const dpi = await new Promise((resolve) => {
      img.onload = () => {
        const pxWidth = img.naturalWidth;
        const inferred = Math.round(pxWidth / slideWidthInches);
        resolve(inferred);
      };
      img.onerror = () => resolve(96);
    });
    detectedDpi = dpi;
    document.getElementById("dpiValue").textContent = `${dpi} dpi`;
    return dpi;
  });
}

async function applyFixedRadiusToSelection(radiusPx, dpi) {
  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/type,width,height,adjustments,count,tags");
    await context.sync();
    const radiusPt = pxToPt(radiusPx, dpi);
    selected.items.forEach((shape) => {
      if (shape.type === PowerPoint.ShapeType.geometricShape && shape.adjustments && shape.adjustments.count > 0) {
        const shortSide = Math.min(shape.width, shape.height);
        const ratio = radiusPt / shortSide;
        shape.adjustments.set(0, ratio);
        shape.tags.add(TAG_KEY_LOCKED, String(radiusPx));
        shape.tags.add(TAG_KEY_DPI, String(dpi));
      }
    });
    await context.sync();
  });
}

async function reapplyLockToSelection() {
  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/type,width,height,adjustments,count,tags");
    await context.sync();
    selected.items.forEach((shape) => {
      if (shape.type === PowerPoint.ShapeType.geometricShape && shape.adjustments && shape.adjustments.count > 0) {
        const lockTag = shape.tags.getItem(TAG_KEY_LOCKED);
        const dpiTag  = shape.tags.getItem(TAG_KEY_DPI);
        lockTag.load("value");
        dpiTag.load("value");
      }
    });
    await context.sync();
    selected.items.forEach((shape) => {
      if (shape.type !== PowerPoint.ShapeType.geometricShape || !shape.adjustments || shape.adjustments.count === 0) return;
      let radiusPx = 0, dpi = detectedDpi || 96;
      try {
        radiusPx = Number(shape.tags.getItem(TAG_KEY_LOCKED).value);
        dpi      = Number(shape.tags.getItem(TAG_KEY_DPI).value) || dpi;
      } catch { return; }
      const radiusPt = pxToPt(radiusPx, dpi);
      const shortSide = Math.min(shape.width, shape.height);
      const ratio = radiusPt / shortSide;
      shape.adjustments.set(0, ratio);
    });
    await context.sync();
  });
}

function enableLock(enabled) {
  const lockToggle = document.getElementById("lockToggle");
  if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
  if (enabled) {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      async () => { if (lockToggle.checked) await reapplyLockToSelection(); }
    );
    pollTimer = setInterval(() => { reapplyLockToSelection().catch(console.error); }, POLL_MS);
  }
}


/* global Office */
Office.onReady(info => {
  if (info.host !== Office.HostType.PowerPoint) return;

  // Expose command functions if you have ribbon commands:
  window.openPane = async (evt) => {
    await Office.addin.showAsTaskpane();
    if (evt && evt.completed) evt.completed();
  };

  // Gate higher sets
  const hasPpt110 = Office.context.requirements.isSetSupported('PowerPointApi', '1.10');
  // ... feature-switch here
});
