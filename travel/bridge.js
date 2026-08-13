// bridge.js — 分派平台的自動化接口。
// 完全獨立的檔案，不修改 app.js / excel-export.js 任何一行；
// 因為 app.js 是以傳統 <script>（非 module）載入，這裡宣告的程式碼跟 app.js 共用同一個全域作用域，
// 可以直接呼叫 app.js 裡的 state / save / parseFile / analyzeRoads / renderPreview / go 等既有函式，
// 就像使用者自己在畫面上操作一樣，不需要另外重寫一套匯入邏輯。
(function () {
  function reply(event, payload) {
    event.source.postMessage(Object.assign({ source: "traffic-suite-travel" }, payload), event.origin);
  }

  window.addEventListener("message", async function (event) {
    const msg = event.data;
    if (!msg || msg.source !== "traffic-suite-shell" || msg.target !== "travel") return;

    if (msg.type === "ensure-project") {
      try {
        let proj = state.projects.find((p) => p.code === msg.code);
        if (!proj) {
          state.projects.push({ code: msg.code, name: msg.name });
        } else {
          proj.name = msg.name;
        }
        state.activeCode = msg.code;
        await save();
        reply(event, { type: "ensure-project-ack", requestId: msg.requestId, ok: true });
      } catch (e) {
        reply(event, { type: "ensure-project-ack", requestId: msg.requestId, ok: false, error: String(e && e.message ? e.message : e) });
      }
      return;
    }

    if (msg.type === "import-file") {
      try {
        const bytes = Uint8Array.from(atob(msg.fileBase64), function (c) {
          return c.charCodeAt(0);
        });
        const file = new File([bytes], msg.fileName);
        state.activeCode = msg.code;
        const parsed = await parseFile(file, msg.year, msg.quarter, msg.defaultSpeed || 50);
        pending = [parsed];
        analyzeRoads();
        renderPreview();
        go("import");
        reply(event, {
          type: "import-file-ack",
          requestId: msg.requestId,
          fileName: msg.fileName,
          ok: parsed.ok,
          error: parsed.ok ? undefined : parsed.error,
          rowCount: parsed.rows.length,
        });
      } catch (e) {
        reply(event, {
          type: "import-file-ack",
          requestId: msg.requestId,
          fileName: msg.fileName,
          ok: false,
          error: String(e && e.message ? e.message : e),
        });
      }
      return;
    }

    if (msg.type === "ping") {
      reply(event, { type: "pong", requestId: msg.requestId });
    }
  });

  // 告知 shell：這個分頁的橋接腳本已經就緒，可以開始收指令了
  if (window.parent && window.parent !== window) {
    window.parent.postMessage({ source: "traffic-suite-travel", type: "bridge-ready" }, "*");
  }
})();
