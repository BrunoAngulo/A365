const DASHBOARD_SOURCE = "a365-dashboard";
const EXTENSION_SOURCE = "a365-extension";

function notifyReady() {
  window.postMessage(
    {
      source: EXTENSION_SOURCE,
      type: "A365_EXTENSION_READY",
    },
    window.location.origin,
  );
}

window.addEventListener("message", (event) => {
  if (event.source !== window || event.data?.source !== DASHBOARD_SOURCE) {
    return;
  }

  if (event.data.type === "A365_EXTENSION_PING") {
    notifyReady();
    return;
  }

  if (event.data.type !== "A365_MATRIX_REPORT_REQUEST") {
    return;
  }

  chrome.runtime.sendMessage(
    {
      type: "A365_MATRIX_REPORT_REQUEST",
      requestId: event.data.requestId,
      payload: event.data.payload,
    },
    (response) => {
      if (chrome.runtime.lastError) {
        window.postMessage(
          {
            source: EXTENSION_SOURCE,
            type: "A365_MATRIX_REPORT_RESPONSE",
            requestId: event.data.requestId,
            ok: false,
            error: chrome.runtime.lastError.message || "La extension no pudo comunicarse con el navegador.",
          },
          window.location.origin,
        );
        return;
      }

      window.postMessage(
        {
          source: EXTENSION_SOURCE,
          type: "A365_MATRIX_REPORT_RESPONSE",
          requestId: event.data.requestId,
          ...(response || {
            ok: false,
            error: "La extension no devolvio respuesta.",
          }),
        },
        window.location.origin,
      );
    },
  );
});

notifyReady();
