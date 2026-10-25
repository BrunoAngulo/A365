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

  if (!["A365_MATRIX_REPORT_REQUEST", "A365_INBOUND_REPORT_REQUEST"].includes(event.data.type)) {
    return;
  }

  const responseType = event.data.type === "A365_INBOUND_REPORT_REQUEST"
    ? "A365_INBOUND_REPORT_RESPONSE"
    : "A365_MATRIX_REPORT_RESPONSE";

  chrome.runtime.sendMessage(
    {
      type: event.data.type,
      requestId: event.data.requestId,
      payload: event.data.payload,
    },
    (response) => {
      if (chrome.runtime.lastError) {
        window.postMessage(
          {
            source: EXTENSION_SOURCE,
            type: responseType,
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
          type: responseType,
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
