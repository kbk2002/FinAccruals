const apiBaseUrl = window.location.origin;

Office.onReady((info) => {
  const officeStatus = document.getElementById("office-status");

  if (info.host === Office.HostType.Excel) {
    officeStatus.textContent = "Office.js loaded inside Excel.";
  } else {
    officeStatus.textContent = "Office.js loaded, but not inside Excel.";
  }

  document.getElementById("btn-health").addEventListener("click", testBackend);
  testBackend();
});

async function testBackend() {
  const apiStatus = document.getElementById("api-status");
  apiStatus.textContent = "Checking backend...";

  try {
    const res = await fetch(`${apiBaseUrl}/api/health`);
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.message || "API check failed");
    }

    apiStatus.textContent = `Backend status: ${data.status}`;
  } catch (err) {
    apiStatus.textContent = `Backend error: ${err.message}`;
  }
}
