document.getElementById("trackerForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const email = document.getElementById("email").value;
  const password = document.getElementById("password").value;
  const tenant = document.getElementById("tenant").value;
  const output = document.getElementById("output");
  output.textContent = "Authentifiziere...";

  try {
    const tokenResp = await fetch(`https://${tenant}/oauth2/token`, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "origin": `https://${tenant}`
      },
      body: new URLSearchParams({
        grant_type: "password",
        username: email,
        password: password,
        client_id: "outlook-addin",
        client_secret: "demo"
      })
    });
    const tokenData = await tokenResp.json();
    if (!tokenResp.ok) throw new Error(tokenData.error_description || "Token Error");

    output.textContent = "Token erhalten. Sende Zeiterfassung...";

    const calendar = Office.context.mailbox.item;
    const title = calendar.subject;
    const start = calendar.start.toISOString();
    const end = calendar.end.toISOString();

    const trackResp = await fetch(`https://${tenant}/v1/time_entries`, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${tokenData.access_token}`,
        "Content-Type": "application/json",
        "origin": `https://${tenant}`
      },
      body: JSON.stringify({ description: title, start_time: start, end_time: end })
    });
    if (!trackResp.ok) throw new Error("Tracking Error");

    output.textContent = "✅ Zeiterfassung erfolgreich!";
  } catch (err) {
    output.textContent = "❌ Fehler: " + err.message;
  }
});
