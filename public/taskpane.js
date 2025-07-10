document.getElementById("trackerForm").addEventListener("submit", async e => {
    e.preventDefault();
    const email = e.target.email.value;
    const password = e.target.password.value;
    const tenant = e.target.tenant.value;
    const out = document.getElementById("output");
    out.textContent = "Authentifiziere...";
    try {
        const resp = await fetch(`https://${tenant}/oauth2/token`, {
            method: "POST",
            headers: {"Content-Type": "application/x-www-form-urlencoded", "origin": `https://${tenant}`},
            body: new URLSearchParams({
                grant_type: "password", username: email, password: password,
                client_id: "outlook-addin", client_secret: "demo"
            })
        });
        const t = await resp.json();
        if (!resp.ok) throw new Error(t.error_description || "Token Error");
        out.textContent = "Token erhalten. Sende Eintrag...";
        const item = Office.context.mailbox.item;
        const start = item.start.toISOString(), end = item.end.toISOString();
        const title = item.subject;
        const r = await fetch(`https://${tenant}/v1/time_entries`, {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${t.access_token}`,
                "Content-Type": "application/json",
                "origin": `https://${tenant}`
            },
            body: JSON.stringify({description: title, start_time: start, end_time: end})
        });
        if (!r.ok) throw new Error("Tracking Error");
        out.textContent = "✅ Erfasst!";
    } catch (err) {
        out.textContent = "❌ " + err.message;
    }
});