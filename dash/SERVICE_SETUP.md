# Running the Dashboard as an Always-On Service

## What this gives you

- ✅ Streamlit runs automatically when your laptop starts
- ✅ Restarts automatically if it crashes
- ✅ Accessible to other people on **BG network only** (blocked on public WiFi)
- ✅ Password-protected — even people on the network need the password
- ✅ Runs hidden in the background — no visible window

You'll get a URL like `http://10.100.5.42:8501` that Simone/team can open in their browsers as long as they're on BG WiFi or VPN.

---

## One-time setup (15 minutes)

### Step 1 — Make sure the app runs manually first

Before automating, verify the manual version works:

```powershell
cd C:\Users\BL3914\OneDrive - Beyond Gravity Services AG\Documents\Development\dash
# Activate your (quality) venv
streamlit run app.py
```

Open `http://localhost:8501` in your browser. You should see the password setup screen. **Set your password now** — remember it, you'll need it every time.

Stop Streamlit (Ctrl+C in the terminal) once you've confirmed it works.

### Step 2 — Adjust `run.bat` to point to your venv

Open `run.bat` in Notepad. Find this line:
```
call C:\Users\%USERNAME%\quality\Scripts\activate.bat
```
Change it to the actual path of your `(quality)` venv.

To find it: in your terminal with venv active, type:
```powershell
where python
```
The path shown ends in `\Scripts\python.exe`. Use everything up to `\Scripts\activate.bat`.

Example: if `where python` shows:
```
C:\Users\BL3914\quality\Scripts\python.exe
```
Then the line should be:
```
call C:\Users\BL3914\quality\Scripts\activate.bat
```

### Step 3 — Open the firewall (Administrator PowerShell)

1. Press Windows key, type "PowerShell"
2. Right-click **Windows PowerShell** → **Run as administrator**
3. Navigate to your dash folder:
   ```powershell
   cd "C:\Users\BL3914\OneDrive - Beyond Gravity Services AG\Documents\Development\dash"
   ```
4. Run:
   ```powershell
   .\open_firewall.ps1
   ```

If PowerShell blocks the script:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force
.\open_firewall.ps1
```

You should see: `SUCCESS: Firewall rule created.`

### Step 4 — Register the auto-start task

Same folder, but this time you don't need Administrator:

```powershell
.\install_service.ps1
```

You should see: `SUCCESS: Task 'QualityBRMDashboard' registered.`

### Step 5 — Start it right now (don't wait for reboot)

```powershell
Start-ScheduledTask -TaskName QualityBRMDashboard
```

Wait 10 seconds, then open `http://localhost:8501` — you should see your login page.

### Step 6 — Find your BG-network URL

```powershell
ipconfig
```

Look for the line **IPv4 Address** under your active network adapter (WiFi or Ethernet). It'll be something like `10.100.5.42` or `192.168.1.42`.

Your dashboard URL is:
```
http://<that-IP>:8501
```

**Test it:** From another computer on BG WiFi, open that URL. You should see the login page.

---

## Sharing with Simone

Send Simone a message like:

> **Quality BRM dashboard is live at http://10.100.5.42:8501**
> Password: [share separately, not in email]
> Only works on BG WiFi or when connected to BG VPN.
> If it doesn't load: my laptop may be off — ping me.

---

## Daily use

You don't need to do anything. When you turn on your laptop and log in, the dashboard starts by itself. When you close the laptop, it stops. When you reopen, it starts again.

**To update NC data:**
1. Drop new `NCR_Cutover_Tracker_YYYY-MM-DD.xlsx` into `data/`
2. Either in the dashboard sidebar click **"Re-run ingest.py"**, or from PowerShell:
   ```powershell
   cd "C:\...\dash"
   python ingest.py
   ```
3. The dashboard picks up the changes immediately.

---

## Managing the service

| Task | Command |
|------|---------|
| Check status | `Get-ScheduledTask -TaskName QualityBRMDashboard` |
| Start manually | `Start-ScheduledTask -TaskName QualityBRMDashboard` |
| Stop | `Stop-ScheduledTask -TaskName QualityBRMDashboard` |
| Remove entirely | `Unregister-ScheduledTask -TaskName QualityBRMDashboard -Confirm:$false` |
| Change password | Delete `.streamlit_password` file, reload the dashboard |

---

## Security layers you now have

1. **Windows Firewall** — only accepts connections from private (BG) networks
2. **Windows session** — only running while you're logged in
3. **Streamlit password** — even network insiders need the password
4. **No public exposure** — the URL is only reachable from BG's network

**What this is NOT:**
- ❌ Not HTTPS (traffic is unencrypted on the LAN — fine for internal, but don't broadcast the URL widely)
- ❌ Not fault-tolerant to your laptop shutting down / losing WiFi
- ❌ Not multi-user (everyone shares the same password)

For a proper "always-on, HTTPS, multi-user" setup, you'd need IT to host it on an internal server. But for a team of 2-5 people using it during work hours, this is plenty.

---

## Troubleshooting

**Dashboard doesn't open at all**
- Check task is running: `Get-ScheduledTask -TaskName QualityBRMDashboard`
- Start manually and watch the terminal window for errors: `.\run.bat`

**Others can't reach my IP**
- Verify Windows shows the network as "Private" not "Public"
  - Settings → Network → click your network → change to Private
- Confirm they're actually on BG WiFi/VPN (not their phone hotspot)
- Try `ping <your-IP>` from their machine

**Forgot the password**
- Delete `.streamlit_password` in your dash folder
- Reload the dashboard — it will prompt to set a new one

**IP address changes when I switch networks**
- Yes, this happens. You'll need to send the updated URL each time.
- Fix: ask BG IT for a reserved DHCP address, or use your machine's DNS name if BG has it configured.
