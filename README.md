# Monitor to DeckLink Output

A Windows WPF application that captures any selected monitor and outputs it via a Blackmagic DeckLink card (Duo, Quad, etc).

---

## Quickest Way to Build

### Step 1 — Install prerequisites (one time)

| Tool | Download |
|------|----------|
| .NET 6 SDK | https://dotnet.microsoft.com/download/dotnet/6.0 |
| Blackmagic Desktop Video | https://www.blackmagicdesign.com/support/family/capture-and-playback |

### Step 2 — Clone the repo

```bat
git clone https://github.com/cedargroveleedsmedia/MonitorToDeckLink.git
cd MonitorToDeckLink
```

### Step 3 — Run the build script

```bat
build.bat
```

That's it. The script will:
- Restore NuGet packages
- Compile a single self-contained `MonitorToDeckLink.exe`
- Drop it in the `publish\` folder
- Offer to launch it immediately

---

## Building in Visual Studio (alternative)

1. Install **Visual Studio Community 2022** (free) with the **.NET desktop development** workload
2. Open `MonitorToDeckLink.csproj`
3. Right-click project → **Add → COM Reference** → search **DeckLink** → check it → OK
4. Set platform to **x64**
5. **Build → Publish** using the `Release-x64` profile, or just press **F5** to run

---

## How it Works

- Uses **DXGI Desktop Duplication** to capture any monitor at full speed via the GPU
- Feeds frames into the **DeckLink Output API** via COM interop on a scheduled playback queue
- Scales the captured image to match the selected output resolution if needed
- Color converts BGRA (screen) → UYVY 4:2:2 (DeckLink) in BT.709

## Supported Output Formats

1080p 23.98 / 24 / 25 / 29.97 / 30 / 50 / 59.94 / 60  
1080i 50 / 59.94 / 60  
720p 50 / 59.94 / 60  
2160p 23.98 / 25 / 29.97 / 30  

## Notes

- The DeckLink Duo appears as **two separate devices** — run two instances of the app to drive both outputs simultaneously
- Run as Administrator if monitor capture is denied
- Frame rate is matched to the selected DeckLink output mode
