# Monitor to DeckLink Output

A Windows WPF application that captures any selected monitor and outputs it via a Blackmagic DeckLink Duo card.

## Requirements

- Windows 10/11 (64-bit)
- .NET 6.0 or later
- Blackmagic Desktop Video software installed (provides DeckLink COM SDK)
- Blackmagic DeckLink Duo (or compatible DeckLink card)
- Visual Studio 2022 (to build)

## Setup

1. Install **Blackmagic Desktop Video** from https://www.blackmagicdesign.com/support/
2. Open `MonitorToDeckLink.sln` in Visual Studio 2022
3. Build in Release mode (x64)
4. Run `MonitorToDeckLink.exe`

## How it Works

- Uses **Windows Graphics Capture API** (DXGI) to capture any monitor
- Feeds frames into the **DeckLink Output API** via COM interop
- Supports common output formats: 1080p25, 1080p29.97, 1080p30, 1080p50, 1080p59.94, 1080p60, 2160p25, 2160p30, 2160p60
- Color conversion from BGRA (screen) to UYVY (DeckLink) is done on CPU

## Notes

- The DeckLink COM interop files (`DeckLinkAPI_x64.tlb`) are installed with Desktop Video
- Recommended: run as Administrator for full monitor capture access
- Frame rate is matched to the selected DeckLink output mode
