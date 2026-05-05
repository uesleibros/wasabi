# Wasabi Example Suite

> [!NOTE]
> These examples are pre-configured to run within the ![](../resources/svg/ms-excel.svg) **Microsoft Excel** environment, as they actively interact with spreadsheets and the Excel application object model.

This folder contains a curated collection of production-ready examples demonstrating how to integrate Wasabi into real-world scenarios. They cover everything from basic live data feeds to advanced non-blocking architectures and strict corporate proxy environments.

## What is included

> [!NOTE]
> <img src="../resources/logo.png" width="20" /> **Wasabi Version used:** [v2.3.1-beta](https://github.com/uesleibros/wasabi/releases/tag/v2.3.1-beta)

| File | Description |
|:---|:---|
| **`Ex01_Binance_Live_Ticker.xlsm`** | Connects to a public crypto stream and updates a spreadsheet cell in real-time. Demonstrates string extraction without freezing the UI. |
| **`Ex02_MQTT_QoS2_Dashboard.xlsm`** | Transforms Excel into a full-duplex MQTT dashboard. Subscribes to topics and guarantees delivery using QoS 2, Ping Jitter, and Offline Queueing. |
| **`Ex03_NonBlocking_UI_OnTime.xlsm`** | The most important architectural pattern. Uses `Application.OnTime` to create a background Event Loop, keeping the Excel UI 100% interactive while listening. |
| **`Ex04_Bot_Command_Interface.xlsm`** | Synchronous Request-Response (RPC) pattern. Sends a command to a server and waits for the exact response before continuing execution. |
| **`Ex05_High_Speed_Batching.xlsm`** | Telemetry and high-throughput logging. Disables Nagle's algorithm (`TCP_NODELAY`) and sends massive arrays of data in a single network burst. |
| **`Ex06_Corporate_Auth_Connection.xlsm`** | Enterprise integration template. Shows strict TLS validation, Custom HTTP Headers (Bearer tokens), subprotocols, and system proxy auto-discovery. |

> [!WARNING]
> If testing `Ex02`, ensure your network allows outbound traffic on port `8084`.

## How to run

1. Download and open any of the `.xlsm` example spreadsheets.
2. **Enable Macros** if prompted by Excel's security warning bar.
3. Run the main public subroutine (e.g., `StartBinanceTicker` or `StartMqttDashboard`) from the Developer tab -> Macros dialog (`Alt + F8`) or directly from the VBA Immediate Window.

## Interpreting the patterns

- **Polling vs. Event Loop** – Wasabi does not use background threads. Examples 1, 2, 4, and 6 use a simple `Do While ... DoEvents` loop, which is great for short scripts but turns the cursor into a loading wheel. **Example 3** demonstrates the definitive way to build complex, long-running WebSocket integrations in Excel.
- **Error Handling** – To keep the examples clean and focused on the API flow, heavy `On Error GoTo` blocks were omitted. In a production environment, ensure you wrap your calls and execute `WebSocketDisconnect` in your teardown routines to prevent handle leaks.
