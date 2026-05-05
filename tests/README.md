# Tests

> [!NOTE]
> These unit tests can run and display results in any ![](../resources/svg/ms-office.svg) **Microsoft Office** program.

This folder contains unit and integration tests for Wasabi. The tests are
written in pure VBA and run inside the VBA IDE without any external tools.

## Prerequisites

Some internal functions of Wasabi must be exposed as `Public` for the tests
to access them. The following functions in `Wasabi.bas` need to be changed
from `Private` to `Public`:

| Function | Purpose |
|:---|:---|
| `Base64Encode` | Test Base64 encoding |
| `SHA1` | Test SHA-1 hashing |
| `GenerateWSKey` | Test WebSocket key generation |
| `ComputeWebSocketAccept` | Test accept key computation |
| `ParseURL` | Test URL parsing |
| `StringToUtf8` | Test UTF-8 conversion |
| `Utf8ToString` | Test UTF-8 to string conversion |
| `BuildWSFrame` | Test frame construction |
| `FillRandomBytes` | (optional) Test randomness |

These functions retain their original behavior and are not part of the
stable public API. Their signatures may change between versions without
notice.

## Running the tests

1. Import `Wasabi.bas` into a new VBA project (Excel is recommended).
2. Import all `.bas` files from this folder.
3. Open the Immediate Window (`Ctrl+G`).
4. Run `RunAllTests` by typing it in the Immediate Window and pressing Enter.

Results are printed to the Immediate Window. A summary line at the end
shows the number of passed and failed tests.
