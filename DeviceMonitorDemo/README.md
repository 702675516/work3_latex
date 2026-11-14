# DeviceMonitorDemo

A small, hardware-free console project for learning C# through a simulated industrial device monitor. It targets .NET 10 and uses only the .NET libraries.

## Run

```powershell
dotnet run
```

Press Enter to stop. The PLC, motion controller, and vision device are simulations; no serial port or TCP device is required.

## Read in this order

1. `Models/DeviceStatus.cs` and `Models/Point2D.cs`: enums, a reference-type record, and a value-type struct.
2. `Devices/IDevice.cs`: the device contract. `Program.cs` stores concrete devices in `List<IDevice>` and injects them into `DeviceMonitor`.
3. `Devices/SimulatedDevice.cs`: encapsulation, abstract members, a virtual method, event invocation, locking, asynchronous I/O simulation, a per-device command semaphore, and async disposal.
4. `Devices/PlcDevice.cs`, `MotionController.cs`, and `VisionDevice.cs`: inheritance, overrides, and polymorphism. The classes are sealed because they are the final simulated device types.
5. `Monitoring/DeviceMonitor.cs`: event subscription/unsubscription, a Lambda handler, `ConcurrentQueue<T>`, `ConcurrentDictionary<TKey,TValue>`, and `IDisposable`.
6. `Pipeline/DeviceDataPipeline.cs`: `Channel<byte[]>`, `await foreach`, `TaskCompletionSource<T>`, parsing, enums, and exception handling.
7. `Program.cs`: `var`, `const`, arrays/collections, `Task.WhenAll`, `Task.WhenAny`, `CancellationToken`, `Task.Run`, and orderly shutdown.

## Important distinctions

- `int`, `double`, `bool`, enums, and structs are value types. A normal assignment copies their value. `Point2D` is an example.
- Classes, arrays, interfaces, delegates, `object`, and `string` are reference types. A variable holds a reference to an object; assignment copies the reference. The object is typically allocated on the managed heap. `DeviceStatus` is a record class.
- `ReadStatusAsync` uses `Task.Delay` to stand in for waiting on a device. `await` does not itself create a thread. Real socket/file/database APIs should use their asynchronous I/O methods.
- `Task.Run` in `Program.cs` demonstrates offloading CPU work. For real CPU-heavy calculations, consider `Task.Run`, `Parallel`, or a worker pool; avoid wrapping naturally asynchronous I/O in `Task.Run`.
- `SemaphoreSlim` serializes commands sent to one simulated device. For a shared serial port, the semaphore would guard the actual write/read transaction.
- `lock` protects the last-status field. `Monitor` is the lower-level API used by `lock`; `Mutex` can coordinate across processes. `Interlocked` handles simple atomic counters without a lock.
- The channel is unbounded to keep the example easy to follow. A long-running production receiver should usually use a bounded channel and define what to do when it fills.
- `Thread` creates and manages a thread directly; `ThreadPool` reuses worker threads; `Task` represents asynchronous work. Prefer Task-based APIs for application code.
- `readonly` is used for fields that should not be reassigned. Add `const` for compile-time fixed values and `readonly` fields for values initialized at runtime.
- `CancellationTokenSource` is disposed by `using`; devices implement `IAsyncDisposable` and release their semaphore in `DisposeAsync`.

## Scope

This is a learning simulation, not a real industrial protocol driver. It demonstrates the listed concepts while omitting hardware discovery, reconnection policy, UI-thread dispatch, protocol framing edge cases, and production logging.
