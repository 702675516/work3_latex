using System.Collections.Concurrent;
using DeviceMonitorDemo.Devices;
using DeviceMonitorDemo.Models;

namespace DeviceMonitorDemo.Monitoring;

// IDisposable 用于解绑事件；否则设备仍持有处理器，可能造成意外调用或对象无法回收。
public sealed class DeviceMonitor : IDisposable
{
    private readonly IReadOnlyList<IDevice> _devices;
    private readonly Dictionary<IDevice, EventHandler<DeviceStatusEventArgs>> _handlers = new();
    private readonly ConcurrentQueue<DeviceStatus> _history = new();
    private readonly ConcurrentDictionary<string, long> _counts = new();

    public DeviceMonitor(IReadOnlyList<IDevice> devices)
    {
        _devices = devices;

        foreach (var device in _devices)
        {
            // Lambda 创建事件处理器；保存它是为了 Dispose 时用同一个委托取消订阅。
            EventHandler<DeviceStatusEventArgs> handler = (_, args) => OnStatusChanged(device, args.Status);
            _handlers.Add(device, handler);
            device.StatusChanged += handler;
        }
    }

    public IReadOnlyCollection<DeviceStatus> History => _history.ToArray();

    private void OnStatusChanged(IDevice device, DeviceStatus status)
    {
        // 多台设备并行上报，使用线程安全集合保存历史和计数。
        _history.Enqueue(status);
        _counts.AddOrUpdate(device.Name, 1, (_, current) => current + 1);
        Console.WriteLine($"[event] {status.Timestamp:HH:mm:ss.fff} {status.DeviceName,-10} {status.State,-7} value={status.Value:F2}");
    }

    public void Dispose()
    {
        foreach (var (device, handler) in _handlers)
        {
            // -= 取消订阅；必须传入注册时保存的那个 handler。
            device.StatusChanged -= handler;
        }

        GC.SuppressFinalize(this);
    }
}
