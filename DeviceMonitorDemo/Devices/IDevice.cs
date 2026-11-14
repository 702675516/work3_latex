using DeviceMonitorDemo.Models;

namespace DeviceMonitorDemo.Devices;

// 接口只描述设备“能做什么”；调用方无需知道具体设备类。
public interface IDevice : IAsyncDisposable
{
    string Name { get; }
    DeviceKind Kind { get; }
    bool IsConnected { get; }

    // event 限制外部只能订阅/取消订阅，只有设备类本身能触发事件。
    event EventHandler<DeviceStatusEventArgs>? StatusChanged;

    Task ConnectAsync(CancellationToken cancellationToken);
    Task DisconnectAsync();
    Task<DeviceStatus> ReadStatusAsync(CancellationToken cancellationToken);
    Task SendCommandAsync(string command, CancellationToken cancellationToken);
}
