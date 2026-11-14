namespace DeviceMonitorDemo.Models;

public sealed class DeviceStatusEventArgs(DeviceStatus status) : EventArgs
{
    public DeviceStatus Status { get; } = status;
}
