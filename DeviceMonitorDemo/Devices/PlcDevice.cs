using System.Threading.Channels;
using DeviceMonitorDemo.Models;

namespace DeviceMonitorDemo.Devices;

// sealed 阻止继续继承；PLC 的专属行为在 override 中实现。
public sealed class PlcDevice(ChannelWriter<byte[]> writer) : SimulatedDevice("PLC-01", DeviceKind.Plc, writer)
{
    protected override DeviceStatus CreateStatus(int sample) => new(
        Name,
        Kind,
        sample % 7 == 0 ? DeviceState.Warning : DeviceState.Running,
        20 + sample % 80,
        DateTimeOffset.UtcNow);
}
