using System.Threading.Channels;
using DeviceMonitorDemo.Models;

namespace DeviceMonitorDemo.Devices;

public sealed class VisionDevice(ChannelWriter<byte[]> writer) : SimulatedDevice("VISION-01", DeviceKind.Vision, writer)
{
    protected override DeviceStatus CreateStatus(int sample) => new(
        Name,
        Kind,
        sample % 11 == 0 ? DeviceState.Warning : DeviceState.Running,
        0.85 + (sample % 15) / 100.0,
        DateTimeOffset.UtcNow);
}
