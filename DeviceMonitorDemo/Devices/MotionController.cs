using System.Threading.Channels;
using DeviceMonitorDemo.Models;

namespace DeviceMonitorDemo.Devices;

public sealed class MotionController(ChannelWriter<byte[]> writer) : SimulatedDevice("MOTION-01", DeviceKind.MotionController, writer)
{
    protected override DeviceStatus CreateStatus(int sample) => new(
        Name,
        Kind,
        sample % 9 == 0 ? DeviceState.Warning : DeviceState.Running,
        Math.Sin(sample / 5.0) * 100,
        DateTimeOffset.UtcNow);
}
