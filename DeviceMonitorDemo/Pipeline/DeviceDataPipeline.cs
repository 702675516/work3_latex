using System.Globalization;
using System.Text;
using System.Threading.Channels;
using DeviceMonitorDemo.Models;

namespace DeviceMonitorDemo.Pipeline;

// 单消费者管线：设备生产 byte[]，解析器按写入顺序逐条读取。
public sealed class DeviceDataPipeline(Channel<byte[]> channel)
{
    // TaskCompletionSource 允许消费者主动完成一个 Task，通知“首条状态已解析”。
    private readonly TaskCompletionSource<DeviceStatus> _firstStatus =
        new(TaskCreationOptions.RunContinuationsAsynchronously);

    public Task<DeviceStatus> FirstStatus => _firstStatus.Task;

    // SingleReader/SingleWriter 是优化提示；此处多个设备写、一个管线读。
    // 无界 Channel 简单易懂，但生产程序应评估有界容量和满载时的处理策略。
    public static Channel<byte[]> CreateChannel() => Channel.CreateUnbounded<byte[]>(
        new UnboundedChannelOptions { SingleReader = true, SingleWriter = false });

    public async Task RunAsync(CancellationToken cancellationToken)
    {
        // await foreach 异步等待新数据，不需要轮询，也不会一直占用线程。
        await foreach (var frame in channel.Reader.ReadAllAsync(cancellationToken))
        {
            try
            {
                var status = Parse(frame);
                _firstStatus.TrySetResult(status);
                Console.WriteLine($"[channel] parsed {status.DeviceName,-10} {status.Kind,-16} value={status.Value:F2}");
            }
            catch (FormatException exception)
            {
                Console.Error.WriteLine($"Bad device frame: {exception.Message}");
            }
            catch (ArgumentException exception)
            {
                Console.Error.WriteLine($"Unknown device value: {exception.Message}");
            }
        }
    }

    private static DeviceStatus Parse(byte[] frame)
    {
        var fields = Encoding.UTF8.GetString(frame).Split('|');
        if (fields.Length != 5)
        {
            throw new FormatException("Expected five fields.");
        }

        var kind = Enum.Parse<DeviceKind>(fields[1]);
        var state = Enum.Parse<DeviceState>(fields[2]);
        var value = double.Parse(fields[3], CultureInfo.InvariantCulture);
        var timestamp = DateTimeOffset.FromUnixTimeMilliseconds(long.Parse(fields[4], CultureInfo.InvariantCulture));
        return new DeviceStatus(fields[0], kind, state, value, timestamp);
    }
}
