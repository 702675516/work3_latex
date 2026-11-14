using System.Globalization;
using System.Text;
using DeviceMonitorDemo.Models;
using System.Threading.Channels;

namespace DeviceMonitorDemo.Devices;

// abstract 基类放共用实现；具体设备通过 override 提供各自的状态数据。
public abstract class SimulatedDevice : IDevice
{
    // readonly 字段只能在声明处或构造函数中重新赋值，避免依赖对象被意外替换。
    private readonly ChannelWriter<byte[]> _writer;
    // lock 保护共享状态；此 object 仅作同步标记，不存业务数据。
    private readonly object _statusLock = new();
    // 同一台设备的命令必须排队发送，避免协议请求交错。
    private readonly SemaphoreSlim _commandGate = new(1, 1);
    private DeviceStatus? _lastStatus;
    private int _connected;
    private int _sampleNumber;
    private int _commandsSent;

    protected SimulatedDevice(string name, DeviceKind kind, ChannelWriter<byte[]> writer)
    {
        Name = name;
        Kind = kind;
        _writer = writer;
    }

    public string Name { get; }
    public DeviceKind Kind { get; }
    // Volatile.Read 让跨线程读取能看到最新连接标志；Interlocked 负责原子写入。
    public bool IsConnected => Volatile.Read(ref _connected) == 1;
    public int CommandsSent => Volatile.Read(ref _commandsSent);

    public event EventHandler<DeviceStatusEventArgs>? StatusChanged;

    public DeviceStatus? LastStatus
    {
        get
        {
            // lock 保证读取和写入 LastStatus 不会同时发生。
            lock (_statusLock)
            {
                return _lastStatus;
            }
        }
    }

    // virtual 表示子类可以在保留基类默认行为的基础上 override。
    public virtual async Task ConnectAsync(CancellationToken cancellationToken)
    {
        // 用延时模拟设备 I/O 等待；真实代码应 await ConnectAsync/ReadAsync 等 I/O API。
        await Task.Delay(100, cancellationToken);
        Interlocked.Exchange(ref _connected, 1);
        Console.WriteLine($"Connected: {Name}");
    }

    public virtual Task DisconnectAsync()
    {
        Interlocked.Exchange(ref _connected, 0);
        Console.WriteLine($"Disconnected: {Name}");
        return Task.CompletedTask;
    }

    public async Task<DeviceStatus> ReadStatusAsync(CancellationToken cancellationToken)
    {
        EnsureConnected();

        // Simulates asynchronous device I/O; it does not create a thread.
        await Task.Delay(Random.Shared.Next(200, 500), cancellationToken);

        // Interlocked 对计数器做原子递增，避免多个调用同时读写时丢失更新。
        var sample = Interlocked.Increment(ref _sampleNumber);
        var status = CreateStatus(sample);
        lock (_statusLock)
        {
            _lastStatus = status;
        }

        // ?. 表示无人订阅时不做事；Invoke 同步通知所有订阅者。
        StatusChanged?.Invoke(this, new DeviceStatusEventArgs(status));
        // 把原始字节写入 Channel，解析工作由另一个消费者负责。
        await _writer.WriteAsync(Encode(status), cancellationToken);
        return status;
    }

    public async Task SendCommandAsync(string command, CancellationToken cancellationToken)
    {
        // WaitAsync 异步等待闸门，不会阻塞当前线程；同一时刻只允许一个命令通过。
        await _commandGate.WaitAsync(cancellationToken);
        try
        {
            EnsureConnected();
            Console.WriteLine($"{Name} received command: {command}");
            await Task.Delay(80, cancellationToken);
            Interlocked.Increment(ref _commandsSent);
        }
        finally
        {
            // 即使发送命令时抛出异常，也必须释放闸门，否则后续命令会永久等待。
            _commandGate.Release();
        }
    }

    // abstract 没有基类实现，要求每个具体设备都定义自己的状态生成规则。
    protected abstract DeviceStatus CreateStatus(int sample);

    private void EnsureConnected()
    {
        if (!IsConnected)
        {
            throw new InvalidOperationException($"Device {Name} is not connected.");
        }
    }

    private static byte[] Encode(DeviceStatus status)
    {
        var line = string.Join('|',
            status.DeviceName,
            status.Kind,
            status.State,
            status.Value.ToString("F2", CultureInfo.InvariantCulture),
            status.Timestamp.ToUnixTimeMilliseconds());
        return Encoding.UTF8.GetBytes(line);
    }

    // async disposal 适合释放需要异步关闭的设备资源；此处用它模拟长期资源生命周期。
    public async ValueTask DisposeAsync()
    {
        await DisconnectAsync();
        _commandGate.Dispose();
        GC.SuppressFinalize(this);
    }
}
