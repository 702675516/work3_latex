using System.Globalization;
using DeviceMonitorDemo.Devices;
using DeviceMonitorDemo.Models;
using DeviceMonitorDemo.Monitoring;
using DeviceMonitorDemo.Pipeline;

// const 是编译期常量；这里只用来制造一小段 CPU 密集型示例工作。
const int sampleWindow = 250_000;
// using var 会在当前作用域结束时自动 Dispose 这个取消源。
using var cancellation = new CancellationTokenSource();

// 把 Channel 的 Writer 注入每个设备，把 Reader 交给独立的解析管线。
var channel = DeviceDataPipeline.CreateChannel();
var pipeline = new DeviceDataPipeline(channel);
var devices = new List<IDevice>
{
    new PlcDevice(channel.Writer),
    new MotionController(channel.Writer),
    new VisionDevice(channel.Writer)
};

// 这里手动把设备集合传入监控器，是最简单的构造函数依赖注入。
using var monitor = new DeviceMonitor(devices);

Console.WriteLine("Device Monitor Demo");
Console.WriteLine("Press Enter to stop. The devices are simulated; no hardware is needed.\n");

try
{
    // WhenAll 等所有连接任务完成；异步等待 I/O，不会为每个设备专门创建线程。
    await Task.WhenAll(devices.Select(device => device.ConnectAsync(cancellation.Token)));

    var pipelineTask = pipeline.RunAsync(cancellation.Token);
    var pollingTasks = devices.Select(device => PollDeviceAsync(device, cancellation.Token)).ToArray();

    // WhenAny 在首条状态和 3 秒超时之间，先完成哪个就先继续。
    var firstSampleOrTimeout = await Task.WhenAny(
        pipeline.FirstStatus,
        Task.Delay(TimeSpan.FromSeconds(3), cancellation.Token));

    if (firstSampleOrTimeout == pipeline.FirstStatus)
    {
        var firstStatus = await pipeline.FirstStatus;
        Console.WriteLine($"First parsed sample: {firstStatus.DeviceName} = {firstStatus.Value:F2}");
    }

    // 用户按 Enter 后发出取消信号；每个异步循环都收到同一个令牌。
    await Console.In.ReadLineAsync(cancellation.Token);
    cancellation.Cancel();
    channel.Writer.TryComplete();

    await ObserveCancellationAsync(Task.WhenAll(pollingTasks));
    await ObserveCancellationAsync(pipelineTask);
}
catch (OperationCanceledException)
{
    // Cancellation is the normal shutdown path for this demo.
}
catch (Exception exception)
{
    Console.Error.WriteLine($"Application error: {exception.Message}");
}
finally
{
    channel.Writer.TryComplete();
    // finally 确保正常退出或异常退出时都会释放设备资源。
    await Task.WhenAll(devices.Select(device => device.DisposeAsync().AsTask()));
    Console.WriteLine("Devices disconnected. Recent samples: " + monitor.History.Count);
}

// CPU 密集计算才适合用 Task.Run 放到线程池；设备 I/O 本身不需要这样包装。
var cpuResult = await Task.Run(() => Enumerable.Range(1, sampleWindow).Sum(value => (long)value), CancellationToken.None);
Console.WriteLine($"Task.Run CPU-work example completed: {cpuResult.ToString(CultureInfo.InvariantCulture)}");

// 每台设备都有自己的轮询 Task；async 方法返回 Task，调用方可以等待和观察异常。
static async Task PollDeviceAsync(IDevice device, CancellationToken cancellationToken)
{
    try
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            await device.ReadStatusAsync(cancellationToken);
        }
    }
    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
    {
        // Expected when the user stops the monitor.
    }
}

static async Task ObserveCancellationAsync(Task task)
{
    try
    {
        await task;
    }
    catch (OperationCanceledException)
    {
        // Expected during shutdown.
    }
}
