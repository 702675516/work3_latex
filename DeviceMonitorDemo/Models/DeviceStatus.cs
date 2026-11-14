namespace DeviceMonitorDemo.Models;

// enum 用有限选项表达设备类别，避免用任意字符串传递类型。
public enum DeviceKind
{
    Plc,
    MotionController,
    Vision
}

public enum DeviceState
{
    Idle,
    Running,
    Warning
}

// record class 适合承载状态数据；它是引用类型，并提供按值比较等生成行为。
public sealed record DeviceStatus(
    string DeviceName,
    DeviceKind Kind,
    DeviceState State,
    double Value,
    DateTimeOffset Timestamp);
