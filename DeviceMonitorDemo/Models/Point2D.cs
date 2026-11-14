namespace DeviceMonitorDemo.Models;

// struct 是值类型；readonly 让创建后的坐标不可变，赋值时复制坐标值。
public readonly struct Point2D(double x, double y)
{
    public double X { get; } = x;
    public double Y { get; } = y;
}
