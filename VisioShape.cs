public class VisioShape
{
    public string? ID { get; set; }
    public string? Name { get; set; }
    public string? Type { get; set; }
    public string? Master { get; set; }

    // Legacy connection properties (kept for backward compatibility)
    public string? BeginConnectedTo { get; set; }
    public string? EndConnectedTo { get; set; }
    public string? ConnectionPoints { get; set; }
    public string? PageName { get; set; }
    public string? Text { get; set; }
    
    // Enhanced connection information (legacy)
    public string? FromPart { get; set; }
    public string? ToPart { get; set; }
    
    // New structured connection information
    public List<ConnectionPoint> Connections { get; set; } = new();

    public List<VisioSection> Sections { get; set; } = new();
    public List<VisioCell> Cells { get; set; } = new();
    public List<VisioShape> Children { get; set; } = new();
    public bool Is1DShape() => Type == "1";
    public bool Is2DShape() => string.IsNullOrEmpty(Type) || Type == "0";
}
