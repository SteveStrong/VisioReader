public class VisioShape
{
    public string? ID { get; set; }
    public string? Name { get; set; }
    public string? Type { get; set; }
    public string? Master { get; set; }

    public string? BeginConnectedTo { get; set; }
    public string? EndConnectedTo { get; set; }
    public string? ConnectionPoints { get; set; }
    public string? PageName { get; set; }

    public List<VisioSection> Sections { get; set; } = new();
    public List<VisioCell> Cells { get; set; } = new();
    public string? Text { get; set; }
    public List<VisioShape> Children { get; set; } = new();

    /// <summary>
    /// Returns true if the shape is 1D (Type == "1"), otherwise false.
    /// </summary>
    public bool Is1DShape() => Type == "1";

    /// <summary>
    /// Returns true if the shape is 2D (Type == "0" or missing), otherwise false.
    /// </summary>
    public bool Is2DShape() => string.IsNullOrEmpty(Type) || Type == "0";
}
