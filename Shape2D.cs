public class Shape2D
{
    public string? ID { get; set; }
    public string? Name { get; set; }
    public string? Text { get; set; }
    public string? Master { get; set; }
    public double? PinX { get; set; }
    public double? PinY { get; set; }
    public double? Width { get; set; }
    public double? Height { get; set; }
    
    // Include page name for all shapes
    public string? PageName { get; set; }
}
