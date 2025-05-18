public class Shape1D
{
    public string? ID { get; set; }
    public string? Name { get; set; }
    public string? Text { get; set; }
    public string? Master { get; set; }
    public double? BeginX { get; set; }
    public double? BeginY { get; set; }
    public double? EndX { get; set; }
    public double? EndY { get; set; }
    
    // Connection-related properties
    public string? BeginConnectedTo { get; set; }
    public string? EndConnectedTo { get; set; }
    public string? ConnectionPoints { get; set; }
    public string? PageName { get; set; }
}
