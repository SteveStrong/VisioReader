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
    
    // Enhanced connection information
    public string? BeginConnectedName { get; set; }  // Name of the shape connected at the beginning
    public string? EndConnectedName { get; set; }    // Name of the shape connected at the end
    public string? FromPart { get; set; }           // FromPart value from the connection
    public string? ToPart { get; set; }             // ToPart value from the connection
    public string? PageName { get; set; }
    
    // Helper method to populate connection names
    public void PopulateConnectionNames(Dictionary<string, Shape2D> shape2DMap)
    {
        if (!string.IsNullOrEmpty(BeginConnectedTo) && shape2DMap.ContainsKey(BeginConnectedTo))
        {
            BeginConnectedName = shape2DMap[BeginConnectedTo].Name;
            
            // Include connected shape text if available
            if (!string.IsNullOrEmpty(shape2DMap[BeginConnectedTo].Text))
            {
                BeginConnectedName += $" ({shape2DMap[BeginConnectedTo].Text})";
            }
        }
        
        if (!string.IsNullOrEmpty(EndConnectedTo) && shape2DMap.ContainsKey(EndConnectedTo))
        {
            EndConnectedName = shape2DMap[EndConnectedTo].Name;
            
            // Include connected shape text if available
            if (!string.IsNullOrEmpty(shape2DMap[EndConnectedTo].Text))
            {
                EndConnectedName += $" ({shape2DMap[EndConnectedTo].Text})";
            }
        }
    }
}
