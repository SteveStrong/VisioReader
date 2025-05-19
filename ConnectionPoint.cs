using System.Text.Json.Serialization;

public class ConnectionPoint
{
    // Basic properties
    public string Id { get; set; } = string.Empty;
    
    public string Name { get; set; } = string.Empty;
    
    // Connection properties - specifically for Shape1D connections
    public string FromSheet { get; set; } = string.Empty;   // The source shape ID
    
    public string ToSheet { get; set; } = string.Empty;     // The target shape ID
    
    public string FromPart { get; set; } = string.Empty;    // FromPart value from the connection
    
    public string ToPart { get; set; } = string.Empty;      // ToPart value from the connection
    
    // Enhanced properties for better readability
    public string FromShapeName { get; set; } = string.Empty;  // Name of the source shape
    
    public string ToShapeName { get; set; } = string.Empty;    // Name of the target shape
    
    // Positional properties (from original ConnectionPoint)
    public double? X { get; set; }
    
    public double? Y { get; set; }
    
    public string DirX { get; set; } = string.Empty;
    
    public string DirY { get; set; } = string.Empty;
}
