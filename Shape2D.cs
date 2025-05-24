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

    // ConnectionPoint objects for this shape
    // This will be populated with connections where this shape is a target
    public List<ConnectionPoint>? IncomingConnections { get; set; }
    
    public void AddIncomingConnection(ConnectionPoint connection)
    {
        if (IncomingConnections == null)
        {
            IncomingConnections = new List<ConnectionPoint>();
        }
        IncomingConnections.Add(connection);
    }
}
