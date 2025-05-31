
public class NodeShape
{
    public string? ID { get; set; }
    public string? Text { get; set; }
    public double? PinX { get; set; }
    public double? PinY { get; set; }
}

public class EdgeShape
{
    public string? ID { get; set; }
    public string? Text { get; set; }
    public string? FromNodeID { get; set; }
    public string? ToNodeID { get; set; }
    public string? FromNodeText { get; set; }
    public string? ToNodeText { get; set; }
}

public class NetworkModel
{
    public List<NodeShape> Nodes { get; set; } = new List<NodeShape>();
    public List<EdgeShape> Edges { get; set; } = new List<EdgeShape>();

    public void AddNode(NodeShape node)
    {
        Nodes.Add(node);
    }

    public void AddEdge(EdgeShape edge)
    {
        Edges.Add(edge);
    }
}