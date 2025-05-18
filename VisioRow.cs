public class VisioRow
{
    public string? IX { get; set; }
    public string? Name { get; set; }
    public List<VisioCell> Cells { get; set; } = new();
}
