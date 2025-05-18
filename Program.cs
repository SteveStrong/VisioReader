using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using System.Collections.Generic;
using FoundryRulesAndUnits.Extensions;
using System.Text.Json.Serialization;

public static class ShapeRegistry
{
    public static Dictionary<string, Shape1D> Shape1DMap { get; } = new();
    public static Dictionary<string, Shape2D> Shape2DMap { get; } = new();
}

class Program
{
    static void Main(string[] args)
    {
        var rootDirectory = Directory.GetCurrentDirectory();
        rootDirectory = rootDirectory.Replace(@"bin\Debug\net9.0", "");

        string drawingDir = Path.Combine(rootDirectory, "Drawings");
        if (!Directory.Exists(drawingDir))
        {
            Console.WriteLine($"Drawing directory not found: {drawingDir}");
            return;
        }
        foreach (var vsdxPath in Directory.GetFiles(drawingDir, "*.vsdx"))
        {
            string drawingName = Path.GetFileNameWithoutExtension(vsdxPath);
            string outputDir = Path.Combine(rootDirectory, drawingName);
            Directory.CreateDirectory(outputDir);

            //clear out the global shape registry
            ShapeRegistry.Shape1DMap.Clear();
            ShapeRegistry.Shape2DMap.Clear();
            // Process each VSDX file

            $"Processing: {vsdxPath}".WriteNote(1);
            var shapes = ExtractShapesFromVsdx(vsdxPath);

            //place the shape from the global registry into the shapes list
            //into an instance of ShapeKnowledgeModel
            var shapeKnowledgeModel = new ShapeKnowledgeModel
            {
                Filename = drawingName,
                Shape1DList = ShapeRegistry.Shape1DMap.Values.ToList(),
                Shape2DList = ShapeRegistry.Shape2DMap.Values.ToList()
            };



            // Save as C# model (for now, just serialize to JSON for demonstration)
            //define the serilizer options
            var serializerOptions = new System.Text.Json.JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull | JsonIgnoreCondition.WhenWritingDefault,
            };
            //now serialize ShapeKnowledgeModel to JSON
            string json = System.Text.Json.JsonSerializer.Serialize(shapeKnowledgeModel, serializerOptions);

            //now write the JSON to the output directory
            string jsonFilePath = Path.Combine(outputDir, $"{drawingName}.json");
            File.WriteAllText(jsonFilePath, json);
            $"Shape knowledge model saved to: {jsonFilePath}".WriteSuccess(1);
        }
    }

    static List<VisioShape> ExtractShapesFromVsdx(string vsdxPath)
    {
        var shapes = new List<VisioShape>();
        using (var package = Package.Open(vsdxPath, FileMode.Open, FileAccess.Read))
        {            var pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            var xdoc = XDocument.Load(pagePart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            var shapesElement = xdoc.Root?.Element(ns + "Shapes");
            if (shapesElement != null)
            {
                foreach (var shapeElem in shapesElement.Elements(ns + "Shape"))
                {
                    var shape = ParseShape(shapeElem, ns);
                    shapes.Add(shape);
                }
            }
        }
        return shapes;
    }

    static void SortShapeByDimension(VisioShape shape, XElement shapeElem, int depth=0)
    {
        // Prefer Type attribute if present
        bool is1D = shape.Is1DShape();
        if (!is1D)
        {
            // Fallback: Check for a cell named "BeginX" to determine if shape is 1D
            var beginXCell = shape.Cells.Find(c => c.Name == "BeginX");
            is1D = beginXCell != null;
        }        if (is1D)
        {
            // 1D shape
            var s1d = new Shape1D
            {
                ID = shape.ID,
                Name = shape.Name,
                Text = shape.Text,
                Master = shape.Master,
                BeginX = double.TryParse(shape.Cells.Find(c => c.Name == "BeginX")?.Value, out var bx) ? bx : null,
                BeginY = double.TryParse(shape.Cells.Find(c => c.Name == "BeginY")?.Value, out var by) ? by : null,
                EndX = double.TryParse(shape.Cells.Find(c => c.Name == "EndX")?.Value, out var ex) ? ex : null,
                EndY = double.TryParse(shape.Cells.Find(c => c.Name == "EndY")?.Value, out var ey) ? ey : null,
                // Copy connection-related properties
                BeginConnectedTo = shape.BeginConnectedTo,
                EndConnectedTo = shape.EndConnectedTo,
                ConnectionPoints = shape.ConnectionPoints,
                PageName = shape.PageName
            };
            if (shape.ID != null)
                ShapeRegistry.Shape1DMap[shape.ID] = s1d;
            $"Shape 1D {shape.Name}: {shape.Text}".WriteSuccess(depth);
        }
        else
        {            // 2D shape
            var s2d = new Shape2D
            {
                ID = shape.ID,
                Name = shape.Name,
                Text = shape.Text,
                Master = shape.Master,
                PinX = double.TryParse(shape.Cells.Find(c => c.Name == "PinX")?.Value, out var px) ? px : null,
                PinY = double.TryParse(shape.Cells.Find(c => c.Name == "PinY")?.Value, out var py) ? py : null,
                Width = double.TryParse(shape.Cells.Find(c => c.Name == "Width")?.Value, out var w) ? w : null,
                Height = double.TryParse(shape.Cells.Find(c => c.Name == "Height")?.Value, out var h) ? h : null,
                // Set the page name for 2D shapes as well
                PageName = shape.PageName
            };            if (shape.ID != null)
                ShapeRegistry.Shape2DMap[shape.ID] = s2d;
            $"Shape 2D {shape.Name}: {shape.Text}".WriteWarning(depth);
        }
    }    static VisioShape ParseShape(XElement shapeElem, XNamespace ns, int depth=0)
    {
        // also add Master and Type to the Shape
        var shape = new VisioShape
        {
            ID = shapeElem.Attribute("ID")?.Value,
            Name = shapeElem.Attribute("Name")?.Value,
            Type = shapeElem.Attribute("Type")?.Value,
            Master = shapeElem.Attribute("Master")?.Value,
            Text = shapeElem.Element(ns + "Text")?.Value?.Trim(),
            // No default string initializations as requested
            // Capture page name from parent document context
            PageName = GetPageName(shapeElem)
        };        foreach (var cellElem in shapeElem.Elements(ns + "Cell"))
        {
            shape.Cells.Add(new VisioCell
            {
                Name = cellElem.Attribute("N")?.Value,
                Value = cellElem.Attribute("V")?.Value,
                U = cellElem.Attribute("U")?.Value,
                F = cellElem.Attribute("F")?.Value
            });
        }
        foreach (var sectionElem in shapeElem.Elements(ns + "Section"))
        {
            var section = new VisioSection { Name = sectionElem.Attribute("N")?.Value ?? string.Empty };
            foreach (var rowElem in sectionElem.Elements(ns + "Row"))
            {
                var row = new VisioRow
                {
                    IX = rowElem.Attribute("IX")?.Value ?? string.Empty,
                    Name = rowElem.Attribute("N")?.Value ?? string.Empty
                };
                foreach (var cellElem in rowElem.Elements(ns + "Cell"))
                {
                    row.Cells.Add(new VisioCell
                    {
                        Name = cellElem.Attribute("N")?.Value ?? string.Empty,
                        Value = cellElem.Attribute("V")?.Value ?? string.Empty,
                        U = cellElem.Attribute("U")?.Value ?? string.Empty,
                        F = cellElem.Attribute("F")?.Value ?? string.Empty
                    });
                }
                section.Rows.Add(row);
            }            shape.Sections.Add(section);
        }
        
        // Extract connection information for the shape
        if (shape.Is1DShape())
        {
            ExtractConnectionInfo(shape, shapeElem, ns);
        }
        
        // Recursively process subshapes
        var shapesElem = shapeElem.Element(ns + "Shapes");
        if (shapesElem != null)
        {
            foreach (var subShapeElem in shapesElem.Elements(ns + "Shape"))
            {
                var subshape = ParseShape(subShapeElem, ns, depth+1);
                shape.Children.Add(subshape);
            }
        }

        //call the Shape2D and Shape1D sorter
        SortShapeByDimension(shape, shapeElem, depth);
        return shape;
    }

    static string GetPageName(XElement shapeElem)
    {
        // Try to get the page name from the parent document
        // This is a simplification - in a full implementation you might need to navigate 
        // through the package structure to get the actual page name
        try
        {
            // Get the closest page element or URI that contains the page name
            // For now, returning a default page name or extracting from parent path
            var ancestorPage = shapeElem.Ancestors().FirstOrDefault(a => a.Name.LocalName == "Page");
            if (ancestorPage != null)
            {
                return ancestorPage.Attribute("Name")?.Value ?? "Unknown";
            }

            // Use owning part URI as fallback - extract page name from URI path
            // Example: /visio/pages/page1.xml -> page1
            return "Page1";  // Default for now, can be enhanced with actual path parsing
        }
        catch
        {
            return "Unknown";
        }
    }

    static void ExtractConnectionInfo(VisioShape shape, XElement shapeElem, XNamespace ns)
    {
        try
        {
            // Extract connection information from the shape's connections
            // This would require analyzing the Connect elements in the document
            
            // Find the shape's connections by looking at Connects sections or specific cells
            var connectSection = shape.Sections.FirstOrDefault(s => s.Name == "Connects");
            if (connectSection != null)
            {
                // Process each connection in the Connects section
                foreach (var row in connectSection.Rows)
                {
                    // Check if this is a beginning connection
                    var fromPartCell = row.Cells.FirstOrDefault(c => c.Name == "FromPart");
                    var toSheetCell = row.Cells.FirstOrDefault(c => c.Name == "ToSheet");
                    
                    if (fromPartCell != null && toSheetCell != null)
                    {
                        // FromPart = 3 typically indicates the beginning of a 1D shape
                        if (fromPartCell.Value == "3")
                        {
                            shape.BeginConnectedTo = toSheetCell.Value;
                        }
                        // FromPart = 6 typically indicates the end of a 1D shape
                        else if (fromPartCell.Value == "6")
                        {
                            shape.EndConnectedTo = toSheetCell.Value;
                        }
                    }
                    
                    // Collect connection point information
                    var connectionPointInfo = new List<string>();
                    foreach (var cell in row.Cells)
                    {
                        connectionPointInfo.Add($"{cell.Name}:{cell.Value}");
                    }
                    
                    // Store all connection points information as a semicolon-delimited list
                    if (connectionPointInfo.Count > 0)
                    {
                        if (string.IsNullOrEmpty(shape.ConnectionPoints))
                        {
                            shape.ConnectionPoints = string.Join(";", connectionPointInfo);
                        }
                        else
                        {
                            shape.ConnectionPoints += "|" + string.Join(";", connectionPointInfo);
                        }
                    }
                }
            }
            else
            {
                // Alternative approach: look for connection-related cells directly
                // This is a fallback if the Connects section is not found
                
                // Look for specific cells that might indicate connections
                var beginConnectCell = shape.Cells.FirstOrDefault(c => c.Name == "BeginConnect" || c.Name == "BeginConnectTo");
                var endConnectCell = shape.Cells.FirstOrDefault(c => c.Name == "EndConnect" || c.Name == "EndConnectTo");
                
                if (beginConnectCell != null)
                {
                    shape.BeginConnectedTo = beginConnectCell.Value;
                }
                
                if (endConnectCell != null)
                {
                    shape.EndConnectedTo = endConnectCell.Value;
                }
            }
        }
        catch (Exception ex)
        {
            // Log exception but continue processing
            Console.WriteLine($"Error extracting connection info for shape {shape.ID}: {ex.Message}");
        }
    }
}
