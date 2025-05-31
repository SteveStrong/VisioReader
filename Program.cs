using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Linq;
using FoundryRulesAndUnits.Extensions;
using System.Text.Json.Serialization;
using System.Text.Json;
using VisioReader; // Add reference to the namespace where EmptyCollectionConverter is defined

public static class ShapeRegistry
{
    public static Dictionary<string, Shape1D> Shape1DMap { get; } = new();
    public static Dictionary<string, Shape2D> Shape2DMap { get; } = new();
}

public class Program
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
            //into an instance of ShapeKnowledgeModel            // Post-processing: Populate connection names for all Shape1D objects
            Console.WriteLine("Post-processing: Populating connection names for all Shape1D objects...");
            
            // First, populate the Shape2D.IncomingConnections lists
            foreach (var shape1D in ShapeRegistry.Shape1DMap.Values)
            {
                // For each connection in the Shape1D
                if (shape1D.Connections == null)
                {
                    continue;
                }
                foreach (var connection in shape1D.Connections)
                {
                    // If the ToSheet refers to a Shape2D, add this connection to that shape's IncomingConnections
                    if (!string.IsNullOrEmpty(connection.ToSheet) && ShapeRegistry.Shape2DMap.ContainsKey(connection.ToSheet))
                    {
                        ShapeRegistry.Shape2DMap[connection.ToSheet].AddIncomingConnection(connection);
                        Console.WriteLine($"Added connection from {shape1D.Name} to Shape2D {ShapeRegistry.Shape2DMap[connection.ToSheet].Name}");
                    }
                }
            }
            
            // Now populate both legacy and new connection information
            foreach (var shape1D in ShapeRegistry.Shape1DMap.Values)
            {
                // Log connection information for debugging
                //Console.WriteLine($"Shape1D ID={shape1D.ID}, Name={shape1D.Name}, ConnectionCount={shape1D.Connections.Count}, BeginConnectedTo={shape1D.BeginConnectedTo ?? "none"}, EndConnectedTo={shape1D.EndConnectedTo ?? "none"}, FromPart={shape1D.FromPart ?? "none"}, ToPart={shape1D.ToPart ?? "none"}");
                
                // Populate both legacy and new connection information
                shape1D.PopulateConnectionNames(ShapeRegistry.Shape2DMap);
                Console.WriteLine($"Populated connections for {shape1D.Name}: BeginConnectedName={shape1D.BeginConnectedName ?? "none"}, EndConnectedName={shape1D.EndConnectedName ?? "none"}");
                

            }
              var shapeKnowledgeModel = new ShapeKnowledgeModel
            {
                Filename = drawingName,
                Shape1DList = ShapeRegistry.Shape1DMap.Values.ToList(),
                Shape2DList = ShapeRegistry.Shape2DMap.Values.ToList()
            };

            // Create and populate NetworkModel with shapes that have text
            var networkModel = CreateNetworkModel();
            Console.WriteLine($"NetworkModel created with {networkModel.Nodes.Count} nodes and {networkModel.Edges.Count} edges");

            // Save as C# model (for now, just serialize to JSON for demonstration)// Define the serializer options
            var serializerOptions = new System.Text.Json.JsonSerializerOptions
            {
                WriteIndented = true,
                // Only ignore null values, but include default values (like empty strings)
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            };
            
            // Add the EmptyCollectionConverter to skip serializing empty collections
            //serializerOptions.Converters.Add(new EmptyCollectionConverter());            // Serialize ShapeKnowledgeModel to JSON
            string json = System.Text.Json.JsonSerializer.Serialize(shapeKnowledgeModel, serializerOptions);

            // Write the JSON to the output directory
            string jsonFilePath = Path.Combine(outputDir, $"{drawingName}.json");
            File.WriteAllText(jsonFilePath, json);
            $"Shape knowledge model saved to: {jsonFilePath}".WriteSuccess(1);            // Serialize NetworkModel to JSON
            string networkJson = System.Text.Json.JsonSerializer.Serialize(networkModel, serializerOptions);
            
            // Write the NetworkModel JSON to the output directory
            string networkJsonFilePath = Path.Combine(outputDir, $"{drawingName}_NetworkModel.json");
            File.WriteAllText(networkJsonFilePath, networkJson);
            $"Network model saved to: {networkJsonFilePath}".WriteSuccess(1);
        }
    }

    static List<VisioShape> ExtractShapesFromVsdx(string vsdxPath)
    {
        var shapes = new List<VisioShape>();
        // Dictionary to store connections: key = FromSheet ID, value = List of connection details (for backward compatibility)
        var connectionsMap = new Dictionary<string, List<(string ToSheet, string FromPart, string ToPart)>>();
        
        // Dictionary to store connection points that will be added to Shape1D objects later
        var shape1DConnectionPointsMap = new Dictionary<string, List<ConnectionPoint>>();
        
        Console.WriteLine("Starting connection extraction...");

        using (var package = Package.Open(vsdxPath, FileMode.Open, FileAccess.Read))
        {
            // Get all page parts from the package
            var pageParts = package.GetParts().Where(part =>
                part.Uri.ToString().StartsWith("/visio/pages/page") &&
                part.Uri.ToString().EndsWith(".xml"));

            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            
            // First pass: Extract all connections from all pages
            foreach (var pagePart in pageParts)
            {
                var xdoc = XDocument.Load(pagePart.GetStream());
                var connects = xdoc.Descendants(ns + "Connect");
                
                foreach (var connect in connects)
                {
                    string fromSheet = connect.Attribute("FromSheet")?.Value ?? "";
                    string toSheet = connect.Attribute("ToSheet")?.Value ?? "";
                    string fromPart = connect.Attribute("FromPart")?.Value ?? "";
                    string toPart = connect.Attribute("ToPart")?.Value ?? "";
                    
                    if (!string.IsNullOrEmpty(fromSheet) && !string.IsNullOrEmpty(toSheet))
                    {
                        // Create a ConnectionPoint object
                        var connectionPoint = new ConnectionPoint
                        {
                            Id = $"{fromSheet}-{toSheet}",
                            FromSheet = fromSheet,
                            ToSheet = toSheet,
                            FromPart = fromPart,
                            ToPart = toPart
                        };
                        
                        // Store the ConnectionPoint object for later use
                        if (!shape1DConnectionPointsMap.ContainsKey(fromSheet))
                        {
                            shape1DConnectionPointsMap[fromSheet] = new List<ConnectionPoint>();
                        }
                        shape1DConnectionPointsMap[fromSheet].Add(connectionPoint);
                        
                        // For backward compatibility, still maintain the old connections map
                        if (!connectionsMap.ContainsKey(fromSheet))
                        {
                            connectionsMap[fromSheet] = new List<(string, string, string)>();
                        }
                        connectionsMap[fromSheet].Add((toSheet, fromPart, toPart));
                        
                        $"Found connection: {fromSheet} -> {toSheet} (FromPart={fromPart}, ToPart={toPart})".WriteSuccess();
                    }
                }
            }
            
            $"Total connections found: {connectionsMap.Count} shapes with connections".WriteSuccess();
            
            // Process all pages and extract all shapes
            foreach (var pagePart in pageParts)
            {
                // Extract page name from the URI - e.g., page1.xml becomes Page1
                string pageName = Path.GetFileNameWithoutExtension(pagePart.Uri.ToString());
                pageName = char.ToUpper(pageName[0]) + pageName.Substring(1);

                var xdoc = XDocument.Load(pagePart.GetStream());
                // Try to get the actual page name from the XML if available
                var pageElement = xdoc.Root;
                if (pageElement?.Attribute("Name") != null)
                {
                    pageName = pageElement.Attribute("Name")?.Value ?? pageName;
                }

                var shapesElement = xdoc.Root?.Element(ns + "Shapes");
                if (shapesElement != null)
                {
                    foreach (var shapeElem in shapesElement.Elements(ns + "Shape"))
                    {
                        var shape = ParseShape(shapeElem, ns);
                        
                        // Ensure the page name is set correctly
                        if (string.IsNullOrEmpty(shape.PageName) || shape.PageName == "Page1" || shape.PageName == "Unknown")
                        {
                            shape.PageName = pageName;
                        }
                        
                        // Apply connection information from the map we built earlier
                        if (shape.ID != null && connectionsMap.TryGetValue(shape.ID, out var connections))
                        {
                            foreach (var (toSheet, fromPart, toPart) in connections)
                            {
                                // For backward compatibility, still maintain the old properties
                                // FromPart = 3, 9 and similar small values typically indicate the beginning of a 1D shape
                                if (fromPart == "3" || fromPart == "9")
                                {
                                    shape.BeginConnectedTo = toSheet;
                                    shape.FromPart = fromPart;
                                }
                                // FromPart = 6, 12 and similar values typically indicate the end of a 1D shape
                                else if (fromPart == "6" || fromPart == "12")
                                {
                                    shape.EndConnectedTo = toSheet;
                                    shape.ToPart = toPart;
                                }
                                
                                // Store detailed connection information for backward compatibility
                                string connectionInfo = $"FromPart:{fromPart};ToPart:{toPart};ToSheet:{toSheet}";
                                if (string.IsNullOrEmpty(shape.ConnectionPoints))
                                {
                                    shape.ConnectionPoints = connectionInfo;
                                }
                                else
                                {
                                    shape.ConnectionPoints += "|" + connectionInfo;
                                }
                                
                                $"Debug: Applied connection for shape {shape.ID}: FromPart={fromPart}, ToPart={toPart}, ToSheet={toSheet}".WriteInfo();
                            }
                        }
                        
                        shapes.Add(shape);
                    }
                }
            }
            
            // After all shapes have been processed and registered, add the ConnectionPoint objects to the Shape1D objects
            Console.WriteLine("Post-processing: Adding ConnectionPoint objects to Shape1D objects...");
            foreach (var shapeID in shape1DConnectionPointsMap.Keys)
            {
                if (ShapeRegistry.Shape1DMap.TryGetValue(shapeID, out var shape1D))
                {
                    shape1D.Connections = shape1DConnectionPointsMap[shapeID];
                    Console.WriteLine($"Added {shape1D.Connections.Count} ConnectionPoint objects to Shape1D {shapeID}");
                }
            }
            
            // Print count of shapes in registry
            Console.WriteLine($"Registered 2D shapes: {ShapeRegistry.Shape2DMap.Count}");
            Console.WriteLine($"Registered 1D shapes: {ShapeRegistry.Shape1DMap.Count}");
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
        }
        
        if (is1D)
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
                // Copy legacy connection-related properties
                BeginConnectedTo = shape.BeginConnectedTo,
                EndConnectedTo = shape.EndConnectedTo,
                ConnectionPoints = shape.ConnectionPoints,
                FromPart = shape.FromPart,
                ToPart = shape.ToPart,
                PageName = shape.PageName,
            };
            foreach (var item in shape.Connections)
            {
                // Add each connection to the Shape1D object
                s1d.AddConnection(item);
            }

            
            // Debug to see if shape.Connections has any items
            //Console.WriteLine($"Debug: Shape {shape.ID} has {shape.Connections.Count} ConnectionPoint objects in VisioShape");
            
            // Debug connection information
            if (!string.IsNullOrEmpty(shape.BeginConnectedTo))
            {
                if (ShapeRegistry.Shape2DMap.ContainsKey(shape.BeginConnectedTo))
                {
                    Console.WriteLine($"Debug: Shape {shape.ID} has BeginConnectedTo = {shape.BeginConnectedTo} (will be populated later)");
                }
                else
                {
                    Console.WriteLine($"Debug: Shape {shape.ID} has BeginConnectedTo = {shape.BeginConnectedTo} but no matching 2D shape found!");
                }
            }
            
            if (!string.IsNullOrEmpty(shape.EndConnectedTo))
            {
                if (ShapeRegistry.Shape2DMap.ContainsKey(shape.EndConnectedTo))
                {
                    Console.WriteLine($"Debug: Shape {shape.ID} has EndConnectedTo = {shape.EndConnectedTo} (will be populated later)");
                }
                else
                {
                    Console.WriteLine($"Debug: Shape {shape.ID} has EndConnectedTo = {shape.EndConnectedTo} but no matching 2D shape found!");
                }
            }
            
            // For backward compatibility, still extract FromPart and ToPart from ConnectionPoints if available
            if (!string.IsNullOrEmpty(shape.ConnectionPoints))
            {
                var parts = shape.ConnectionPoints.Split('|');
                foreach (var part in parts)
                {
                    if (part.Contains("FromPart:"))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(part, @"FromPart:([^;]+)");
                        if (match.Success)
                        {
                            s1d.FromPart = match.Groups[1].Value;
                        }
                    }
                    if (part.Contains("ToPart:"))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(part, @"ToPart:([^;]+)");
                        if (match.Success)
                        {
                            s1d.ToPart = match.Groups[1].Value;
                        }
                    }
                }
            }
            
            // Debug ConnectionPoint objects

            
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
          // No need to extract connection info here anymore, as it's handled in ExtractShapesFromVsdx
        
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

    static NetworkModel CreateNetworkModel()
    {
        var networkModel = new NetworkModel();
        
        // Add 2D shapes as nodes (only if they have text)
        foreach (var shape2D in ShapeRegistry.Shape2DMap.Values)
        {
            if (!string.IsNullOrWhiteSpace(shape2D.Text))
            {
                var nodeShape = new NodeShape
                {
                    ID = shape2D.ID,
                    Text = shape2D.Text,
                    PinX = shape2D.PinX,
                    PinY = shape2D.PinY
                };
                
                networkModel.AddNode(nodeShape);
                Console.WriteLine($"Added node: {nodeShape.Text} (ID: {nodeShape.ID})");
            }
        }
        
        // Add 1D shapes as edges (only if they have text)
        foreach (var shape1D in ShapeRegistry.Shape1DMap.Values)
        {
            if (string.IsNullOrWhiteSpace(shape1D.Text))
                continue;

            if (shape1D.Connections == null || shape1D.Connections.Count == 0)
                continue;
            
            var edgeShape = new EdgeShape
            {
                ID = shape1D.ID,
                Text = shape1D.Text,
                FromNodeID = null,
                ToNodeID = null
            };
                

            // For 1D shapes with connections, try to map to connected 2D shapes
            var fromConnection = shape1D.Connections.FirstOrDefault(c => !string.IsNullOrEmpty(c.ToSheet));
            if (fromConnection != null)
            {
                edgeShape.FromNodeID = fromConnection.ToSheet;
            }
            

            var toConnection = shape1D.Connections.LastOrDefault(c => !string.IsNullOrEmpty(c.ToSheet));
            if (toConnection != null && toConnection != fromConnection)
            {
                edgeShape.ToNodeID = toConnection.ToSheet;
            }
            
                
                
            // Fall back to legacy connection properties if available
            if (string.IsNullOrEmpty(edgeShape.FromNodeID) && !string.IsNullOrEmpty(shape1D.BeginConnectedTo))
            {
                edgeShape.FromNodeID = shape1D.BeginConnectedTo;
            }
            
            if (string.IsNullOrEmpty(edgeShape.ToNodeID) && !string.IsNullOrEmpty(shape1D.EndConnectedTo))
            {
                edgeShape.ToNodeID = shape1D.EndConnectedTo;
            }
            
            networkModel.AddEdge(edgeShape);
            $"Added edge: {edgeShape.Text} (ID: {edgeShape.ID}, From: {edgeShape.FromNodeID ?? "none"}, To: {edgeShape.ToNodeID ?? "none"})".WriteSuccess(1);
            
        }
        
        return networkModel;
    }
}
