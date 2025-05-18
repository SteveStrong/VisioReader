using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public class ShapeKnowledgeModel
{
    public string Filename { get; set; } = string.Empty;
    public List<Shape2D>? Shape2DList { get; set; }
    public List<Shape1D>? Shape1DList { get; set; }

}

