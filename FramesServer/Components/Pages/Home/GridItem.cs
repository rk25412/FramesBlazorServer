using System.ComponentModel.DataAnnotations;

namespace FramesServer.Components.Pages.Home;

public class GridItem
{
    [Required] public string Name { get; set; } = string.Empty;

    [Required] [Range(1, 100)] public decimal? Rate { get; set; }

    [Required] public List<int> Lines { get; set; } = [];

    public int TotalLines => Lines.Sum();
    
    public decimal TotalAmt => Lines.Sum(x => x * Rate ?? 0);
    
    public string GridKey { get; set; } = Guid.NewGuid().ToString("N");
}
